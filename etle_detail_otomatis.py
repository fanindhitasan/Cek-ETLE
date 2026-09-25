#!/usr/bin/env python3
"""Pemeriksa ETLE PMJ berurutan dari Excel menggunakan HTTP GET resmi halaman.

Tidak mencoba melewati CAPTCHA, autentikasi, atau pembatasan akses. Jika situs
berubah atau meminta verifikasi manusia, proses dihentikan agar dapat ditinjau.
"""

from __future__ import annotations

import argparse
import hashlib
import unicodedata
from zoneinfo import ZoneInfo
import html
import json
from html.parser import HTMLParser
import re
import sys
import time
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from openpyxl.styles import Font


BASE_URL = "https://etle-pmj.id/"
DEFAULT_DELAY_SECONDS = 1.0
DEFAULT_TIMEOUT_SECONDS = 45
MAX_ATTEMPTS = 3

INPUT_COLUMNS = ["Nomor Polisi", "Nomor Rangka", "Nomor Mesin"]
OUTPUT_COLUMNS = [
    "Status Proses",
    "Hasil ETLE",
    "Pesan Website",
    "HTTP Status",
    "Waktu Cek",
    "Jumlah Percobaan",
    "Waktu Pelanggaran",
    "Lokasi Pelanggaran",
    "Tipe Pelanggaran",
    "Status",
    "Detail",
    "Frekuensi Terkena ETLE",
]


def clean_identifier(value: Any) -> str:
    """Membersihkan identifier tanpa mengubah nol di awal."""
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return re.sub(r"\s+", "", text)


def strip_html(document: str) -> str:
    document = re.sub(r"<script\b[^>]*>.*?</script>", " ", document, flags=re.I | re.S)
    document = re.sub(r"<style\b[^>]*>.*?</style>", " ", document, flags=re.I | re.S)
    document = re.sub(r"<[^>]+>", " ", document)
    document = html.unescape(document)
    return re.sub(r"\s+", " ", document).strip()


# Adapter konservatif: hanya membaca tabel dengan header yang dikenali.
# Struktur respons positif belum diverifikasi dengan data kendaraan nyata.
FIELDS = ["Waktu Pelanggaran", "Lokasi Pelanggaran", "Tipe Pelanggaran", "Status", "Detail", "Referensi Resmi", "Detail URL"]
ALIASES = {
    "nomor referensi": "Referensi Resmi", "no referensi": "Referensi Resmi",
    "id pelanggaran": "Referensi Resmi", "nomor tilang": "Referensi Resmi",
    "waktu": "Waktu Pelanggaran", "waktu pelanggaran": "Waktu Pelanggaran",
    "tanggal": "Waktu Pelanggaran", "tanggal pelanggaran": "Waktu Pelanggaran",
    "tanggal kejadian": "Waktu Pelanggaran", "jam": "Waktu Pelanggaran",
    "lokasi": "Lokasi Pelanggaran", "lokasi pelanggaran": "Lokasi Pelanggaran",
    "lokasi kejadian": "Lokasi Pelanggaran", "tempat kejadian": "Lokasi Pelanggaran",
    "pelanggaran": "Tipe Pelanggaran", "jenis pelanggaran": "Tipe Pelanggaran",
    "tipe pelanggaran": "Tipe Pelanggaran", "status": "Status",
    "status pelanggaran": "Status", "status tilang": "Status",
    "detail": "Detail", "keterangan": "Detail", "detail pelanggaran": "Detail",
}


class ResultTables(HTMLParser):
    """Menangkap sel tabel dan, bila ada, link <a href> di baris yang sama
    (mis. tombol "Detail"/"Lihat" di halaman hasil), sejajar per baris."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables, self.table, self.row, self.cell = [], None, None, None
        self.links, self.table_links, self.row_link = [], None, None
        self.depth = 0

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.depth += 1
            if self.depth == 1:
                self.table = []
                self.table_links = []
        if self.depth != 1:
            return
        if tag == "tr":
            self.row = []
            self.row_link = None
        elif tag in ("th", "td"):
            self.cell = []
        elif tag == "br" and self.cell is not None:
            self.cell.append(" ")
        elif tag == "a" and self.row is not None and self.row_link is None:
            href = dict(attrs).get("href", "")
            if href and not href.strip().lower().startswith(("javascript:", "#")):
                self.row_link = href.strip()

    def handle_data(self, data):
        if self.depth == 1 and self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if self.depth == 1:
            if tag in ("td", "th") and self.cell is not None:
                if self.row is not None:
                    self.row.append(" ".join(" ".join(self.cell).split()))
                self.cell = None
            elif tag == "tr" and self.row is not None:
                self.table.append(self.row)
                self.table_links.append(self.row_link or "")
                self.row = None
                self.row_link = None
            elif tag == "table":
                self.tables.append(self.table)
                self.links.append(self.table_links)
                self.table = None
                self.table_links = None
        if tag == "table":
            self.depth = max(0, self.depth - 1)


def parse_etle_response(document: str) -> tuple[str, str, dict]:
    extra = {field: "" for field in FIELDS}
    extra["Frekuensi Terkena ETLE"] = ""
    plain = strip_html(document)
    alerts = re.findall(r"title\s*:\s*['\"]([^'\"]+)['\"]", document, flags=re.I)
    message = " | ".join(html.unescape(x).strip() for x in alerts)
    combined = f"{message} {plain}".lower()
    if any(x in combined for x in ("captcha", "verify you are human", "verifikasi manusia", "access denied", "too many requests")):
        return "VERIFIKASI DIPERLUKAN", message or "Verifikasi diperlukan", extra
    parser = ResultTables()
    parser.feed(document)
    records = []
    ambiguous = False
    for table, table_links in zip(parser.tables, parser.links):
        for header_index, headers in enumerate(table):
            mapped = [ALIASES.get(h.strip().lower().rstrip(":")) for h in headers]
            known = set(mapped) - {None}
            # Kata 'pelanggaran' pada teks petunjuk bukan bukti ada pelanggaran.
            if not {"Waktu Pelanggaran", "Lokasi Pelanggaran"}.issubset(known):
                continue
            body_links = table_links[header_index + 1:]
            for row_index, cells in enumerate(table[header_index + 1:]):
                if len(cells) != len(headers):
                    ambiguous = True
                    continue
                record = {field: "" for field in FIELDS}
                for field, value in zip(mapped, cells):
                    if field and value:
                        record[field] = (record[field] + " " + value).strip()
                if not record["Waktu Pelanggaran"] or not record["Lokasi Pelanggaran"]:
                    ambiguous = True
                    continue
                record["Detail"] = json.dumps(dict(zip(headers, cells)), ensure_ascii=False)
                href = body_links[row_index] if row_index < len(body_links) else ""
                # Link "Detail" resmi dari halaman hasil itu sendiri, bukan tebakan;
                # kosong kalau baris ini memang tidak menyertakan link (mis. sudah lunas).
                record["Detail URL"] = urllib.parse.urljoin(BASE_URL, href) if href else ""
                records.append(record)
            break
    if records:
        extra["records"] = records
        extra["Frekuensi Terkena ETLE"] = len(records)
        extra["partial"] = ambiguous
        return "ADA DATA TERBACA", "Hanya baris halaman ini; kelengkapan/paginasi belum diverifikasi" + ("; sebagian baris tidak dikenali" if ambiguous else ""), extra
    if any(x in combined for x in ("data tidak ditemukan", "tidak ada data", "no data available", "data not found")):
        extra["Frekuensi Terkena ETLE"] = 0
        return "TIDAK DITEMUKAN", message or "Data tidak ditemukan; bukan jaminan bebas pelanggaran", extra
    return "RESPONS TIDAK DIKENALI", message or "Detail belum dikenali. Perlu contoh HTML hasil pelanggaran untuk menyesuaikan parser.", extra


def atomic_write_excel(df: pd.DataFrame, output_path: Path) -> None:
    """Menyimpan checkpoint melalui file sementara agar output tidak setengah jadi."""
    temp_path = output_path.with_name(f".{output_path.stem}.tmp{output_path.suffix}")
    df.to_excel(temp_path, index=False, sheet_name="Hasil ETLE")
    temp_path.replace(output_path)


def request_one(
    session: requests.Session,
    nopol: str,
    norangka: str,
    nomesin: str,
    delay_seconds: float,
    timeout_seconds: int,
) -> tuple[str, str, int | None, int, dict]:
    params = {
        "aksi": "cek",
        "nopol": nopol,
        "norangka": norangka,
        "nomesin": nomesin,
    }
    last_error = ""
    last_status: int | None = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = session.get(BASE_URL, params=params, timeout=timeout_seconds)
            last_status = response.status_code

            if response.status_code == 429:
                return "DIHENTIKAN", "HTTP 429: terlalu banyak permintaan", 429, attempt, {}
            if response.status_code in {401, 403}:
                return "DIHENTIKAN", f"HTTP {response.status_code}: akses ditolak", response.status_code, attempt, {}
            response.raise_for_status()

            result, message, extra = parse_etle_response(response.text)
            return result, message, response.status_code, attempt, extra
        except requests.RequestException as exc:
            last_error = type(exc).__name__  # Jangan simpan URL berisi nomor rangka/mesin.
            if attempt == MAX_ATTEMPTS:
                return "GAGAL", last_error, last_status, attempt, {}
        finally:
            # Jeda berlaku setelah SETIAP request, termasuk request yang gagal.
            time.sleep(delay_seconds)

    return "GAGAL", last_error or "Kesalahan tidak diketahui", last_status, MAX_ATTEMPTS, {}


def validate_columns(df: pd.DataFrame) -> None:
    missing = [column for column in INPUT_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError("Kolom wajib tidak ditemukan: " + ", ".join(missing))



DETAIL_COLUMNS = ["Nomor Polisi", "ID Pelanggaran", "Dasar ID", "Kualitas ID",
                  "Urutan Tampilan", "Penanda Waktu", *FIELDS,
                  "Status Pembayaran", "Waktu Cek"]


def normalize(value):
    return " ".join(unicodedata.normalize("NFKC", str(value or "")).upper().split())


def parse_time(value):
    # Format tanggal numerik eksplisit; tanggal tanpa jam tidak diberi jam rekaan.
    value = " ".join(str(value).strip().replace("T", " ").split())
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
                "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
                "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return None


def event_id(plate, record):
    reference = normalize(record.get("Referensi Resmi"))
    raw_time = record.get("Waktu Pelanggaran", "")
    dt = parse_time(raw_time)
    canonical_time = dt.isoformat() if dt else normalize(raw_time)
    if reference:
        parts = [plate, "REFERENSI", reference]
        basis, quality = "Referensi resmi", "Referensi tersedia"
    else:
        parts = [plate, "DETAIL", canonical_time,
                 normalize(record.get("Lokasi Pelanggaran")),
                 normalize(record.get("Tipe Pelanggaran"))]
        basis = "Waktu + lokasi + tipe"
        quality = "Bersyarat: detail harus konsisten"
        if not all(parts[2:]) or not dt or not re.search(r"\d{1,2}:\d{2}", raw_time):
            quality = "PERLU VERIFIKASI: identitas/waktu kurang lengkap"
    digest = hashlib.sha256(json.dumps(parts, ensure_ascii=False).encode()).hexdigest()[:20].upper()
    return f"{plate}-{digest}", basis, quality


def detail_rows(plate, records, checked):
    rows = []
    dates = [parse_time(r.get("Waktu Pelanggaran", "")) for r in records]
    latest = max((d for d in dates if d), default=None)
    all_dates_known = all(d is not None for d in dates)
    ordered = sorted(zip(records, dates), key=lambda pair: pair[1] or datetime.min, reverse=True)
    for i, (record, dt) in enumerate(ordered, 1):
        key, basis, quality = event_id(plate, record)
        # Jangan menyimpulkan pembayaran dari status umum seperti "selesai".
        status = normalize(record.get("Status"))
        payment = {"SUDAH DIBAYAR": "Sudah Dibayar", "BELUM DIBAYAR": "Belum Dibayar",
                   "LUNAS": "Sudah Dibayar", "BELUM LUNAS": "Belum Dibayar"}.get(status, "Belum Diketahui")
        label = "Waktu tidak dikenali" if dt is None else (
            ("Terbaru dari hasil cek" if all_dates_known else "Terbaru dari waktu yang terbaca")
            if dt == latest else "Lebih lama")
        rows.append({"Nomor Polisi": plate, "ID Pelanggaran": key, "Dasar ID": basis,
                     "Kualitas ID": quality, "Urutan Tampilan": i, "Penanda Waktu": label,
                     **record, "Status Pembayaran": payment, "Waktu Cek": checked})
    # Jangan diam-diam menggabungkan dua kejadian dengan identitas sama.
    counts = {}
    for row in rows:
        counts[row["ID Pelanggaran"]] = counts.get(row["ID Pelanggaran"], 0) + 1
    for row in rows:
        if counts[row["ID Pelanggaran"]] > 1:
            row["Kualitas ID"] = "PERLU VERIFIKASI: ID berulang dalam respons"
    return rows


def save_result(summary, details, output):
    tmp = output.with_name("." + output.stem + ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Rekap Kendaraan")
        pd.DataFrame(details, columns=DETAIL_COLUMNS).to_excel(writer, index=False, sheet_name="Detail Pelanggaran")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                ws.column_dimensions[cells[0].column_letter].width = 25
            header = {cell.value: cell.column_letter for cell in ws[1]}
            link_cols = [header[name] for name in ("Detail URL", "Link Pelanggaran Terbaru") if name in header]
            for row in ws:
                for cell in row:
                    if isinstance(cell.value, str):
                        # Nilai eksternal ditulis sebagai teks, bukan formula Excel.
                        cell.data_type = "s"
            for col in link_cols:
                for cell in ws[col][1:]:
                    if cell.value:
                        # Hyperlink murni ke URL string, bukan formula HYPERLINK() Excel.
                        cell.hyperlink = cell.value
                        cell.font = Font(color="0563C1", underline="single")
    tmp.replace(output)


def run(args):
    if args.delay < 2:
        raise ValueError("Jeda minimal 2 detik")
    input_path = Path(args.input).expanduser().resolve()
    output = Path(args.output).expanduser().resolve() if args.output else Path(
        "hasil_etle_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".xlsx").resolve()
    if output == input_path or output.exists():
        raise ValueError("Gunakan nama output baru agar file yang ada tidak tertimpa.")
    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    df = pd.read_excel(input_path, sheet_name=sheet, dtype=str).fillna("")
    validate_columns(df)
    df = df[INPUT_COLUMNS].copy()
    for col in INPUT_COLUMNS:
        df[col] = df[col].map(clean_identifier)
    df["Nomor Polisi"] = df["Nomor Polisi"].str.replace(r"[^A-Z0-9]", "", regex=True)
    df = df.drop_duplicates().reset_index(drop=True)
    if df["Nomor Polisi"].duplicated().any():
        raise ValueError("Ada plat sama dengan nomor rangka/mesin berbeda. Perbaiki master terlebih dahulu.")
    summary = df[["Nomor Polisi"]].copy()
    for col in ("Status Proses", "Hasil ETLE", "Jumlah Pelanggaran Terbaca", "Jumlah ID Berbeda",
                "Waktu Pelanggaran Terbaru Terbaca", "Catatan", "HTTP Status", "Waktu Cek",
                "Link Pelanggaran Terbaru"):
        summary[col] = ""
    summary["Status Proses"] = "BELUM DIPERIKSA"
    details = []
    output.parent.mkdir(parents=True, exist_ok=True)
    save_result(summary, details, output)
    exit_code = 0
    with requests.Session() as session:
        session.headers.update({"User-Agent": "ETLEFleetChecker/2.0", "Accept-Language": "id-ID,id;q=0.9"})
        for index, row in df.iterrows():
            plate = row["Nomor Polisi"]
            if not all(row[c] for c in INPUT_COLUMNS):
                summary.at[index, "Status Proses"] = "DATA TIDAK LENGKAP"
                save_result(summary, details, output)
                continue
            print(f"[{index + 1}/{len(df)}] Memeriksa {plate}", flush=True)
            result, message, http, attempts, extra = request_one(
                session, plate, row["Nomor Rangka"], row["Nomor Mesin"], args.delay, args.timeout)
            checked = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%Y-%m-%d %H:%M:%S WIB")
            records = extra.get("records", [])
            rows = detail_rows(plate, records, checked)
            details.extend(rows)
            # detail_rows mengurutkan menurun berdasarkan waktu, jadi baris pertama = terbaru.
            latest_link = rows[0].get("Detail URL", "") if rows else ""
            parsed_times = [parse_time(r.get("Waktu Pelanggaran", "")) for r in records]
            newest = max((d for d in parsed_times if d), default=None)
            summary.at[index, "Status Proses"] = "SELESAI" if result in {"ADA DATA TERBACA", "TIDAK DITEMUKAN"} else result
            summary.at[index, "Hasil ETLE"] = result
            summary.at[index, "Jumlah Pelanggaran Terbaca"] = len(rows) if rows else (0 if result == "TIDAK DITEMUKAN" else "")
            summary.at[index, "Jumlah ID Berbeda"] = len({r["ID Pelanggaran"] for r in rows}) if rows else (0 if result == "TIDAK DITEMUKAN" else "")
            summary.at[index, "Waktu Pelanggaran Terbaru Terbaca"] = newest.isoformat(sep=" ") if newest else ""
            summary.at[index, "Catatan"] = message
            summary.at[index, "HTTP Status"] = str(http or "")
            summary.at[index, "Waktu Cek"] = checked
            summary.at[index, "Link Pelanggaran Terbaru"] = latest_link
            save_result(summary, details, output)
            if result in {"VERIFIKASI DIPERLUKAN", "DIHENTIKAN"}:
                exit_code = 3
                break
    print(f"Hasil: {output}")
    return exit_code


def build_parser():
    parser = argparse.ArgumentParser(description="ETLE: satu baris per pelanggaran, tanpa membaca riwayat lama.",
        epilog="Instal: py -m pip install pandas openpyxl requests tzdata. Input: Nomor Polisi, Nomor Rangka, Nomor Mesin. Parser positif belum diverifikasi dengan respons nyata.")
    parser.add_argument("--input", required=True, help="Excel master kendaraan")
    parser.add_argument("--sheet", default="0", help="Nama atau indeks sheet")
    parser.add_argument("--output", help="Nama file baru; default memakai waktu running")
    parser.add_argument("--delay", type=float, default=2.0)
    parser.add_argument("--timeout", type=int, default=45)
    return parser


if __name__ == "__main__":
    try:
        raise SystemExit(run(build_parser().parse_args()))
    except KeyboardInterrupt:
        print("Dihentikan. Hasil kendaraan yang sudah selesai tersimpan.", file=sys.stderr)
        raise SystemExit(130)
