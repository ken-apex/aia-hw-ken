#!/usr/bin/env python3
"""ดึงเอกสารทางการของหุ้นไทยจาก ก.ล.ต. iDisc ลงคลังหุ้น `C:\\ทุกบริษัท\\<TICKER>\\`

ได้งบการเงิน, MD&A และ 56-1 ตั้งชื่อ `YYYY-MM-DD_<ชนิด>_<งวด>` และแปลง PDF เป็น .txt ไว้คู่กัน
เขียนที่มาและวันที่ดึงของทุกไฟล์ลง `_reference/fetch-log.md` ไม่เขียนทับไฟล์ที่มีอยู่แล้ว

ตัวอย่าง:
    python output/fetch_stock.py SCC
    python output/fetch_stock.py QH --periods 4
"""
import argparse
import datetime as dt
import filecmp
import html
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://market.sec.or.th/public/idisc"
DOWNLOAD = BASE + "/Download?FILEID={fid}"
# iDisc ปฏิเสธ User-Agent ทั่วไป ต้องระบุตัวตน
UA = "aia-hw-ken/1.0 (AI Investor Alpha course homework)"


def get(url, timeout=90):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    time.sleep(1)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def listing(kind, sym):
    """คืน (url, ชื่อบริษัท, [(เซลล์ในแถว, FILEID), ...]) ของหน้ารายการ"""
    url = f"{BASE}/en/Viewmore/{kind}?searchSymbol={sym}"
    text = get(url).decode("utf-8", "ignore")
    title = re.search(r"<title>(.*?)</title>", text, re.S | re.I)
    company = re.sub(r"\s+", " ", title.group(1)).strip() if title else ""
    rows = []
    for row in re.findall(r"<tr.*?</tr>", text, re.S | re.I):
        fid = re.search(r"FILEID=([^\"'&<>\s]+)", row)
        if not fid:
            continue
        cells = [re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", c))).strip()
                 for c in re.findall(r"<td.*?</td>", row, re.S | re.I)]
        rows.append((cells, fid.group(1)))
    return url, company, rows


def be(year):
    return int(year) + 543


def fs_rows(rows, n):
    """เลือกงบรวมของ n งวดล่าสุด ถ้างวดไหนไม่มีงบรวมจึงใช้งบเฉพาะกิจการ"""
    by_period = {}
    for cells, fid in rows:
        year, kind, period = cells[0], cells[2], cells[3]
        label = f"FY-{be(year)}" if period == "Year" else f"{period}-{be(year)}"
        if label not in by_period or kind == "Consolidated":
            by_period[label] = fid
    return list(by_period.items())[:n]


def mda_label(heading):
    m = re.search(r"Quarter\s*(\d)\s*/\s*(\d{4})", heading) or \
        re.search(r"Quarter\s*(\d)\s+Ending\s+\d+\s+\w+\s+(\d{4})", heading)
    if m:
        quarter, year = m.group(1), m.group(2)
        return f"FY-{be(year)}" if quarter == "4" else f"Q{quarter}-{be(year)}"
    m = re.search(r"Year\w*\s+Ending\s+\d+\s+\w+\s+(\d{4})", heading)
    return f"FY-{be(m.group(1))}" if m else None


def iso(ddmmyyyy):
    return dt.datetime.strptime(ddmmyyyy, "%d/%m/%Y").date().isoformat()


def fs_suffix(name):
    low = name.lower()
    if "auditor" in low:
        return "_AUDITOR_REPORT"
    if "note" in low:
        return "_NOTES"
    if "financial_statement" in low or low.endswith((".xlsx", ".xls")):
        return ""
    return "_" + re.sub(r"\W+", "_", Path(name).stem)


def add_page_marks(txt):
    """ใส่เลขหน้าแบบเดียวกับไฟล์ของ CPALL เพื่อให้อ้างที่มาระบุหน้าได้"""
    pages = txt.read_text(encoding="utf-8", errors="replace").rstrip("\f\n").split("\f")
    body = "".join(f"\n\n===== หน้า {i} =====\n{p}" for i, p in enumerate(pages, 1))
    txt.write_text(body, encoding="utf-8")


class Library:
    def __init__(self, root, sym):
        self.top = Path(root) / sym
        (self.top / "_analysis").mkdir(parents=True, exist_ok=True)
        (self.top / "_reference").mkdir(parents=True, exist_ok=True)
        self.today = dt.date.today().isoformat()
        self.saved, self.skipped = [], []

    def put(self, src, name, url):
        """ย้ายไฟล์เข้าชั้นบนสุด ไม่เขียนทับของเดิม
        ถ้าชื่อซ้ำแต่เนื้อหาต่างกัน เก็บเป็น _v2 เพราะบริษัทอาจออกฉบับแก้ไขในวันเดียวกัน"""
        stem, suffix = Path(name).stem, Path(name).suffix
        dest, version = self.top / name, 1
        while dest.exists():
            if filecmp.cmp(src, dest, shallow=False):
                self.skipped.append(dest.name)
                return
            version += 1
            dest = self.top / f"{stem}_v{version}{suffix}"
        name = dest.name
        shutil.move(str(src), dest)
        self.saved.append((name, url))
        if dest.suffix.lower() == ".pdf":
            self.to_text(dest)

    def to_text(self, pdf):
        txt = pdf.with_suffix(".txt")
        if txt.exists():
            return
        r = subprocess.run(["pdftotext", "-layout", str(pdf), str(txt)], capture_output=True)
        if r.returncode != 0 or not txt.exists():
            print(f"  ! แปลงเป็นข้อความไม่ได้: {pdf.name}", file=sys.stderr)
            return
        add_page_marks(txt)
        if txt.stat().st_size < 500:
            print(f"  ! ข้อความสั้นผิดปกติ อาจเป็นภาพสแกน: {txt.name}", file=sys.stderr)

    def write_log(self):
        if not self.saved:
            return
        log = self.top / "_reference" / "fetch-log.md"
        lines = [] if log.exists() else ["# ที่มาของเอกสาร\n", "| ไฟล์ | URL | วันที่ดึง |", "|---|---|---|"]
        lines += [f"| {n} | {u} | {self.today} |" for n, u in self.saved]
        with log.open("a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")


def unpack(fid, tmp):
    url = DOWNLOAD.format(fid=fid)
    blob = get(url, timeout=300)
    folder = Path(tempfile.mkdtemp(dir=tmp))
    pack = folder / "package.zip"
    pack.write_bytes(blob)
    with zipfile.ZipFile(pack) as z:
        z.extractall(folder)
    pack.unlink()
    return url, [p for p in folder.rglob("*") if p.is_file()]


def fetch(sym, root, periods):
    lib = Library(root, sym)
    with tempfile.TemporaryDirectory() as tmp:
        _, company, rows = listing("fs-full", sym)
        print(f"{sym}: {company}")
        chosen = fs_rows(rows, periods)
        for label, fid in chosen:
            url, files = unpack(fid, tmp)
            m = re.search(r"FIN(\d{2})(\d{2})(\d{4})", fid)
            date = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
            for f in files:
                lib.put(f, f"{date}_FS_{label}{fs_suffix(f.name)}{f.suffix.lower()}", url)

        wanted = {label for label, _ in chosen}
        newest = max(dt.date.fromisoformat(n[:10]) for n, _ in lib.saved) if lib.saved else None
        _, _, rows = listing("fs-mda", sym)
        for cells, fid in rows:
            label, date = mda_label(cells[2]), iso(cells[0])
            # เก็บ MD&A ที่ออกหลังงบล่าสุดด้วย เพราะบางบริษัทออก MD&A ก่อนงบเต็ม
            if label in wanted or (label and newest and dt.date.fromisoformat(date) > newest):
                pdf = Path(tempfile.mkdtemp(dir=tmp)) / "mda.pdf"
                url = DOWNLOAD.format(fid=fid)
                pdf.write_bytes(get(url))
                lib.put(pdf, f"{date}_MDA_{label}.pdf", url)

        _, _, rows = listing("fs-r561", sym)
        if rows:
            cells, fid = rows[0]
            year, date = be(cells[1]), iso(cells[2])
            url, files = unpack(fid, tmp)
            pdfs = [f for f in files if f.suffix.lower() == ".pdf"]
            for f in pdfs:
                extra = "" if len(pdfs) == 1 else "_" + re.sub(r"\W+", "_", f.stem)
                lib.put(f, f"{date}_56-1_{year}{extra}.pdf", url)
    lib.write_log()
    print(f"  ได้ใหม่ {len(lib.saved)} ไฟล์ · มีอยู่แล้ว {len(lib.skipped)} ไฟล์")
    for name, _ in lib.saved:
        print(f"  + {name}")
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("symbol", help="ชื่อย่อหุ้น เช่น SCC")
    ap.add_argument("--periods", type=int, default=4, help="จำนวนงวดของงบล่าสุด (ค่าเริ่มต้น 4)")
    ap.add_argument("--root", default=r"C:\ทุกบริษัท", help="รากของคลังหุ้น")
    args = ap.parse_args()
    return fetch(args.symbol.upper(), args.root, args.periods)


if __name__ == "__main__":
    raise SystemExit(main())
