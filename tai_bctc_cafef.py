"""
Script tải Báo Cáo Tài Chính Q1/2026 từ CafeF.vn
================================================
Tác giả: Claude AI
Mục tiêu: Tải toàn bộ BCTC Q1/2026 (Cân đối kế toán, KQKD, Lưu chuyển tiền tệ)
           cho tất cả công ty niêm yết trên HOSE và HNX từ cafef.vn
Định dạng xuất: File Excel (.xlsx), mỗi công ty 1 file, lưu vào thư mục BCTC_Q1_2026/

Yêu cầu:
    pip install requests beautifulsoup4 openpyxl

Cách chạy:
    python tai_bctc_cafef.py

    Hoặc chỉ tải một sàn:
    python tai_bctc_cafef.py --exchange HOSE
    python tai_bctc_cafef.py --exchange HNX

Lưu ý: Chạy trực tiếp trên máy tính của bạn (không qua proxy/VPN).
        Thời gian ước tính: 3-5 giờ cho toàn bộ ~700 công ty.
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import time
import json
import argparse
import logging
from datetime import datetime
from pathlib import Path

# ─── Cấu hình ───────────────────────────────────────────────────────────────
YEAR = 2026
QUARTER = 1
OUTPUT_DIR = "BCTC_Q1_2026"
LOG_FILE = "tai_bctc.log"
DELAY_BETWEEN_REQUESTS = 1.5   # giây, tránh bị chặn
DELAY_BETWEEN_COMPANIES = 2.0  # giây giữa các công ty
MAX_RETRIES = 3

REPORT_TYPES = {
    "BSHEET":   "Cân đối kế toán",
    "INCSTA":   "Kết quả kinh doanh",
    "CASHFLOW": "Lưu chuyển tiền tệ",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://cafef.vn/",
}

# ─── Logging ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ─── Lấy danh sách công ty ───────────────────────────────────────────────────
def fetch_company_list_from_vndirect(exchange: str) -> list[dict]:
    """Lấy danh sách cổ phiếu từ VNDirect API (phương án 1)."""
    url = (
        f"https://api.vndirect.com.vn/v4/stocks"
        f"?q=type:STOCK~status:LISTED~comGroupCode:{exchange}"
        f"&size=2000&sort=code"
    )
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        data = r.json()
        return [
            {"ticker": item["code"], "name": item.get("name", ""), "exchange": exchange}
            for item in data.get("data", [])
        ]
    except Exception as e:
        log.warning(f"VNDirect API lỗi cho {exchange}: {e}")
        return []


def fetch_company_list_from_ssi(exchange: str) -> list[dict]:
    """Lấy danh sách cổ phiếu từ SSI API (phương án 2)."""
    url = f"https://iboard-query.ssi.com.vn/v2/stock/exchange/{exchange}?size=2000"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        data = r.json()
        items = data.get("data", data.get("items", []))
        return [
            {"ticker": item.get("sym", item.get("symbol", "")),
             "name": item.get("stockName", item.get("name", "")),
             "exchange": exchange}
            for item in items
            if item.get("sym", item.get("symbol", ""))
        ]
    except Exception as e:
        log.warning(f"SSI API lỗi cho {exchange}: {e}")
        return []


def fetch_company_list_from_cafef_page(exchange: str) -> list[dict]:
    """Lấy danh sách từ trang cafef (phương án 3 – parse HTML)."""
    exchange_lower = exchange.lower()
    results = []
    page = 1
    while True:
        url = f"https://cafef.vn/du-lieu/{exchange_lower}.chn?page={page}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            r.raise_for_status()
            soup = BeautifulSoup(r.text, "html.parser")
            # tìm các link dạng /du-lieu/hose/XXX-...
            links = soup.find_all("a", href=True)
            found = 0
            for link in links:
                href = link["href"]
                if f"/du-lieu/{exchange_lower}/" in href:
                    parts = href.split("/")
                    if len(parts) >= 4:
                        slug = parts[-1].replace(".chn", "")
                        ticker = slug.split("-")[0].upper()
                        if 1 <= len(ticker) <= 5 and ticker.isalpha():
                            name = link.get_text(strip=True)
                            if {"ticker": ticker, "exchange": exchange} not in [
                                {"ticker": x["ticker"], "exchange": x["exchange"]} for x in results
                            ]:
                                results.append({"ticker": ticker, "name": name, "exchange": exchange})
                                found += 1
            if found == 0:
                break
            page += 1
            time.sleep(1)
        except Exception as e:
            log.warning(f"Cafef page lỗi trang {page}: {e}")
            break
    return results


def get_company_list(exchange: str) -> list[dict]:
    """Lấy danh sách công ty, thử nhiều nguồn."""
    log.info(f"Đang lấy danh sách công ty {exchange}...")

    # Thử VNDirect trước
    companies = fetch_company_list_from_vndirect(exchange)
    if companies:
        log.info(f"  VNDirect: {len(companies)} công ty {exchange}")
        return companies

    # Thử SSI
    companies = fetch_company_list_from_ssi(exchange)
    if companies:
        log.info(f"  SSI: {len(companies)} công ty {exchange}")
        return companies

    # Parse từ cafef
    companies = fetch_company_list_from_cafef_page(exchange)
    if companies:
        log.info(f"  Cafef HTML: {len(companies)} công ty {exchange}")
        return companies

    # Fallback: danh sách cứng từ knowledge base (các mã phổ biến)
    log.warning(f"  Không lấy được list từ API, dùng danh sách cứng cho {exchange}")
    return get_hardcoded_list(exchange)


def get_hardcoded_list(exchange: str) -> list[dict]:
    """Danh sách fallback các mã phổ biến."""
    hose = [
        "ACB","AGR","ACV","ANV","APH","BCG","BCM","BID","BMI","BMP","BVH","BWE",
        "CAV","CII","CMG","CTD","CTG","CTP","CTS","DCM","DHC","DIG","DPM","DXG",
        "EIB","EVF","FPT","FRT","GAS","GEX","GMD","GVR","HAG","HAH","HCM","HDB",
        "HDC","HDG","HHV","HMC","HNG","HPG","HSG","HT1","HUT","IDC","IJC","IMP",
        "KBC","KDC","KDH","KHG","KOS","LCG","LDG","LGC","LHG","LPB","MBB","MCH",
        "MHC","MSB","MSN","MWG","NAB","NCT","NDN","NKG","NLG","NVL","OCB","OIL",
        "PAC","PAN","PDR","PHP","PNJ","POW","PPF","PPC","PTB","PVD","PVS","PVT",
        "QNS","REE","SAB","SAM","SAV","SBT","SCS","SGN","SHB","SJS","SKG","SRC",
        "SSB","SSI","STB","SVC","TCB","TCH","TDM","TLG","TNH","TPB","TV2","TVS",
        "VCB","VCG","VCI","VGC","VHC","VHM","VIB","VIC","VIX","VJC","VLB","VND",
        "VNM","VOS","VPB","VRE","VSC","VTJ","YEG",
    ]
    hnx = [
        "ACE","APC","APG","BAB","BBS","BCC","BIC","BKC","BLF","BPC","BST","BTP",
        "BTT","BVS","C32","C47","CAN","CAP","CCI","CDN","CEO","CLG","CMI","CMS",
        "CNT","CSC","CST","CTB","CTF","CTI","CTN","CTS","DBC","DC4","DCL","DCS",
        "DDG","DHT","DL1","DMC","DNM","DNP","DNS","DPC","DST","DTD","DTL","DTP",
        "DXP","EBS","FCN","FID","GDT","GHC","GKM","GLT","GLX","GPH","HAD","HAP",
        "HAS","HAT","HBC","HBH","HC3","HCC","HCI","HCT","HEJ","HGM","HID","HII",
        "HLC","HLD","HLG","HMH","HNA","HNM","HOM","HPC","HPN","HPS","HPX","HQC",
        "HRC","HSA","HTC","HTG","HTI","HTL","HTN","HTP","HUB","HUT","HVA","HVN",
        "ICG","ICI","IDJ","ILS","IME","INC","INN","ITC","IVS","JVC","KBC","KBS",
        "KHB","KHL","KHG","KMR","KPF","KSD","KSF","L14","L35","L43","L61","LAS",
        "LAW","LBE","LCF","LCS","LDP","LHC","LIG","LIX","LLC","LM7","LM8","LTC",
        "LUT","MCO","MDC","MEL","MIC","MIM","MNC","MNS","MSP","MVN","NAG","NAP",
        "NAS","NBC","NBB","NBT","NCG","NDF","NDN","NGK","NHP","NHT","NLC","NMT",
        "NQT","NRC","NST","NTB","NTP","NTS","NVB","NVT","OGC","PBP","PCG","PDB",
        "PFL","PGC","PGI","PGL","PGS","PHC","PHN","PHP","PHS","PIV","PIX","PLO",
        "PMC","PMP","PPI","PPP","PPS","PPY","PRE","PSC","PSG","PSH","PSI","PSW",
        "PTC","PTH","PTK","PTO","PTP","PTS","PTT","PVB","PVC","PVE","PVG","PVL",
        "PVM","PVR","PVV","PXA","PXI","PXL","PXM","QHW","QND","QST","QTC","RCL",
        "RHC","S55","S96","S99","SBS","SCI","SCL","SCR","SEB","SGC","SGD","SGH",
        "SHA","SHB","SHC","SHI","SHL","SHN","SHX","SIE","SII","SIP","SJ1","SJC",
        "SJD","SJE","SJF","SKS","SLS","SMA","SMB","SMN","SNC","SOB","SPC","SPH",
        "SPI","SRA","SRB","SRC","SRF","SRI","SRT","SRS","SRU","STA","STB","STC",
        "TDN","TDT","TED","TET","TGP","THB","THD","THS","TIG","TIN","TIX","TKG",
        "TLD","TLH","TLT","TMS","TMT","TNA","TNB","TNG","TNP","TNT","TOC","TPH",
        "TPQ","TPS","TQL","TRC","TRS","TRT","TSB","TTC","TTH","TTI","TTL","TTS",
        "TTZ","TUG","TVC","TVD","TVE","TVG","TVH","TVN","TVP","TVQ","TVS","TVT",
        "TVU","TVW","TVX","TVY","TXM","UDC","UDJ","UDS","UIV","UNI","UPC","V12",
        "V21","VBC","VBH","VC1","VC2","VC3","VC5","VC6","VC7","VC9","VCC","VCF",
        "VCG","VCH","VCM","VCR","VCS","VCT","VDB","VDL","VDP","VDS","VEC","VEF",
        "VFR","VGP","VGS","VGT","VHE","VHL","VHT","VIE","VIG","VIN","VIR","VIS",
        "VIT","VIW","VKC","VLA","VMC","VMD","VNA","VNC","VND","VNE","VNF","VNG",
        "VNH","VNI","VNL","VNN","VNP","VNR","VNS","VNT","VNW","VOC","VPC","VPG",
        "VQC","VRG","VSA","VSC","VSG","VSH","VSI","VSM","VSN","VSP","VSR","VST",
        "VTA","VTB","VTC","VTD","VTG","VTH","VTI","VTK","VTL","VTM","VTP","VTQ",
        "VTS","VTT","VTV","VTX","VTV","VTZ","VUA","VUG","VXB",
    ]
    if exchange.upper() == "HOSE":
        return [{"ticker": t, "name": "", "exchange": "HOSE"} for t in hose]
    else:
        return [{"ticker": t, "name": "", "exchange": "HNX"} for t in hnx]


# ─── Lấy dữ liệu tài chính ──────────────────────────────────────────────────
def fetch_financial_data(ticker: str, report_type: str, session: requests.Session) -> list[dict]:
    """Lấy một loại BCTC cho một mã cổ phiếu từ cafef."""
    slug_map = {
        "BSHEET":   "can-doi-ke-toan",
        "INCSTA":   "ket-qua-kinh-doanh",
        "CASHFLOW": "luu-chuyen-tien-te",
    }
    slug = slug_map.get(report_type, "bao-cao-tai-chinh")
    url = (
        f"https://cafef.vn/du-lieu/bao-cao-tai-chinh/"
        f"{ticker.lower()}/{report_type}/{YEAR}/{QUARTER}/0/0/{slug}.chn"
    )

    for attempt in range(MAX_RETRIES):
        try:
            r = session.get(url, timeout=20)
            if r.status_code == 404:
                log.debug(f"  {ticker}/{report_type}: 404 - không có dữ liệu")
                return []
            r.raise_for_status()
            return parse_financial_table(r.text, ticker, report_type)
        except requests.exceptions.Timeout:
            log.warning(f"  {ticker}/{report_type}: Timeout lần {attempt+1}")
            time.sleep(2 ** attempt)
        except requests.exceptions.RequestException as e:
            log.warning(f"  {ticker}/{report_type}: Lỗi lần {attempt+1}: {e}")
            time.sleep(2 ** attempt)
    return []


def parse_financial_table(html: str, ticker: str, report_type: str) -> list[dict]:
    """Parse bảng dữ liệu tài chính từ HTML cafef."""
    soup = BeautifulSoup(html, "html.parser")
    rows = []

    # Tìm tất cả các row chứa dữ liệu tài chính
    # Cafef render dữ liệu trong các list-item và table
    # Dữ liệu thường xuất hiện trong dạng: "Tên chỉ tiêu VALUE1 VALUE2 VALUE3 VALUE4"

    # Thử tìm bảng dữ liệu chính
    tables = soup.find_all("table")
    for table in tables:
        table_rows = table.find_all("tr")
        for tr in table_rows:
            cells = tr.find_all(["td", "th"])
            if cells:
                row_data = [c.get_text(strip=True) for c in cells]
                if len(row_data) >= 2 and any(row_data):
                    rows.append(row_data)

    if not rows:
        # Thử parse từ list items (cách cafef hiển thị BCTC)
        items = soup.find_all("li")
        for item in items:
            text = item.get_text(strip=True)
            if text and len(text) > 5:
                # Tách tên chỉ tiêu và giá trị
                parts = text.split()
                if len(parts) >= 2:
                    rows.append([text])

    # Chuyển sang list dict
    result = []
    for i, row in enumerate(rows):
        if row and any(r.strip() for r in row):
            result.append({
                "stt": i + 1,
                "chi_tieu": row[0] if row else "",
                "q2_2025": row[1] if len(row) > 1 else "",
                "q3_2025": row[2] if len(row) > 2 else "",
                "q4_2025": row[3] if len(row) > 3 else "",
                "q1_2026": row[4] if len(row) > 4 else "",
                "tang_truong": row[5] if len(row) > 5 else "",
            })

    return result


# ─── Xuất Excel ─────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
SUBHEADER_FILL = PatternFill("solid", fgColor="2E75B6")
SUBHEADER_FONT = Font(bold=True, color="FFFFFF", size=9)
DATA_FONT = Font(size=9)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def write_excel(ticker: str, company_name: str, exchange: str,
                all_data: dict[str, list[dict]], output_path: Path):
    """Ghi dữ liệu BCTC vào file Excel với định dạng đẹp."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # xóa sheet mặc định

    report_written = False

    for report_type, display_name in REPORT_TYPES.items():
        data = all_data.get(report_type, [])
        if not data:
            continue

        ws = wb.create_sheet(title=display_name[:31])
        report_written = True

        # Tiêu đề chính
        ws.merge_cells("A1:G1")
        ws["A1"] = f"{ticker} – {company_name} ({exchange})"
        ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
        ws["A1"].alignment = Alignment(horizontal="center")

        ws.merge_cells("A2:G2")
        ws["A2"] = f"BÁO CÁO TÀI CHÍNH – {display_name.upper()} – QUÝ {QUARTER}/{YEAR}"
        ws["A2"].font = Font(bold=True, size=11)
        ws["A2"].alignment = Alignment(horizontal="center")

        ws.merge_cells("A3:G3")
        ws["A3"] = f"Nguồn: cafef.vn | Tải lúc: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Đơn vị: triệu đồng"
        ws["A3"].font = Font(italic=True, size=9, color="666666")
        ws["A3"].alignment = Alignment(horizontal="center")

        # Header cột
        headers = ["STT", "CHỈ TIÊU", "Q2/2025", "Q3/2025", "Q4/2025", "Q1/2026", "TĂNG TRƯỞNG"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = THIN_BORDER

        # Dữ liệu
        for row_idx, row in enumerate(data, 6):
            ws.cell(row=row_idx, column=1, value=row.get("stt", "")).border = THIN_BORDER

            chi_tieu_cell = ws.cell(row=row_idx, column=2, value=row.get("chi_tieu", ""))
            chi_tieu_cell.font = DATA_FONT
            chi_tieu_cell.border = THIN_BORDER

            for col, key in enumerate(["q2_2025", "q3_2025", "q4_2025", "q1_2026", "tang_truong"], 3):
                val = row.get(key, "")
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.font = DATA_FONT
                cell.alignment = Alignment(horizontal="right")
                cell.border = THIN_BORDER

        # Chiều rộng cột
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 45
        for col_letter in ["C", "D", "E", "F", "G"]:
            ws.column_dimensions[col_letter].width = 18

        # Freeze header
        ws.freeze_panes = "A6"

    if report_written:
        wb.save(output_path)
        return True
    return False


# ─── Tiến trình chính ────────────────────────────────────────────────────────
def process_exchange(exchange: str, output_dir: Path, session: requests.Session,
                     success_log: list, fail_log: list):
    """Xử lý toàn bộ một sàn."""
    companies = get_company_list(exchange)
    log.info(f"Sàn {exchange}: {len(companies)} công ty")

    exch_dir = output_dir / exchange
    exch_dir.mkdir(parents=True, exist_ok=True)

    for i, company in enumerate(companies, 1):
        ticker = company["ticker"]
        name = company.get("name", "")

        log.info(f"[{i}/{len(companies)}] {exchange} – {ticker} {name}")

        # Bỏ qua nếu đã tải
        out_file = exch_dir / f"{ticker}_BCTC_Q{QUARTER}_{YEAR}.xlsx"
        if out_file.exists():
            log.info(f"  → Đã tồn tại, bỏ qua")
            success_log.append({"ticker": ticker, "exchange": exchange, "status": "skip"})
            continue

        # Lấy dữ liệu 3 loại báo cáo
        all_data = {}
        has_data = False

        for rtype in REPORT_TYPES:
            data = fetch_financial_data(ticker, rtype, session)
            all_data[rtype] = data
            if data:
                has_data = True
            time.sleep(DELAY_BETWEEN_REQUESTS)

        if not has_data:
            log.warning(f"  → Không có dữ liệu Q{QUARTER}/{YEAR} cho {ticker}")
            fail_log.append({"ticker": ticker, "exchange": exchange, "reason": "no_data"})
        else:
            ok = write_excel(ticker, name, exchange, all_data, out_file)
            if ok:
                log.info(f"  → Đã lưu: {out_file.name}")
                success_log.append({"ticker": ticker, "exchange": exchange, "status": "ok"})
            else:
                fail_log.append({"ticker": ticker, "exchange": exchange, "reason": "empty_excel"})

        time.sleep(DELAY_BETWEEN_COMPANIES)


def main():
    parser = argparse.ArgumentParser(description="Tải BCTC Q1/2026 từ cafef.vn")
    parser.add_argument("--exchange", choices=["HOSE", "HNX", "ALL"], default="ALL",
                        help="Sàn cần tải (mặc định: ALL)")
    args = parser.parse_args()

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    log.info("=" * 60)
    log.info(f"BẮT ĐẦU TẢI BCTC Q{QUARTER}/{YEAR} TỪ CAFEF.VN")
    log.info(f"Thời gian bắt đầu: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    log.info(f"Thư mục lưu: {output_dir.resolve()}")
    log.info("=" * 60)

    session = requests.Session()
    session.headers.update(HEADERS)
    # Warm-up request để lấy cookies
    try:
        session.get("https://cafef.vn", timeout=10)
        time.sleep(1)
    except Exception:
        pass

    success_log = []
    fail_log = []

    exchanges = ["HOSE", "HNX"] if args.exchange == "ALL" else [args.exchange]
    for exch in exchanges:
        process_exchange(exch, output_dir, session, success_log, fail_log)

    # Lưu báo cáo kết quả
    summary_file = output_dir / "ket_qua_tai.json"
    with open(summary_file, "w", encoding="utf-8") as f:
        json.dump({
            "thoi_gian": datetime.now().isoformat(),
            "tong_thanh_cong": len([x for x in success_log if x["status"] == "ok"]),
            "tong_bo_qua": len([x for x in success_log if x["status"] == "skip"]),
            "tong_loi": len(fail_log),
            "thanh_cong": success_log,
            "loi": fail_log,
        }, f, ensure_ascii=False, indent=2)

    log.info("=" * 60)
    log.info(f"HOÀN THÀNH!")
    log.info(f"  ✓ Thành công: {len([x for x in success_log if x['status'] == 'ok'])} công ty")
    log.info(f"  ↷ Bỏ qua (đã có): {len([x for x in success_log if x['status'] == 'skip'])} công ty")
    log.info(f"  ✗ Lỗi/không có dữ liệu: {len(fail_log)} công ty")
    log.info(f"  → File lưu tại: {output_dir.resolve()}")
    log.info(f"  → Báo cáo chi tiết: {summary_file}")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
