"""Tool thu thập & tổng hợp báo cáo tài chính HOSE + HNX.

Chạy toàn sàn (mặc định) hoặc theo danh sách mã. Mỗi doanh nghiệp -> 1 file Excel
5 sheet: Tổng quan, Bảng cân đối kế toán, Kết quả kinh doanh, Lưu chuyển tiền tệ,
Chỉ số tài chính.

Ví dụ:
    python main.py --symbols FPT                 # thử 1 mã
    python main.py --symbols FPT,VCB,SSI         # vài mã khác ngành
    python main.py --limit 10                    # 10 mã đầu (smoke test)
    python main.py                               # toàn bộ HOSE + HNX
    python main.py --force                        # tạo lại cả file đã có
"""
from __future__ import annotations

import argparse
import csv
import io
import logging
import os
import sys
import time
from datetime import datetime

# Ép UTF-8 cho stdout để in tiếng Việt trên Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import config
import data_source
import excel_builder

try:
    from tqdm import tqdm
except ImportError:  # tqdm là tuỳ chọn
    def tqdm(x, **k):
        return x


class _OwnLogsFilter(logging.Filter):
    """Chỉ cho log của tool ('bctc') hiện ra console; chặn log nội bộ vnstock."""

    def filter(self, record):
        return record.name.startswith("bctc")


def _setup_logging(out_dir: str):
    os.makedirs(out_dir, exist_ok=True)
    log_path = os.path.join(out_dir, "run.log")
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    file_h = logging.FileHandler(log_path, encoding="utf-8")
    file_h.setFormatter(fmt)
    console_h = logging.StreamHandler(sys.stderr)
    console_h.setFormatter(fmt)
    console_h.addFilter(_OwnLogsFilter())  # console gọn, file log đầy đủ

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.handlers = [file_h, console_h]
    return logging.getLogger("bctc")


def _load_done(progress_path: str) -> set:
    done = set()
    if os.path.exists(progress_path):
        with open(progress_path, encoding="utf-8") as f:
            for row in csv.reader(f):
                if len(row) >= 2 and row[1] == "OK":
                    done.add(row[0])
    return done


def _append_progress(progress_path: str, symbol: str, status: str, msg: str = ""):
    new = not os.path.exists(progress_path)
    with open(progress_path, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["symbol", "status", "time", "message"])
        w.writerow([symbol, status, datetime.now().isoformat(timespec="seconds"), msg])


def parse_args():
    p = argparse.ArgumentParser(description="Thu thập BCTC HOSE + HNX -> Excel/mã.")
    p.add_argument("--exchanges", default=",".join(config.EXCHANGES),
                   help="Sàn, phân tách bởi dấu phẩy (mặc định HOSE,HNX)")
    p.add_argument("--symbols", default="",
                   help="Danh sách mã tự chọn (bỏ qua --exchanges), vd: FPT,HPG")
    p.add_argument("--period", default=config.PERIOD, choices=["quarter", "year"])
    p.add_argument("--periods", type=int, default=config.DISPLAY_PERIODS,
                   help="Số kỳ hiển thị (mặc định 3)")
    p.add_argument("--output", default="output", help="Thư mục kết quả")
    p.add_argument("--sleep", type=float, default=0.0,
                   help="Nghỉ thêm giữa các mã (giây); mặc định 0 vì đã có điều tiết theo phút")
    p.add_argument("--retries", type=int, default=3, help="Số lần thử lại khi lỗi")
    p.add_argument("--limit", type=int, default=0, help="Chỉ xử lý N mã đầu (0 = tất cả)")
    p.add_argument("--force", action="store_true", help="Tạo lại cả file đã có")
    return p.parse_args()


def main():
    args = parse_args()
    out_dir = os.path.abspath(args.output)
    log = _setup_logging(out_dir)
    progress_path = os.path.join(out_dir, "progress.csv")

    # 1) Xác định danh sách mã
    if args.symbols.strip():
        symbols = [s.strip().upper() for s in args.symbols.split(",") if s.strip()]
        # Vẫn tra sàn để đặt file đúng thư mục
        try:
            listing = data_source.list_companies()
            ex_map = dict(zip(listing["symbol"], listing["exchange"]))
        except Exception:
            ex_map = {}
        companies = [(s, ex_map.get(s, "KHAC")) for s in symbols]
    else:
        exchanges = [e.strip().upper() for e in args.exchanges.split(",")]
        log.info("Lấy danh sách mã trên sàn: %s", exchanges)
        listing = data_source.list_companies(exchanges)
        companies = list(zip(listing["symbol"], listing["exchange"]))

    if args.limit > 0:
        companies = companies[: args.limit]

    done = set() if args.force else _load_done(progress_path)
    total = len(companies)
    log.info("Tổng số mã cần xử lý: %d (đã xong trước đó: %d)", total, len(done))

    ok = fail = skip = 0
    for symbol, exchange in tqdm(companies, desc="BCTC", unit="mã"):
        ex_dir = os.path.join(out_dir, exchange)
        os.makedirs(ex_dir, exist_ok=True)
        out_path = os.path.join(ex_dir, f"{symbol}.xlsx")

        if not args.force and (symbol in done or os.path.exists(out_path)):
            skip += 1
            continue

        try:
            fin = data_source.get_financials(
                symbol, period=args.period, n_display=args.periods,
                retries=args.retries,
            )
            excel_builder.build_company_workbook(fin, out_path)
            _append_progress(progress_path, symbol, "OK")
            ok += 1
        except data_source.NoDataError as e:
            fail += 1
            _append_progress(progress_path, symbol, "NO_DATA", str(e))
            log.warning("[%s] không có dữ liệu: %s", symbol, e)
        except Exception as e:  # noqa: BLE001
            fail += 1
            _append_progress(progress_path, symbol, "ERROR", str(e))
            log.exception("[%s] lỗi: %s", symbol, e)

        time.sleep(args.sleep)

    summary = (
        f"HOÀN TẤT: thành công={ok}, thất bại={fail}, bỏ qua(đã có)={skip}, "
        f"tổng={total}. Kết quả tại: {out_dir}"
    )
    log.info(summary)
    with open(os.path.join(out_dir, "run_summary.txt"), "w", encoding="utf-8") as f:
        f.write(datetime.now().isoformat() + "\n" + summary + "\n")
    print("\n" + summary)


if __name__ == "__main__":
    main()
