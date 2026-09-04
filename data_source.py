"""Lớp bọc mọi lời gọi thư viện vnstock.

Cô lập vnstock ở đây để nếu thư viện đổi API thì chỉ sửa 1 file.
Trả về DataFrame đã chuẩn hoá: cột 'Chỉ tiêu' + các cột kỳ (mới nhất bên trái).
"""
from __future__ import annotations

import collections
import contextlib
import io
import os
import re
import threading
import time
import warnings

import pandas as pd

import config

warnings.filterwarnings("ignore")


@contextlib.contextmanager
def _silence():
    """Chặn banner quảng cáo / log mà vnstock in ra stdout+stderr."""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        yield


# --------------------------------------------------------------------------- #
# Điều tiết tốc độ (rolling window) — bắt buộc vì gói Guest chặn 20 req/phút
# và tự kết thúc tiến trình khi vượt. Giữ MỌI cửa sổ 60s <= MAX_RPM.
# --------------------------------------------------------------------------- #
_REQ_TIMES = collections.deque()
_RL_LOCK = threading.Lock()


def _rate_limit():
    """Chặn nhịp trước mỗi lời gọi API để không vượt config.MAX_RPM/phút."""
    with _RL_LOCK:
        now = time.time()
        while _REQ_TIMES and now - _REQ_TIMES[0] >= 60:
            _REQ_TIMES.popleft()
        if len(_REQ_TIMES) >= config.MAX_RPM:
            wait = 60 - (now - _REQ_TIMES[0]) + 0.2
            if wait > 0:
                time.sleep(wait)
            now = time.time()
            while _REQ_TIMES and now - _REQ_TIMES[0] >= 60:
                _REQ_TIMES.popleft()
        _REQ_TIMES.append(time.time())


# Import vnstock trong vùng im lặng (banner in ngay khi import)
with _silence():
    from vnstock import Company, Finance, Listing, register_user

# Nạp API key nếu đặt qua biến môi trường VNSTOCK_API_KEY (để khả chuyển sang
# máy khác). Nếu không có, vnstock tự đọc key đã lưu ở ~/.vnstock/api_key.json.
_API_KEY = os.environ.get("VNSTOCK_API_KEY", "").strip()
if _API_KEY:
    with _silence():
        try:
            register_user(_API_KEY)
        except Exception:
            pass


class NoDataError(Exception):
    """Không có dữ liệu cho mã này."""


# --------------------------------------------------------------------------- #
# Danh sách công ty
# --------------------------------------------------------------------------- #
# Chuẩn hoá tên sàn giữa các nguồn (VCI dùng 'HSX' cho HOSE)
_EXCHANGE_ALIASES = {"HSX": "HOSE", "HOSE": "HOSE", "HNX": "HNX", "UPCOM": "UPCOM"}


def list_companies(exchanges=None) -> pd.DataFrame:
    """Trả DataFrame cổ phiếu trên các sàn yêu cầu: symbol, organ_name, exchange.

    Dùng nguồn Listing mặc định (schema sạch) và chuẩn hoá giá trị vì mỗi nguồn
    đặt tên khác nhau: type 'stock'/'STOCK', sàn 'HOSE'/'HSX'.
    """
    exchanges = [e.upper() for e in (exchanges or config.EXCHANGES)]
    _rate_limit()
    with _silence():
        df = Listing().symbols_by_exchange()

    if "organ_name" not in df.columns:
        df["organ_name"] = df.get("organ_short_name", "")
    df = df.copy()
    df["type"] = df["type"].astype(str).str.lower()
    df["exchange"] = df["exchange"].astype(str).str.upper().map(
        lambda x: _EXCHANGE_ALIASES.get(x, x)
    )
    df = df[(df["type"] == config.STOCK_TYPE) & (df["exchange"].isin(exchanges))]
    df = df[["symbol", "organ_name", "exchange"]].drop_duplicates("symbol")
    df = df.sort_values(["exchange", "symbol"]).reset_index(drop=True)
    return df


# --------------------------------------------------------------------------- #
# Tiện ích xử lý kỳ
# --------------------------------------------------------------------------- #
_PERIOD_RE = re.compile(r"^(\d{4})(?:[-_ ]?Q(\d))?$")


def _parse_period(label: str):
    """'2026-Q1' -> (2026, 1); '2025' -> (2025, 0). None nếu không phải cột kỳ."""
    m = _PERIOD_RE.match(str(label).strip())
    if not m:
        return None
    year = int(m.group(1))
    quarter = int(m.group(2)) if m.group(2) else 0
    return (year, quarter)


def _clean_statement(df: pd.DataFrame, n_keep: int) -> pd.DataFrame:
    """Chuẩn hoá 1 báo cáo: giữ cột chỉ tiêu + n kỳ mới nhất, bỏ dòng rỗng.

    - Bỏ các cột định danh nội bộ (item_en, item_id).
    - Sắp xếp cột kỳ giảm dần theo (năm, quý) rồi lấy n kỳ đầu.
    - Loại các dòng mà toàn bộ kỳ đều trống hoặc bằng 0 (không phục vụ phân tích).
    """
    if df is None or df.empty or config.ITEM_COL not in df.columns:
        raise NoDataError("báo cáo rỗng")

    period_cols = [c for c in df.columns if _parse_period(c) is not None]
    if not period_cols:
        raise NoDataError("không tìm thấy cột kỳ")

    period_cols.sort(key=_parse_period, reverse=True)
    keep = period_cols[:n_keep]

    out = df[[config.ITEM_COL] + keep].copy()
    out = out.rename(columns={config.ITEM_COL: "Chỉ tiêu"})
    for c in keep:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    # Bỏ dòng không có tên hoặc mọi kỳ đều trống/0
    out = out[out["Chỉ tiêu"].notna() & (out["Chỉ tiêu"].astype(str).str.strip() != "")]
    mask_all_empty = out[keep].fillna(0).eq(0).all(axis=1)
    out = out[~mask_all_empty].reset_index(drop=True)

    if out.empty:
        raise NoDataError("mọi chỉ tiêu đều rỗng")
    return out


# --------------------------------------------------------------------------- #
# Lấy toàn bộ dữ liệu cho 1 mã
# --------------------------------------------------------------------------- #
def _overview(symbol: str) -> dict:
    try:
        _rate_limit()
        with _silence():
            ov = Company(symbol=symbol, source=config.SOURCE).overview()
        row = ov.iloc[0].to_dict()
    except Exception:
        row = {}
    fields = [
        "organ_name", "sector", "current_price", "market_cap", "issue_share",
        "rating", "target_price", "is_bank", "foreigner_percentage",
        "listing_date", "com_group_code", "company_profile",
    ]
    return {k: row.get(k) for k in fields}


def get_financials(
    symbol: str,
    period: str = None,
    n_display: int = None,
    retries: int = 3,
    retry_wait: float = 2.0,
) -> dict:
    """Lấy 4 báo cáo + tổng quan cho 1 mã, đã chuẩn hoá.

    Trả về dict: {symbol, overview, balance, income, cashflow, periods, raw_full}
    - balance/income/cashflow: DataFrame n_display kỳ gần nhất.
    - raw_full: dict báo cáo giữ tối đa kỳ (dùng tính TTM).
    Ném NoDataError nếu không có dữ liệu cốt lõi.
    """
    period = period or config.PERIOD
    n_display = n_display or config.DISPLAY_PERIODS

    last_err = None
    for attempt in range(retries):
        try:
            with _silence():
                fin = Finance(symbol=symbol, source=config.SOURCE)
                _rate_limit()
                bs_raw = fin.balance_sheet(period=period, lang="vi")
                _rate_limit()
                is_raw = fin.income_statement(period=period, lang="vi")
                _rate_limit()
                cf_raw = fin.cash_flow(period=period, lang="vi")
            break
        except Exception as e:  # lỗi mạng / rate-limit -> thử lại
            last_err = e
            if attempt < retries - 1:
                time.sleep(retry_wait * (attempt + 1))
    else:
        raise NoDataError(f"lỗi tải BCTC: {last_err}")

    n_full = max(n_display, config.TTM_QUARTERS)
    balance = _clean_statement(bs_raw, n_display)
    income = _clean_statement(is_raw, n_display)
    cashflow = _clean_statement(cf_raw, n_display)

    # Bản đầy đủ (tối đa kỳ) cho tính TTM
    raw_full = {
        "balance": _clean_statement(bs_raw, n_full),
        "income": _clean_statement(is_raw, n_full),
    }

    periods = [c for c in balance.columns if c != "Chỉ tiêu"]

    return {
        "symbol": symbol,
        "overview": _overview(symbol),
        "balance": balance,
        "income": income,
        "cashflow": cashflow,
        "periods": periods,
        "raw_full": raw_full,
    }
