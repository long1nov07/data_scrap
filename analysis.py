"""Tính chỉ số tài chính & khối chỉ tiêu then chốt từ báo cáo đã chuẩn hoá.

Tự tính thay vì dùng ratio() của vnstock vì bản community trả về kỳ cũ (2018).
Mọi phép chia đều an toàn với giá trị trống -> trả về None (hiển thị trống).
"""
from __future__ import annotations

import pandas as pd

import config


def _lookup(df: pd.DataFrame, name, period: str):
    """Lấy giá trị 1 chỉ tiêu tại 1 kỳ; None nếu không có.

    `name` là chuỗi hoặc danh sách tên thay thế (thử lần lượt, lấy tên đầu tiên
    có mặt trong báo cáo) -> hỗ trợ cả DN phi tài chính lẫn ngân hàng/CK.
    """
    if df is None or df.empty or period not in df.columns:
        return None
    names = [name] if isinstance(name, str) else list(name)
    for nm in names:
        hit = df[df["Chỉ tiêu"] == nm]
        if not hit.empty:
            val = hit.iloc[0][period]
            if pd.isna(val):
                return None
            return float(val)
    return None


def _div(a, b):
    """Chia an toàn -> None nếu thiếu dữ liệu hoặc mẫu số ~ 0."""
    if a is None or b is None or b == 0:
        return None
    return a / b


def _growth(cur, prev):
    """Tăng trưởng % (theo tỷ lệ, 0.12 = 12%). Dựa trên |mẫu số|."""
    if cur is None or prev is None or prev == 0:
        return None
    return (cur - prev) / abs(prev)


def key_indicators(fin: dict) -> pd.DataFrame:
    """Bảng chỉ tiêu then chốt qua các kỳ + cột %QoQ (kỳ mới nhất so kỳ liền trước)."""
    periods = fin["periods"]
    src_map = {"BS": (fin["balance"], config.BS),
               "IS": (fin["income"], config.IS),
               "CF": (fin["cashflow"], config.CF)}

    rows = []
    for label, src, key in config.KEY_INDICATORS:
        df, names = src_map[src]
        item_name = names[key]
        vals = [_lookup(df, item_name, p) for p in periods]
        qoq = _growth(vals[0], vals[1]) if len(vals) >= 2 else None
        rows.append([label] + vals + [qoq])

    cols = ["Chỉ tiêu"] + periods + ["%QoQ"]
    return pd.DataFrame(rows, columns=cols)


def financial_ratios(fin: dict) -> pd.DataFrame:
    """Bảng chỉ số tài chính tự tính, mỗi cột là 1 kỳ.

    Nhóm: Sinh lời, Thanh khoản, Đòn bẩy, Hiệu quả, Tăng trưởng.
    Ngân hàng/CK thiếu chỉ tiêu -> ô để trống.
    """
    periods = fin["periods"]
    bs, is_, cf = fin["balance"], fin["income"], fin["cashflow"]

    def b(k, p):  # noqa: E741 - helper ngắn cho dễ đọc
        return _lookup(bs, config.BS[k], p)

    def i(k, p):
        return _lookup(is_, config.IS[k], p)

    rows = []

    def add(label, fn, pct=False, suffix=""):
        vals = []
        for p in periods:
            v = fn(p)
            vals.append(v)
        rows.append({"label": label, "vals": vals, "pct": pct, "suffix": suffix})

    # --- Sinh lời ---
    add("Biên LN gộp", lambda p: _div(i("loi_nhuan_gop", p), i("doanh_thu_thuan", p)), pct=True)
    add("Biên LN ròng", lambda p: _div(i("ln_sau_thue", p), i("doanh_thu_thuan", p)), pct=True)
    add("ROA (quý)", lambda p: _div(i("ln_sau_thue", p), b("tong_tai_san", p)), pct=True)
    add("ROE (quý)", lambda p: _div(i("ln_sau_thue", p), b("von_chu_so_huu", p)), pct=True)
    # --- Thanh khoản ---
    add("Hệ số thanh toán hiện hành",
        lambda p: _div(b("tai_san_ngan_han", p), b("no_ngan_han", p)), suffix=" lần")
    add("Hệ số thanh toán nhanh",
        lambda p: _div(
            (b("tai_san_ngan_han", p) or 0) - (b("hang_ton_kho", p) or 0)
            if b("tai_san_ngan_han", p) is not None else None,
            b("no_ngan_han", p)), suffix=" lần")
    # --- Đòn bẩy ---
    add("Nợ phải trả / VCSH (D/E)",
        lambda p: _div(b("no_phai_tra", p), b("von_chu_so_huu", p)), suffix=" lần")
    add("Nợ vay / VCSH",
        lambda p: _div((b("vay_ngan_han", p) or 0) + (b("vay_dai_han", p) or 0)
                       if (b("vay_ngan_han", p) is not None or b("vay_dai_han", p) is not None)
                       else None,
                       b("von_chu_so_huu", p)), suffix=" lần")
    add("Nợ phải trả / Tổng tài sản",
        lambda p: _div(b("no_phai_tra", p), b("tong_tai_san", p)), pct=True)
    # --- Hiệu quả (theo quý) ---
    add("Vòng quay hàng tồn kho (quý)",
        lambda p: _div(i("gia_von", p), b("hang_ton_kho", p)), suffix=" lần")
    add("Vòng quay tổng tài sản (quý)",
        lambda p: _div(i("doanh_thu_thuan", p), b("tong_tai_san", p)), suffix=" lần")

    # --- Tăng trưởng (so kỳ liền trước) ---
    def growth_row(label, src, key):
        df = fin["income"] if src == "IS" else fin["balance"]
        name = (config.IS if src == "IS" else config.BS)[key]
        vals = []
        for idx, p in enumerate(periods):
            cur = _lookup(df, name, p)
            prev = _lookup(df, name, periods[idx + 1]) if idx + 1 < len(periods) else None
            vals.append(_growth(cur, prev))
        rows.append({"label": label, "vals": vals, "pct": True, "suffix": ""})

    growth_row("Tăng trưởng doanh thu (QoQ)", "IS", "doanh_thu_thuan")
    growth_row("Tăng trưởng LNST (QoQ)", "IS", "ln_sau_thue")

    # Dựng DataFrame
    data = []
    for r in rows:
        data.append([r["label"]] + r["vals"])
    df_out = pd.DataFrame(data, columns=["Chỉ số"] + periods)
    # Lưu metadata định dạng để excel_builder dùng
    df_out.attrs["fmt"] = [(r["pct"], r["suffix"]) for r in rows]
    return df_out


def valuation_snapshot(fin: dict) -> dict:
    """Định giá hiện tại: P/E (TTM), P/B, vốn hoá — từ overview + BCTC đầy đủ."""
    ov = fin["overview"]
    mc = ov.get("market_cap")
    full_is = fin["raw_full"]["income"]
    full_bs = fin["raw_full"]["balance"]
    is_periods = [c for c in full_is.columns if c != "Chỉ tiêu"]

    # LNST cổ đông mẹ lũy kế 4 quý gần nhất
    ttm = 0.0
    got = 0
    for p in is_periods[:config.TTM_QUARTERS]:
        v = _lookup(full_is, config.IS["ln_cong_ty_me"], p)
        if v is not None:
            ttm += v
            got += 1
    ttm = ttm if got > 0 else None

    latest_p = [c for c in full_bs.columns if c != "Chỉ tiêu"][0]
    equity = _lookup(full_bs, config.BS["von_chu_so_huu"], latest_p)

    return {
        "market_cap": mc,
        "pe_ttm": _div(mc, ttm),
        "pb": _div(mc, equity),
        "ln_ttm": ttm,
        "von_chu_so_huu": equity,
    }
