"""Cấu hình cho tool báo cáo tài chính.

Tập trung mọi hằng số & mapping tên chỉ tiêu ở đây để dễ chỉnh khi vnstock
đổi nhãn hoặc khi muốn thêm/bớt chỉ tiêu.
"""

# --- Giới hạn tốc độ ---
# Trần request/phút của vnstock theo gói: Guest 20, Community (có API key) 60,
# Sponsor 180-600. Tool đặt trần chủ động DƯỚI ngưỡng vì gói Guest tự kết thúc
# tiến trình khi vượt. Đã đăng ký API key Community -> dùng 55 (đệm dưới 60).
# Nếu quay lại gói Guest thì hạ về 18.
MAX_RPM = 55

# --- Nguồn & phạm vi ---
SOURCE = "VCI"                     # Nguồn dữ liệu vnstock (VCI ổn định cho BCTC)
EXCHANGES = ["HOSE", "HNX"]        # Sàn cần lấy
STOCK_TYPE = "stock"              # Lọc chỉ cổ phiếu (bỏ trái phiếu, CW, ETF...)
PERIOD = "quarter"                # 'quarter' hoặc 'year'
DISPLAY_PERIODS = 3                # Số kỳ hiển thị (3 quý gần nhất)
TTM_QUARTERS = 4                   # Số quý dùng tính lũy kế 12 tháng (P/E TTM)

# Tên các cột định danh do vnstock trả về (sẽ bị loại khỏi phần hiển thị)
ID_COLS = ["item", "item_en", "item_id"]
ITEM_COL = "item"                 # Cột chứa tên chỉ tiêu tiếng Việt

# --- Tên sheet trong file Excel (đúng thứ tự) ---
SHEET_OVERVIEW = "Tổng quan"
SHEET_BALANCE = "Bảng cân đối kế toán"
SHEET_INCOME = "Kết quả kinh doanh"
SHEET_CASHFLOW = "Lưu chuyển tiền tệ"
SHEET_RATIOS = "Chỉ số tài chính"

# --- Mapping tên chỉ tiêu (theo nhãn tiếng Việt vnstock trả về) ---
# Mỗi chỉ tiêu là DANH SÁCH tên thay thế: tên [0] cho DN phi tài chính, các tên
# sau cho ngân hàng/CK (cấu trúc BCTC khác). Lấy tên KHỚP ĐẦU TIÊN có trong báo cáo.
BS = {
    "tong_tai_san": ["TỔNG CỘNG TÀI SẢN", "TỔNG TÀI SẢN"],
    "tai_san_ngan_han": ["TÀI SẢN NGẮN HẠN"],
    "tien": ["Tiền và tương đương tiền", "Tiền mặt, vàng bạc, đá quý"],
    "phai_thu": ["Các khoản phải thu"],
    "hang_ton_kho": ["Hàng tồn kho, ròng"],
    "tai_san_dai_han": ["TÀI SẢN DÀI HẠN"],
    "no_phai_tra": ["NỢ PHẢI TRẢ", "TỔNG NỢ PHẢI TRẢ"],
    "no_ngan_han": ["Nợ ngắn hạn"],
    "no_dai_han": ["Nợ dài hạn"],
    "vay_ngan_han": ["Vay ngắn hạn"],
    "vay_dai_han": ["Vay dài hạn"],
    "von_chu_so_huu": ["Vốn chủ sở hữu", "VỐN CHỦ SỞ HỮU"],
    "von_gop": ["Vốn góp", "Vốn điều lệ"],
}
IS = {
    "doanh_thu_thuan": ["Doanh thu thuần", "Tổng thu nhập hoạt động"],
    "gia_von": ["Giá vốn hàng bán"],
    "loi_nhuan_gop": ["Lợi nhuận gộp"],
    "dt_tai_chinh": ["Doanh thu hoạt động tài chính"],
    "cp_tai_chinh": ["Chi phí tài chính"],
    "cp_lai_vay": ["Chi phí lãi vay"],
    "cp_ban_hang": ["Chi phí bán hàng"],
    "cp_qldn": ["Chi phí quản lý doanh nghiệp"],
    "ln_hdkd": ["Lãi/(lỗ) từ hoạt động kinh doanh"],
    "ln_truoc_thue": ["Lãi/(lỗ) trước thuế", "Tổng lợi nhuận/lỗ trước thuế"],
    "ln_sau_thue": ["Lãi/(lỗ) thuần sau thuế", "Lợi nhuận sau thuế"],
    "ln_cong_ty_me": ["Lợi nhuận của Cổ đông của Công ty mẹ", "Cổ đông của Công ty mẹ"],
    "eps": ["Lãi cơ bản trên cổ phiếu (VND)"],
}
CF = {
    "cfo": ["Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh"],
    "cfi": ["Lưu chuyển tiền thuần từ hoạt động đầu tư"],
    "cff": ["Lưu chuyển tiền thuần từ hoạt động tài chính"],
    "capex": ["Tiền chi để mua sắm, xây dựng TSCĐ và các tài sản dài hạn khác"],
    "luu_chuyen_thuan": ["Lưu chuyển tiền thuần trong kỳ"],
    "tien_cuoi_ky": ["Tiền và tương đương tiền cuối kỳ"],
}

# Khối "chỉ tiêu then chốt" hiển thị ở sheet Tổng quan (nhãn -> nguồn.key)
KEY_INDICATORS = [
    ("Doanh thu thuần", "IS", "doanh_thu_thuan"),
    ("Lợi nhuận gộp", "IS", "loi_nhuan_gop"),
    ("LN thuần sau thuế", "IS", "ln_sau_thue"),
    ("LNST cổ đông công ty mẹ", "IS", "ln_cong_ty_me"),
    ("EPS (VND)", "IS", "eps"),
    ("Tổng tài sản", "BS", "tong_tai_san"),
    ("Nợ phải trả", "BS", "no_phai_tra"),
    ("Vốn chủ sở hữu", "BS", "von_chu_so_huu"),
    ("Dòng tiền HĐKD (CFO)", "CF", "cfo"),
]
