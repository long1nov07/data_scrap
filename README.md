# Tool thu thập & tổng hợp Báo cáo tài chính HOSE + HNX

Tự động tải báo cáo tài chính của các doanh nghiệp niêm yết trên **HOSE** và **HNX**
(3 quý gần nhất) và xuất **mỗi doanh nghiệp một file Excel** gồm 5 sheet:

| Sheet | Nội dung |
|-------|----------|
| **Tổng quan** | Thông tin mã, giá, vốn hoá, ngành, P/E (TTM), P/B + bảng chỉ tiêu then chốt qua 3 kỳ |
| **Bảng cân đối kế toán** | Toàn bộ khoản mục, 3 quý gần nhất + %QoQ |
| **Kết quả kinh doanh** | Doanh thu → lợi nhuận, 3 quý + %QoQ |
| **Lưu chuyển tiền tệ** | Dòng tiền HĐKD / đầu tư / tài chính, 3 quý + %QoQ |
| **Chỉ số tài chính** | Chỉ số tự tính: sinh lời, thanh khoản, đòn bẩy, hiệu quả, tăng trưởng |

Nguồn dữ liệu: thư viện [`vnstock`](https://github.com/thinh-vu/vnstock) (miễn phí, lấy từ VCI).

## Cài đặt

```bash
python -m pip install -r requirements.txt
```

Yêu cầu Python 3.10+.

## API key & tốc độ

vnstock giới hạn số request/phút theo gói:

| Gói | Trần | `MAX_RPM` khuyến nghị | Thời gian chạy full (~700 mã) |
|-----|------|----------------------|-------------------------------|
| Guest (không key) | 20/phút | 18 | ~2–2,5 giờ |
| **Community (API key miễn phí)** | 60/phút | 55 | **~30 phút** |
| Sponsor | 180–600/phút | 170+ | vài phút |

> ⚠️ Gói Guest **tự kết thúc tiến trình** khi vượt 20/phút. Tool có bộ điều tiết
> chủ động (`MAX_RPM` trong `config.py`) giữ mọi cửa sổ 60s dưới ngưỡng.

**Nạp API key** (lấy miễn phí tại https://vnstocks.com/login):

- Cách 1 — biến môi trường (khả chuyển giữa máy):
  ```bash
  # Windows PowerShell
  $env:VNSTOCK_API_KEY = "vnstock_xx...";  python main.py
  ```
- Cách 2 — đăng ký 1 lần, vnstock tự lưu vào `~/.vnstock/api_key.json` và dùng lại:
  ```python
  import vnstock; vnstock.register_user("vnstock_xx...")
  ```

Sau khi có key, đặt `MAX_RPM = 55` trong `config.py` (mặc định đã là 55).

## Sử dụng

```bash
# Thử 1 mã (khuyến nghị chạy đầu tiên để kiểm tra)
python main.py --symbols FPT

# Vài mã khác ngành
python main.py --symbols FPT,VCB,SSI

# Smoke test 10 mã đầu của toàn sàn
python main.py --limit 10

# Chạy đầy đủ toàn bộ HOSE + HNX (~700 mã, ~15-25 phút)
python main.py

# Chạy lại: tự bỏ qua mã đã có file (resume). Dùng --force để tạo lại.
python main.py
python main.py --force
```

Kết quả nằm trong thư mục `output/<SÀN>/<MÃ>.xlsx`, ví dụ `output/HOSE/FPT.xlsx`.

### So sánh các mã cùng ngành

Sau khi đã có các file mã, tạo bảng so sánh theo ngành:

```bash
python compare.py                        # đọc ./output, xuất output/So_sanh_nganh.xlsx
python compare.py --out-file D:/ss.xlsx  # đổi nơi lưu
```

- Đọc trực tiếp từ các file trong `output/HOSE` và `output/HNX` (không gọi lại API).
- Nhóm theo **ngành (sector)** — 19 nhóm. Xuất 1 file Excel: sheet **Mục lục** +
  mỗi ngành 1 sheet, mỗi dòng 1 mã (xếp theo vốn hoá), có **AutoFilter** để tự
  lọc/sắp xếp và dòng **Trung vị ngành**.
- Chỉ tiêu so sánh: Định giá (Vốn hoá, Giá, P/E, P/B), Sinh lời (ROE, ROA, biên
  LN gộp/ròng), Quy mô & tăng trưởng (Doanh thu, LNST quý, tăng trưởng DT QoQ),
  An toàn (Nợ/VCSH, thanh toán hiện hành).
- Đơn vị tiền trong bảng so sánh: **tỷ đồng** (riêng Giá là VND).

### Tham số

| Tham số | Mặc định | Ý nghĩa |
|---------|----------|---------|
| `--exchanges` | `HOSE,HNX` | Sàn cần lấy |
| `--symbols` | (trống) | Danh sách mã tự chọn, bỏ qua `--exchanges` |
| `--period` | `quarter` | `quarter` hoặc `year` |
| `--periods` | `3` | Số kỳ hiển thị |
| `--output` | `output` | Thư mục kết quả |
| `--sleep` | `0.6` | Nghỉ giữa các mã (giây) – tăng nếu bị chặn tốc độ |
| `--retries` | `3` | Số lần thử lại khi lỗi mạng |
| `--limit` | `0` | Chỉ xử lý N mã đầu (0 = tất cả) |
| `--force` | tắt | Tạo lại cả file đã có |

## Cơ chế chạy bền bỉ

- **Resume/checkpoint:** ghi `output/progress.csv`; chạy lại tự bỏ qua mã đã xong → an toàn khi bị gián đoạn giữa chừng.
- **Rate-limit + retry:** nghỉ `--sleep` giây giữa các mã, tự thử lại (backoff) khi lỗi tải.
- **Log:** `output/run.log` (đầy đủ), `output/run_summary.txt` (tóm tắt), cột trạng thái trong `progress.csv` (`OK` / `NO_DATA` / `ERROR`).

## Cấu trúc mã nguồn

```
config.py         # Nguồn, sàn, tên sheet, mapping tên chỉ tiêu (có alias ngân hàng)
data_source.py    # Bọc vnstock: list_companies(), get_financials()
analysis.py       # Tự tính chỉ số tài chính, chỉ tiêu then chốt, định giá
excel_builder.py  # Dựng file Excel 5 sheet (định dạng, màu, %QoQ)
main.py           # CLI, điều phối, checkpoint, rate-limit, logging
```

## Ghi chú & giới hạn

- **Bản community của vnstock giới hạn 4 kỳ** báo cáo → tool lấy 3 quý gần nhất là phù hợp.
- Hàm `ratio()` của vnstock (bản community) trả về kỳ cũ nên tool **tự tính chỉ số** từ 3 báo cáo để đảm bảo đúng kỳ.
- **Doanh nghiệp phi tài chính:** đầy đủ mọi chỉ số.
- **Ngân hàng:** báo cáo thô đầy đủ; đã hỗ trợ ROA, ROE, biên LN, đòn bẩy, P/E, P/B. Các chỉ số không áp dụng (thanh toán hiện hành, vòng quay HTK…) để trống.
- **Công ty chứng khoán / bảo hiểm:** báo cáo thô đầy đủ; một số chỉ số để trống do cấu trúc khác.
- Đơn vị số liệu trên các sheet báo cáo: **VND**.
