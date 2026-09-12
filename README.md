# Trợ lý AI xử lý công việc (OpenRouter + Excel)

Web app mobile-first cho phép:
- Chat với AI (nhớ lịch sử hội thoại trong 2 ngày)
- Upload file Excel, gắn với tên khách hàng (bạn tự chọn khi upload)
- Nhờ AI xử lý dữ liệu theo yêu cầu, tải về file Excel kết quả
- Gán file đã xử lý cho một nhân viên cụ thể
- Xem/xuất thống kê cuối ngày: khách hàng nào → gửi cho nhân viên nào

## 1. Cài đặt

Cần cài [Node.js](https://nodejs.org) bản **22.13 trở lên** (khuyến nghị dùng bản LTS mới nhất) trên máy chủ/máy tính sẽ chạy app — vì app dùng module SQLite có sẵn trong Node (`node:sqlite`), không cần cài thêm gì khác, không cần trình biên dịch C++.

```bash
cd wife-chatbot
npm install
cp .env.example .env
```

Mở file `.env`, dán API key lấy từ https://openrouter.ai/keys vào `OPENROUTER_API_KEY`.

> Khi chạy sẽ thấy dòng cảnh báo `ExperimentalWarning: SQLite is an experimental feature...` — đây là cảnh báo bình thường của Node.js, không phải lỗi, app vẫn chạy đúng.

## 2. Chạy thử trên máy tính

```bash
npm start
```

Mở trình duyệt tại `http://localhost:3000`.

## 3. Deploy lên Render.com — Gói Free

⚠️ Bạn đã chọn dùng gói **Free** và chấp nhận: mỗi lần deploy lại code (hoặc service tự ngủ/khởi động lại sau 15 phút không dùng) thì **toàn bộ dữ liệu sẽ mất** (database, file đã upload, hồ sơ khách hàng, lịch sử chat...). Phù hợp nếu bạn chỉ cần dùng thử hoặc dùng theo phiên ngắn hạn, không cần lưu lâu dài.

1. Đưa code lên 1 repo GitHub (`git init && git add . && git commit -m "init" && git push`).
2. Vào [dashboard.render.com](https://dashboard.render.com) → **New** → **Blueprint** → chọn repo đó (file `render.yaml` có sẵn trong repo sẽ tự cấu hình gói Free, không có ổ đĩa cố định).
3. Render sẽ hỏi bạn nhập `OPENROUTER_API_KEY` → dán key từ https://openrouter.ai/keys.
4. Bấm **Apply**, đợi build xong → nhận link dạng `https://wife-chatbot-xxxx.onrender.com`.
5. Mở link đó trên điện thoại vợ bạn, có thể **"Thêm vào màn hình chính"** để dùng như app thật.

Lưu ý: gói Free sẽ **ngủ sau 15 phút không có ai truy cập**, lần mở lại đầu tiên sẽ mất khoảng 30-60 giây để "thức dậy" — đây là điều bình thường của gói Free, không phải lỗi.

*(Nếu sau này muốn giữ dữ liệu lâu dài không bị mất, có thể nâng lên gói Starter + gắn Persistent Disk — lúc đó chỉ cần đổi `plan: free` thành `plan: starter` trong `render.yaml` và thêm phần `disk` trỏ vào biến `DATA_DIR`, code đã hỗ trợ sẵn việc này.)*

## 4. Chạy thử trong mạng nhà (không deploy, chỉ test nhanh)

1. Chạy `npm start` trên máy tính.
2. Tìm địa chỉ IP máy tính (Windows: `ipconfig`, Mac: `ifconfig`).
3. Trên điện thoại (cùng wifi), mở trình duyệt và vào `http://<IP-máy-tính>:3000`.

## 5. Quy tắc chuẩn hoá SĐT / Seri sim (tự động)

Hệ thống này giả định dữ liệu chỉ gồm 2 loại: **SĐT** và **Seri sim**. Mỗi khi có dữ liệu mới (upload file, dán tay, AI xử lý, đính kèm trong chat), hệ thống **tự động nhận diện và chuẩn hoá**:

- **SĐT**: nhận diện chuỗi số dài 9 (giữ nguyên), 10 số bắt đầu bằng `0`, hoặc 11 số bắt đầu bằng `84` → luôn chuyển về **9 số cuối**.
- **Seri sim**: nhận diện chuỗi số dài 10 bắt đầu bằng `1`/`2` (giữ nguyên), hoặc 20 số → lấy **11 số cuối rồi bỏ đi số cuối cùng**, còn lại đúng **10 số**.
- Các giá trị khác (số tiền, số lượng, chữ...) không bị ảnh hưởng.

⚠️ Lưu ý: nếu file Excel gốc lưu Seri sim (20 số) dưới dạng **số (Number)** thay vì **văn bản (Text)**, bản thân Excel/JavaScript có thể đã làm tròn/mất vài chữ số cuối trước khi vào hệ thống (giới hạn của định dạng số). Để an toàn tuyệt đối, nên định dạng cột Seri sim thành **Text** trong file Excel gốc trước khi upload.

## 6. Thu hồi số liệu đã giao cho nhân viên

Ở tab **👤 Phân việc**, mỗi nhân viên trong danh sách "đã gán" có nút **×** để **thu hồi riêng** số liệu đã giao cho người đó (đưa các dòng đó về trạng thái chưa gán). Ngoài ra có nút **"🔄 Thu hồi tất cả"** để đặt lại toàn bộ phân việc của hồ sơ khách hàng đó về trạng thái ban đầu. Sau khi thu hồi, số liệu sẽ tự động biến mất khỏi thống kê cuối ngày và có thể gán lại cho người khác.

## 7. Tin nhắn bàn giao (SMS) để gửi Zalo

Mỗi khi gán dữ liệu cho 1 nhân viên (tab Phân việc), hệ thống tự sinh sẵn 1 đoạn tin nhắn liệt kê từng cặp **SĐT - Seri sim** vừa giao, hiện ngay trong ô có thể copy. Vợ bạn chỉ cần bấm **"📋 Sao chép tin nhắn"** rồi dán trực tiếp vào Zalo gửi cho nhân viên, không cần soạn tay.

## 4. Cách hoạt động của "bộ nhớ 2 ngày"

Mọi tin nhắn chat được lưu trong SQLite (`data.sqlite`) kèm timestamp. Khi AI trả lời, nó chỉ nhìn thấy các tin nhắn trong vòng 48 giờ gần nhất — tin nhắn cũ hơn sẽ tự bị xoá định kỳ (mỗi giờ) và không còn ảnh hưởng tới ngữ cảnh.

## 5. Luồng sử dụng thực tế cho vợ bạn

1. Vào tab **Khách hàng** → chọn file Excel → gõ tên khách hàng (vd: "Công ty ABC") → Tải lên.
2. Gõ yêu cầu xử lý (vd: "tính tổng đơn hàng theo ngày") → bấm **Xử lý & tạo file Excel** → tải file kết quả về.
3. Sang tab **Phân việc** → chọn đúng file vừa xử lý → gõ tên nhân viên → **Gán cho nhân viên**.
4. Cuối ngày, vào tab **Thống kê** → xem hoặc xuất file Excel tổng hợp "khách hàng nào đã gửi cho nhân viên nào".

## 6. Giới hạn hiện tại / có thể mở rộng thêm

- Hiện tại **chưa có đăng nhập** — ai có link đều dùng chung 1 phiên. Nếu cần riêng tư/bảo mật hơn (chỉ vợ bạn dùng được), nên thêm 1 lớp đăng nhập bằng mật khẩu đơn giản.
- Xử lý Excel qua AI hiện giới hạn khoảng 500 dòng/lần để tránh vượt giới hạn ngữ cảnh — nếu file lớn hơn, có thể chia nhỏ hoặc chuyển sang xử lý bằng công thức cố định (không qua AI) cho nhanh và rẻ hơn.
- Danh sách nhân viên/khách hàng hiện nhập tay tự do — có thể làm thành danh sách chọn sẵn (dropdown) để tránh gõ sai tên trùng lặp.
