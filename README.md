🚢 AutoCAD .NET Plugin Template
Template chuẩn hóa dành cho việc phát triển các công cụ trên nền tảng AutoCAD .NET API. Bộ khung này được xây dựng dựa trên kiến trúc Module hóa và nguyên lý Clean Code, giúp bóc tách dữ liệu hình học (MTO) chính xác và nhanh chóng.

🏗️ Cấu trúc dự án (Architecture)
Dự án được phân chia thành các Module độc lập để dễ dàng bảo trì và mở rộng:

📁 Models: Chứa các lớp định nghĩa đối tượng (POCO). Tách biệt hoàn toàn với thư viện AutoCAD.

📁 Services: "Trái tim" của Plugin. Chứa ExtractionService thực hiện logic

📁 UI: Giao diện người dùng dựa trên WPF tích hợp vào AutoCAD PaletteSet.

📁 Utilities: Các hàm tiện ích dùng chung (Tính COG, đổi đơn vị, định dạng tọa độ WCS).

📁 Commands: Nơi đăng ký lệnh (CommandMethod) và quản lý vòng đời của Palette (Singleton).

Test quá trình checkin = Commit (luôn phải có comments xem là nội dung thay đổi là gì)
Quá trình checkout = Pull (kéo dự án từ Github về local)

B1: Trước khi làm thì luôn PULL về để đồng bộ hóa với hệ thống
B2: Khi kết thúc phiên làm việc thì cần Commit & Sync để đồng bộ lên hệ thống
Note: Trường hợp xảy ra Conflict

- Nhiều hơn 1 người cùng sửa 1 nội dung và up lên hệ thống?

* Phân phạm vi công việc cụ thể, tránh chồng lấn
* Nếu có chồng lấn, Github cho phép tất cả các update đó đều được thực hiện đồng bộ nhưng báo Conflict trên hệ thống, khi đó admin sẽ quyết định

Module 1 <--> function a
Module 2 <--> function a
Module 3 <--> function a
Nhưng khi cần thay đổi function a để các hệ thống về sau đơn giản hơn, tối ưu hơn thì một số user có thể tùy chỉnh function a này?

Test chia Branch để up lên hệ thống chứ không trực tiếp commit thẳng lên nhánh main. Khi đó admin sẽ quyết định phiên bản mới có ok không, nếu có thì admin sẽ merge vào nhánh main, nếu không thì sẽ reject để làm lại

Test pull from local
HƯỚNG DẪN SỬ DỤNG NHANH ADD-IN "DATA CENTER" (INVENTOR)
Add-in Data Center là công cụ giúp tra cứu, chỉnh sửa iProperties và đồng bộ dữ liệu Vault/BOM trực tiếp trên giao diện thiết kế 3D của Autodesk Inventor cực kỳ nhanh chóng.

---

1. HƯỚNG DẪN CÀI ĐẶT
   Bước 1: Copy toàn bộ thư mục InventorDrawingPlugin (bên trong có chứa 2 file là .dll và .addin).
   Bước 2: Dán (Paste) thư mục đó vào đường dẫn sau trên máy tính của bạn:

C:\Users\<Tên_Người_Dùng>\AppData\Roaming\Autodesk\Inventor 2023\Addins
(Ví dụ trên máy của bạn, <Tên_Người_Dùng> sẽ là Vuqu)

⚠️ LƯU Ý QUAN TRỌNG: Thư mục AppData mặc định bị hệ điều hành Windows ẩn đi. Để nhìn thấy nó, bạn làm như sau:

Windows 10: Mở File Explorer -> Bấm lên thẻ View ở trên cùng -> Tích chọn vào ô Hidden items.

Windows 11: Mở File Explorer -> Bấm nút View (Có icon con mắt hoặc menu thả xuống) -> Chọn Show -> Tích chọn Hidden items.

2. KHỞI ĐỘNG CÔNG CỤ
   Mở phần mềm Autodesk Inventor.

Mở một bản vẽ Lắp ráp (Assembly - .iam) hoặc Chi tiết (Part - .ipt) bất kỳ.

Nhìn lên thanh công cụ (Ribbon), chọn thẻ Add-Ins.

Tìm nhóm công cụ có tên Data Center và bấm vào nút Model Data (Có biểu tượng Icon của công ty).

Một bảng giao diện (Palette) sẽ xuất hiện bên cạnh màn hình làm việc của bạn.

3. CÁC TÍNH NĂNG CHÍNH
   🔍 Tra cứu tức thì (One-Click Inspect)
   Chỉ cần click chuột trái vào bất kỳ một chi tiết (Part), bề mặt (Face), hoặc cạnh (Edge) nào trên mô hình 3D, toàn bộ thông tin iProperties, Vật liệu, Khối lượng và Dữ liệu BOM (Item No, Số lượng) sẽ tự động hiện ra ngay lập tức trên bảng Data Center.

✍️ Sửa dữ liệu siêu tốc (Pop-up Edit)
Bạn không cần phải mở bảng iProperties thủ công như trước nữa!

Những ô thông tin có viền màu xanh dương (Part Number, Title, Description, Revision, Designer) là những ô cho phép sửa trực tiếp.

Cách sửa: Click đúp (Double-click) vào ô muốn sửa. Một hộp thoại nhỏ sẽ bật lên. Gõ thông tin mới (Hỗ trợ gõ cả Tiếng Anh và Tiếng Việt mượt mà) và nhấn phím Enter để lưu thẳng vào file.

☁️ Tích hợp Vault (Smart Vault Integration)
Nút "Open File Location": Bấm để mở thư mục chứa file 3D hiện tại trên Windows.

Nút "Get Vault Drawing": Bấm để Add-in tự động quét trên hệ thống Vault, tải về và mở ngay lập tức bản vẽ kỹ thuật 2D (.idw / .dwg) của chi tiết bạn đang chọn mà không cần thao tác tìm kiếm thủ công.

Chúc bạn có những giờ làm việc năng suất với công cụ mới này!
