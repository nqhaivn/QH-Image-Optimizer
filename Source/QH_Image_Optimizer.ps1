# Ép mã hóa console/log về UTF-8 BOM để log không bị lỗi dấu.
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$OutputEncoding = $utf8Bom
[Console]::OutputEncoding = $utf8Bom

# Lưu đường dẫn script để map function đúng trong runspace.
$script:SourceScriptPath = $PSCommandPath

# Bật/tắt giữ cửa sổ console khi chạy script.
$script:KeepConsoleOpen = $true

if (-not $script:KeepConsoleOpen) {
    if (-not ("Win32.NativeMethods" -as [type])) {
        Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @"
[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@
    }
    $consoleHandle = [Win32.NativeMethods]::GetConsoleWindow()
    if ($consoleHandle -ne [IntPtr]::Zero) {
        [Win32.NativeMethods]::ShowWindow($consoleHandle, 0) | Out-Null
    }
}

# Cờ nội bộ để dot-source file trong runspace mà không khởi tạo UI.
if ($null -eq $script:LogicBootstrap) {
    $script:LogicBootstrap = $false
}

# Danh sách input hợp lệ chuẩn, dùng chung cho UI và encode.
$script:supportedExtensions = @(
    ".jpg", ".jpeg", ".jpe", ".jfif", ".jif", ".pjpeg", ".pjp",
    ".png", ".bmp", ".tif", ".tiff", ".ico", ".icon", ".apng",
    ".webp", ".jxl", ".avif", ".heic", ".heif", ".gif"
)

# ==============================
# NGÔN NGỮ
# ==============================
# LogNoValidTasks là báo lỗi trong trường hợp chỉ có 1 ảnh, ảnh bị lỗi, bị khóa, hoặc không phải ảnh thật sự (ví dụ file rar đổi đuôi thành jpg).
$strings = @{}
$strings = @{
    vi = @{
        Title = "Tối Ưu Ảnh QH 1.0"
        AddFiles = "Thêm ảnh"
        AddFolder = "Thêm thư mục"
        ClearList = "Dọn sạch"
        Guide = "Hướng dẫn"
        Encode = "Chuyển đổi"
        LangVi = "Tiếng Việt"
        LangEn = "English"
        TooltipAddFiles = "Thêm ảnh vào danh sách để bắt đầu chuyển đổi"
        TooltipAddFolder = "Thêm thư mục chứa các tệp ảnh vào danh sách để bắt đầu chuyển đổi"
        TooltipClear = "Làm trống danh sách"
        TooltipGuide = "Mở hướng dẫn"
        TooltipEncode = "Bắt đầu chuyển đổi các tệp ảnh"
        OnlyParent = "Chỉ thư mục mẹ"
        IncludeSub = "Bao gồm thư mục con"
        UnsupportedMsg = "Không có ảnh hợp lệ. Phần mềm chỉ hỗ trợ đọc các định dạng sau: <EXT>"
        Ok = "OK"
        DialogDefaultTitle = "Thông báo"
        GuideContentUpdating = "【 1 】 GIỚI THIỆU`n`n    Tối Ưu Ảnh QH 1.0 là là giải pháp giao diện đồ họa (GUI) được phát triển trên nền tảng ImageMagick, chuyên tối ưu hóa quy trình xử lý ảnh hàng loạt. Phần mềm tập trung vào khả năng nén ảnh thông minh thông qua việc cân bằng giữa hệ số chất lượng và kích thước điểm ảnh. Kết quả thực nghiệm cho thấy dung lượng tệp tin có thể được tối ưu hóa chỉ còn 10-30% so với nguyên bản, trong khi vẫn duy trì độ sắc nét tiệm cận bản gốc.`n`n    Lưu ý kỹ thuật:`n    - Phần mềm bắt buộc cần ImageMagick để vận hành (đã đóng gói kèm). Nếu tệp này bị mất, vui lòng tải bản Portable (`"portable-Q16-HDRI-x64`" cho 64-bit hoặc `"x86`" cho 32-bit) tại: https://github.com/ImageMagick/ImageMagick/releases và đặt vào cùng thư mục chương trình.`n    - Để đạt chất lượng AVIF tốt nhất, phần mềm mặc định sử dụng avifenc.exe. Nếu bạn muốn dùng trình xử lý của ImageMagick, chỉ cần xóa avifenc.exe trong thư mục ImageMagick, hệ thống sẽ tự động chuyển đổi phương thức.`n    - Nếu máy tính của bạn đã có ImageMagick và đưa vào biến môi trường (PATH), bạn không cần đặt magick.exe cùng thư mục với phần mềm này. Phần mềm sẽ tự động tìm ImageMagick từ PATH.`n`n【 2 】 ĐỊNH DẠNG HỖ TRỢ`n`n    2.1 Đầu vào: Hỗ trợ các định dạng JPG, JPEG, JPE, JFIF, JIF, PJPEG, PJP, PNG, BMP, TIF, TIFF, ICO, ICON, APNG, WEBP, JXL, AVIF, HEIC, HEIF, GIF`n`n    2.2 Đầu ra: JPG, PNG, TIFF, ICO, AVIF, WEBP, JXL.`n    a) Đối với chế độ `"Giữ nguyên định dạng`": `n        - Phần mềm sẽ xuất ra đúng định dạng gốc của ảnh đầu vào. `n        - Ngoại trừ những định dạng sau sẽ được chuyển đổi sang JPG: JPG, JPEG, JPE, JFIF, JIF, PJPEG, PJP, HEIC, HEIF. `n        - TIFF, TIF sẽ được chuyển đổi sang TIFF. `n        - PNG, APNG, BMP, ICO, ICON sẽ được chuyển đổi sang PNG.`n`n    b) Lời khuyên: Thông thường hãy chọn JPG. Nếu cần lưu trữ bảo toàn chi tiết đến từng điểm ảnh (Pixel-perfect), hãy chọn PNG, TIF hoặc các định dạng hiện đại như WEBP, JXL, AVIF ở mức chất lượng 100%.`n`n`n【 3 】 CHẤT LƯỢNG VÀ THAM SỐ`n`n    Mức mặc định 80% phù hợp với hơn 90% nhu cầu thông thường. Bạn có thể tùy chỉnh tăng/giảm theo mục đích sử dụng.`n`n    Cách thức hoạt động:`n    - Với JPG, JXL, WEBP: Chất lượng được điều khiển qua tham số -quality của Magick. Ví dụ: Chọn 75% tương đương -quality 75. Riêng 100% của WEBP và JXL sẽ kích hoạt nén không mất dữ liệu (Lossless).`n`n    - Với AVIF: Để khắc phục nhược điểm cho chất lượng thấp của Magick đối với định dạng này, tác giả ưu tiên dùng avifenc.exe. `n    Bảng quy đổi dựa trên thực nghiệm (so sánh 1.840 ảnh qua MS-SSIM và Butteraugli và điều chỉnh để đạt chất lượng tương đương JPG):`n        Chọn | avifenc  | magick quality`n        100% | lossless | 100`n        95%  | crf 13   | 95`n        90%  | crf 16   | 90`n        85%  | crf 24   | 80`n        80%  | crf 28   | 70`n        75%  | crf 31   | 60`n        70%  | crf 36   | 50`n        65%  | crf 38   | 48`n        60%  | crf 40   | 44`n        50%  | crf 42   | 40`n        40%  | crf 44   | 36`n        30%  | crf 47   | 30`n        20%  | crf 51   | 20`n        10%  | crf 58   | 10`n`n    - Với ICO và định dạng khác: Magick tự động xử lý.`n`n    Lưu ý: Ở cùng một mức chất lượng, AVIF cho dung lượng nhỏ hơn so với JXL và WEBP.`n`n【 4 】 QUY TẮC THAY ĐỔI KÍCH THƯỚC`n`n    Kích thước tối đa: Là giới hạn cho cạnh dài nhất của ảnh. Ví dụ ảnh gốc 7680x5120, nếu đặt tối đa 3840, ảnh đầu ra sẽ là 3840x2560 (giữ đúng tỷ lệ).`n`n    Giới hạn thu nhỏ: Đây là `"chốt chặn`" để ảnh không bị thu quá nhỏ, gây mất chi tiết vĩnh viễn. `n        - Công thức: Kích thước đầu ra không nhỏ hơn (Gốc/Tỷ lệ). `n        - Ví dụ: Đặt tối đa 3840 nhưng giới hạn tỷ lệ 1.5. Với ảnh gốc 7680, kết quả sẽ là 5120 (vì 7680/1.5 = 5120, lớn hơn mức 3840 bạn muốn).`n`n    Tại sao phải Giới hạn thu nhỏ? Bởi vì có những bức ảnh mà hiện tại bạn muốn thu nhỏ để giảm dung lượng khi lưu trữ, nhưng có thể muốn khôi phục về kích thước gốc sau này. Nếu ảnh gốc lớn và chứa nhiều chi tiết, như ảnh chụp một lễ hội với hàng trăm người, việc thu nhỏ quá mức sẽ dẫn đến hậu quả là ảnh đầu ra vĩnh viễn mất đi các chi tiết nhỏ. Qua kinh nghiệm từng phóng to/thu nhỏ hàng nghìn ảnh, tác giả nhận thấy đối với phần lớn trường hợp, chỉ nên thu nhỏ bằng 1.5 cho đến 2 lần kích thước gốc. Tuy nhiên, nếu ảnh của bạn chỉ gồm các chi tiết lớn thì hoàn toàn có thể thu nhỏ hơn nữa.`n`n    Nếu Giới hạn thu nhỏ được đặt thành 1, kích thước ảnh đầu ra sẽ không được nhỏ hơn (Kích thước ảnh gốc)/1, đồng nghĩa với việc giữ nguyên kích thước gốc. `n`n【 5 】 MẸO VÀ BÁO LỖI`n`n    Bạn có thể kéo và thả tệp ảnh hoặc thư mục chứa ảnh vào phần mềm, danh sách các ảnh hợp lệ sẽ được liệt kê và sẵn sàng chuyển đổi.`n`n    Nếu bạn gặp lỗi `"Không thể đọc ảnh gốc hoặc ghi ảnh đã chuyển đổi vào ổ đĩa`" thì có thể do một trong các nguyên nhân:`n        - Ảnh gốc đã bị thu hồi quyền đọc hoặc bị khóa.`n        - Không có quyền lưu ảnh đã chuyển đổi vào ổ đĩa của ảnh gốc.`n        - Ổ đĩa đã đầy nên không thể lưu ảnh đã chuyển đổi.`n    `n【 6 】 VỀ TÁC GIẢ`n`n    Phần mềm Tối Ưu Ảnh QH 1.0 được viết bởi facebook.com/nqhaivn/, phát hành ngày 24/03/2026.`n`n    Mọi hành vi chỉnh sửa lại mã nguồn mà chưa được sự đồng ý bằng văn bản của tác giả thì đều là trộm cắp.`n`n    Mọi góp ý, báo lỗi, có nhu cầu tùy chỉnh phần mềm theo sở thích của bản thân hoặc muốn đề nghị thiết kế phần mềm khác, vui lòng liên lạc với email: nqhai86@gmail.com`n`n    Trân trọng!"
        GuideTitle = "Hướng dẫn"
        StatusName = "Tên"
        StatusSize = "Kích thước"
        StatusBytes = "Dung lượng"
        LogConvertSuccess = "✔ Thành công"
        LogMagickMissing = "Không tìm thấy ImageMagick (magick.exe). Chương trình này cần ImageMagick để hoạt động. Mở Hướng dẫn để biết thêm chi tiết."
        LogNoValidTasks = "Từ chối: ImageMagick không có quyền truy cập hoặc không thể xử lý định dạng này."
        LogEncodeFailed = "{0} ❌ Thất bại | Lý do: {1}"
        LogOutputWriteFailed = "Không thể đọc ảnh gốc hoặc ghi ảnh đã chuyển đổi vào ổ đĩa. Mở Hướng dẫn để biết thêm chi tiết."
        LogIcoOnlyPngOrTiffSupported = "ICO chỉ có thể chuyển đổi sang PNG hoặc TIFF"
        LogMultiFrameToIcoNotSupported = "Không thể chuyển đổi ảnh đa khung hình sang ICO"
        LogRetryAttempt = "Thử lại lần {0}: {1}"
        LogRunspaceError = "Có lỗi xảy ra khi chạy. Vui lòng thử lại."
        LogAddedFilesWithErrors = "Đã thêm {0} ảnh vào danh sách. Trong đó {1} ảnh bị lỗi, không thể lấy được thông tin về kích thước."
        LogAddedFiles = "Đã thêm {0} ảnh vào danh sách."
        MissingFileMsg = "Tệp ảnh không còn tồn tại! Nguyên nhân: Đã đổi tên, bị di chuyển hoặc bị xóa."
        Remove = "Xóa khỏi danh sách"
        OpenFolder = "Mở trong thư mục"
        PickScope = "Bạn muốn thêm vào chỉ những tệp ảnh trong thư mục mẹ, hay toàn bộ tệp ảnh trong các thư mục con?"
        EncTo = "Chuyển đổi sang"
        Quality = "Chất lượng"
        MaxSize = "Kích thước tối đa *"
        Minimize = "Giới hạn thu nhỏ *"
        Prefix = "Thêm tiền tố"
        Suffix = "Thêm hậu tố"
        SaveWhere = "Lưu ảnh đầu ra"
        KeepMeta = "Siêu dữ liệu"
        RunMode = "Chế độ chạy"
        SaveSame = "Cùng chỗ ảnh gốc"
        SaveNew = "Trong thư mục mới"
        MetaAll = "Giữ lại"
        MetaNone = "Xóa bỏ"
        Yes = "Có"
        No = "Không"
        Parallel = "Đa luồng"
        Sequential = "Tuần tự"
        Original = "Giữ nguyên định dạng"
        Q100 = "100%"
        Q95 = "95%"
        Q90 = "90%"
        Q85 = "85%"
        Q80 = "80%"
        Q75 = "75%"
        Q70 = "70%"
        Q65 = "65%"
        Q60 = "60%"
        Q50 = "50%"
        Q40 = "40%"
        Q30 = "30%"
        Q20 = "20%"
        Q10 = "10%"
        Lossless = "Nguyên bản"
        IcoQualityNormal = "Tối ưu"
        IcoQualityBest = "Tốt nhất"
        InvalidChar = "Không được nhập vào các ký tự {0} bởi vì Windows không cho phép."
        LoadingList = "Đang nạp danh sách ảnh..."
        TooltipEncTo = "Bạn muốn chuyển đổi ảnh gốc sang định dạng nào? `n`nLưu ý: Chế độ 'Giữ nguyên định dạng' sẽ xuất ra định dạng gốc của ảnh đầu vào. Tuy nhiên, để đảm bảo tính tương thích tốt nhất, các nhóm sau sẽ được quy đổi về định dạng chuẩn: `n- Các biến thể của JPG (JFIF, JPE...) và HEIC/HEIF sẽ được lưu dưới dạng: .jpg`n- Các biến thể của TIFF sẽ được lưu dưới dạng: .tiff `n- Các định dạng PNG, APNG, ICO, ICON và BMP sẽ được lưu dưới dạng: .png"
        TooltipQuality = "Chọn chất lượng ảnh đầu ra. `n`nChú ý: `n- Khi chuyển đổi sang PNG hoặc TIFF, chất lượng luôn là 100%.`n- Khi chuyển đổi sang ICO, chất lượng 'Tối ưu' sẽ đưa ảnh về 256 màu và xóa nền trong suốt."
        TooltipMaxSize = "Thiết lập cạnh dài nhất của ảnh đầu ra.`n`nVí dụ: Ảnh gốc 6000x4000 px, nếu đặt 3840 thì ảnh đầu ra sẽ là 3840x2560 px. `n`nNếu ảnh gốc vốn đã nhỏ hơn mức này, kích thước sẽ được giữ nguyên."
        TooltipMinimize = "Hệ số bảo toàn chi tiết: Ngăn không cho ảnh bị thu nhỏ quá mức so với gốc.`n`nCông thức: (Kích thước gốc / Hệ số này). `n`nKhuyến nghị: Đặt 1.5 đến 2 để sau này có thể phóng to lại mà không bị vỡ hình. `n`nVí dụ: Ảnh gốc 4500x3000 px, nếu đặt hệ số 1.5 thì ảnh đầu ra tối thiểu phải là 3000x2000 px, ngay cả khi mục tiêu 'Kích thước tối đa' của bạn thấp hơn mức đó."
        TooltipPrefix = "Thêm tiền tố cho ảnh đầu ra. Công thức: <Tiền tố><Tên ảnh gốc>. Nếu tên gốc là 'photo.jpg' và tiền tố là 'New_' thì ảnh đầu ra sẽ có tên 'New_photo.jpg'"
        TooltipSuffix = "Thêm hậu tố cho ảnh đầu ra. Công thức: <Tên ảnh gốc><Hậu tố>. Nếu tên gốc là 'photo.jpg' và hậu tố là '_New' thì ảnh đầu ra sẽ có tên 'photo_New.jpg'"
        TooltipSaveWhere = "Chọn nơi lưu ảnh sau khi xử lý.`n`n⚠️ CHÚ Ý: Nếu chọn 'Cùng chỗ ảnh gốc' và định dạng trùng nhau, ảnh gốc sẽ bị GHI ĐÈ (mất ảnh cũ). Hãy kiểm tra kỹ thiết lập trước khi chạy hàng loạt.`n`n'Trong thư mục mới': Phần mềm tự tạo thư mục riêng để bảo vệ ảnh gốc. Tên thư mục sẽ bao gồm thời gian tạo và các thông số thiết lập.`nVí dụ: 20260325_052511_max_3840_limit_1.5_JPG_quality_75"
        TooltipMeta = "Gồm thông tin về: thời gian tạo, sửa, chụp của ảnh; góc xoay; độ sáng; thiết bị chụp; tọa độ địa lý (nếu máy ảnh hoặc điện thoại hỗ trợ)... `n`nViệc xóa bỏ siêu dữ liệu sẽ giúp bảo vệ quyền riêng tư khi chia sẻ ảnh, ngăn người khác biết được những thông tin trên. Tuy nhiên, bạn cũng sẽ mất thông tin hữu ích về thời gian chụp. `n`nNếu là ảnh phóng sự hoặc cần được lưu trữ làm tư liệu cho tương lai thì nên giữ lại siêu dữ liệu."
        TooltipRun = "Chạy đa luồng cho tốc độ xử lý cao hơn, nhưng cần tận dụng tối đa tài nguyên của CPU. Nếu máy tính của bạn cấu hình yếu hoặc đang chạy nhiều tác vụ khác, chọn 'Tuần tự' để an toàn hơn"
        AfterEncodeTitle = "Thông báo"
        AfterEncodeTemplate = "➤ Đã chuyển đổi: Thành công {0} ảnh | Thất bại {1} ảnh | Dung lượng ban đầu: {2} | Dung lượng hiện tại: {3} ({4}) | Thời gian hoàn thành: {5} phút {6} giây."
        IcoTitle = "Chọn những kích thước được lồng trong ico"
        ToolsMissingTitle = "Thiếu ImageMagick"
        ToolsMissingMsg = "Không tìm thấy magick.exe. Vui lòng mở Hướng dẫn để biết thêm chi tiết."
    }
    en = @{
        Title = "QH Image Optimizer 1.0"
        AddFiles = "Add Images"
        AddFolder = "Add Folder"
        ClearList = "Clear List"
        Guide = "User Guide"
        Encode = "Convert"
        LangVi = "Tiếng Việt"
        LangEn = "English"
        TooltipAddFiles = "Add images to the list to start converting"
        TooltipAddFolder = "Add a folder containing image files to the list to start converting"
        TooltipClear = "Clear the list"
        TooltipGuide = "Open user guide"
        TooltipEncode = "Start converting image files"
        OnlyParent = "Parent folder only"
        IncludeSub = "Include subfolders"
        UnsupportedMsg = "No valid images found. The software only supports the following formats: <EXT>"
        Ok = "OK"
        DialogDefaultTitle = "Notification"
        GuideContentUpdating = "【 1 】 INTRODUCTION`n`n    QH Image Optimizer 1.0 is a Graphical User Interface (GUI) solution developed on the ImageMagick platform, specialized for optimizing batch image processing workflows. The software focuses on intelligent image compression by balancing quality factors and pixel dimensions. Experimental results show that file sizes can be optimized to just 10-30% of the original while maintaining sharpness near-identical to the source.`n`n    Technical Notes:`n    - ImageMagick is required for operation (pre-packaged). If this file is missing, please download the Portable version (`"portable-Q16-HDRI-x64`" for 64-bit or `"x86`" for 32-bit) at: https://github.com/ImageMagick/ImageMagick/releases and place it in the same program directory.`n    - To achieve the best AVIF quality, the software defaults to using avifenc.exe. If you prefer to use ImageMagick's processor, simply delete avifenc.exe; the system will automatically switch methods.`n    - If ImageMagick is already installed and added to your system's Environment Variables (PATH), you do not need to place magick.exe in the program folder. The software will automatically detect and use ImageMagick from the PATH.`n`n【 2 】 SUPPORTED FORMATS`n`n    2.1 Input: Supports JPG, JPEG, JPE, JFIF, JIF, PJPEG, PJP, PNG, BMP, TIF, TIFF, ICO, ICON, APNG, WEBP, JXL, AVIF, HEIC, HEIF, GIF.`n`n    2.2 Output: JPG, PNG, TIFF, ICO, AVIF, WEBP, JXL.`n    a) For `"Keep original format`" mode: `n        - The software will output the exact original format of the input image. `n        - Except for the following formats, which will be converted to JPG: JPG, JPEG, JPE, JFIF, JIF, PJPEG, PJP, HEIC, HEIF. `n        - TIFF and TIF will be converted to TIFF. `n        - PNG, APNG, BMP, ICO, ICON will be converted to PNG.`n    b) Advice: Generally, choose JPG. If you need pixel-perfect preservation, select PNG, TIF, or modern formats like WEBP, JXL, and AVIF at 100% quality.`n`n`n【 3 】 QUALITY AND PARAMETERS`n`n    The default level of 80% is suitable for over 90% of common needs. You can customize this up or down based on your purpose.`n`n    How it works:`n    - For JPG, JXL, WEBP: Quality is controlled via Magick's -quality parameter. Example: Selecting 75% is equivalent to -quality 75. Note that 100% for WEBP and JXL will activate Lossless compression.`n`n    - For AVIF: To overcome Magick's low-quality drawbacks for this format, the author prioritizes avifenc.exe.`n    Conversion table based on experiments (comparing 1,840 images via MS-SSIM and Butteraugli, adjusted to match JPG quality):`n        Chọn | avifenc  | magick quality`n        100% | lossless | 100`n        95%  | crf 13   | 95`n        90%  | crf 16   | 90`n        85%  | crf 24   | 80`n        80%  | crf 28   | 70`n        75%  | crf 31   | 60`n        70%  | crf 36   | 50`n        65%  | crf 38   | 48`n        60%  | crf 40   | 44`n        50%  | crf 42   | 40`n        40%  | crf 44   | 36`n        30%  | crf 47   | 30`n        20%  | crf 51   | 20`n        10%  | crf 58   | 10`n`n    - For ICO and other formats: Magick handles these automatically.`n`n    Note: At equivalent quality levels, AVIF achieves higher compression efficiency than JXL and WEBP.`n`n`n【 4 】 RESIZING RULES`n`n    Max Resolution: The limit for the longest side of the image. For example, if the original is 7680x5120 and Max Resolution is set to 3840, the output will be 3840x2560 (maintaining aspect ratio).`n`n    Downscale Limit: This is a `"safety stop`" to prevent images from being shrunk too small, causing permanent loss of detail.`n        - Formula: Output size will not be smaller than (Original/Ratio).`n        - Example: Max Resolution set to 3840 but Downscale Limit is 1.5. For a 7680px original, the result will be 5120px (because 7680/1.5 = 5120, which is larger than your desired 3840).`n`n    Why limit the scale ratio? You might want to shrink images now to save storage but may wish to restore them to original size later. If an original image is large and highly detailed (e.g., a festival photo with hundreds of people), over-shrinking will permanently erase fine details. Based on experience scaling thousands of images, the author recommends shrinking by only 1.5 to 2 times the original size in most cases. However, if your image only contains large details, you can shrink it further.`n`n    If the Downscale Limit is set to 1, the output size will not be smaller than (Original Size)/1, which effectively keeps the original dimensions.`n`n`n【 5 】 TIPS AND ERROR REPORTS`n`n    You can drag and drop image files or folders containing images into the software; the list of valid images will be populated and ready for conversion.`n`n    If you encounter the error `"Unable to read the original image or write the converted image to the disk,`" it may be due to:`n        - The original image has restricted read permissions or is locked.`n        - No permission to save the converted image to the original disk/folder.`n        - The disk is full, preventing the converted image from being saved.`n`n`n【 6 】 ABOUT THE AUTHOR`n`n    QH Image Optimizer 1.0 was written by facebook.com/nqhaivn/, released on March 24, 2026.`n`n    Any act of modifying the source code without written consent from the author is considered theft.`n`n    For suggestions, bug reports, software customization requests, or other software design inquiries, please contact email: nqhai86@gmail.com`n`n    Sincerely!"
        GuideTitle = "User Guide"
        StatusName = "Name"
        StatusSize = "Dimensions"
        StatusBytes = "Size"
        LogConvertSuccess = "✔ Success"
        LogMagickMissing = "ImageMagick (magick.exe) not found. This program requires ImageMagick to function. Open the Guide for more details."
        LogNoValidTasks = "Rejected: ImageMagick does not have access or cannot process this format."
        LogEncodeFailed = "{0} ❌ Failed | Reason: {1}"
        LogOutputWriteFailed = "Unable to read the original image or write the converted image to the disk. Open the Guide for more details."
        LogIcoOnlyPngOrTiffSupported = "ICO can only be converted to PNG or TIFF"
        LogMultiFrameToIcoNotSupported = "Cannot convert multi-frame image to ICO"
        LogRetryAttempt = "Retry attempt {0}: {1}"
        LogRunspaceError = "An error occurred during execution. Please try again."
        LogAddedFilesWithErrors = "Added {0} images to the list. Of which {1} images had errors and dimensions could not be retrieved."
        LogAddedFiles = "Added {0} images to the list."
        MissingFileMsg = "The image file no longer exists! Reason: Renamed, moved, or deleted."
        Remove = "Remove from list"
        OpenFolder = "Open in folder"
        PickScope = "Do you want to add images only from the parent folder, or all images in subfolders?"
        EncTo = "Convert to"
        Quality = "Quality"
        MaxSize = "Max Resolution *"
        Minimize = "Downscale Limit *"
        Prefix = "Add Prefix"
        Suffix = "Add Suffix"
        SaveWhere = "Output Destination"
        KeepMeta = "Metadata"
        RunMode = "Run Mode"
        SaveSame = "Same as original"
        SaveNew = "In a new folder"
        MetaAll = "Keep"
        MetaNone = "Remove"
        Yes = "Yes"
        No = "No"
        Parallel = "Multi-threaded"
        Sequential = "Sequential"
        Original = "Keep original format"
        Q100 = "100%"
        Q95 = "95%"
        Q90 = "90%"
        Q85 = "85%"
        Q80 = "80%"
        Q75 = "75%"
        Q70 = "70%"
        Q65 = "65%"
        Q60 = "60%"
        Q50 = "50%"
        Q40 = "40%"
        Q30 = "30%"
        Q20 = "20%"
        Q10 = "10%"
        Lossless = "Original"
        IcoQualityNormal = "Optimized"
        IcoQualityBest = "Best"
        InvalidChar = "Characters {0} are not allowed because Windows does not permit them."
        LoadingList = "Loading image list..."
        TooltipEncTo = "Which format do you want to convert the original image to? `n`nNote: 'Original Format' preserves the source extension. However, for better compatibility, these groups will be standardized: `n- JPG variants (JFIF, JPE...) & HEIC/HEIF → .jpg `n- TIFF variants → .tiff `n- PNG, APNG, ICO, ICON & BMP → .png"
        TooltipQuality = "Select output image quality. `n`nNote: `n- When converting to PNG or TIFF, quality is always 100%. `n- When converting to ICO, 'Optimized' quality will reduce the image to 256 colors and remove the transparent background."
        TooltipMaxSize = "Sets the maximum length for the longest side of the output image.`n`nExample: If the original is 6000x4000 px and Max Resolution is 3840, the output will be 3840x2560 px. `n`nIf the original is already smaller than this value, it will remain unchanged."
        TooltipMinimize = "Detail Preservation Factor: Prevents the image from being downscaled too much relative to the original.`n`nFormula: (Original Size / This Factor). `n`nRecommendation: Set between 1.5 and 2 to ensure the image can be upscaled in the future without losing clarity. `n`nExample: For a 4500x3000 px original, a factor of 1.5 ensures the output is at least 3000x2000 px, regardless of the Max Resolution setting."
        TooltipPrefix = "Add a prefix to the output image. Formula: <Prefix><Original Name>. If the name is 'photo.jpg' and the prefix is 'New_', the output will be 'New_photo.jpg'"
        TooltipSuffix = "Add a suffix to the output image. Formula: <Original Name><Suffix>. If the name is 'photo.jpg' and the suffix is '_New', the output will be 'photo_New.jpg'"
        TooltipSaveWhere = "Choose where to save the processed images.`n`n⚠️ WARNING: Selecting 'Same as original' with the same format will OVERWRITE your original files (original data will be lost). Double-check your settings before batch processing.`n`n'In a new folder': The software creates a separate folder to protect your originals. The folder name will include a timestamp and the configuration settings.`nExample: 20260325_052511_max_3840_limit_1.5_JPG_quality_75"
        TooltipMeta = "Includes info on: creation/modification/capture time, rotation, brightness, device, GPS coordinates... `n`nRemoving metadata protects privacy when sharing, but you will lose useful capture time info. `n`nKeep metadata if these are documentary or archival photos."
        TooltipRun = "Multi-threading offers higher processing speed but uses more CPU resources. If your computer is low-spec or running many tasks, choose 'Sequential' for safety."
        AfterEncodeTitle = "Notification"
        AfterEncodeTemplate = "➤ Converted: Success {0} | Failed {1} | Original Size: {2} | Current Size: {3} ({4}) | Time: {5} min {6} sec."
        IcoTitle = "Select sizes to be embedded in the .ico file"
        ToolsMissingTitle = "ImageMagick Missing"
        ToolsMissingMsg = "magick.exe not found. Please open the Guide for more details."
    }
}

$script:currentLang = "vi"
$script:LangChanging = $false

# =====================================
# THIẾT LẬP LOGIC
# =====================================
#
# Tất cả biến cấu hình của logic sẽ được khởi tạo trong hàm
# Initialize-LogicFromSettings dựa trên tham số UI.

# =====================================
# HÀM HỖ TRỢ ĐƯỜNG DẪN / ĐỊNH DẠNG
# =====================================

# Nhóm hàm này tập trung vào chuẩn hóa đường dẫn, kích thước file và định dạng chuỗi.
# Chuyển đường dẫn sang dạng \\?\ để hỗ trợ path dài trên Windows.
function Convert-ToLongPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $full = [System.IO.Path]::GetFullPath($Path)
    if ($full -match '^\\\\\?\\') {
        return $full
    }

    # Nếu là UNC path thì phải dùng tiền tố \\?\UNC\ để đúng chuẩn long-path.
    if ($full -match '^\\\\') {
        return ("\\?\UNC\" + $full.TrimStart('\'))
    }

    return "\\?\$full"
}

# Lấy kích thước file theo byte.
function Get-FileSizeBytes {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-Item -LiteralPath $Path).Length
}

# Chuyển số sang chuỗi invariant để tránh lỗi dấu chấm thập phân.
function ConvertTo-InvariantString {
    param([Parameter(Mandatory = $true)][object]$Value)

    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        return ([double]$Value).ToString("0.################", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return [string]$Value
}

# Tìm đường dẫn exe: ưu tiên quét thư mục rồi mới PATH (khi bật PreferSearchRoots).
function Resolve-ExecutablePath {
    param(
        [Parameter(Mandatory = $true)][string]$ExeName,
        [Parameter(Mandatory = $true)][string[]]$SearchRoots,
        [switch]$PreferSearchRoots,
        [switch]$NoRecurse
    )

    $searchRootsFirst = $PreferSearchRoots.IsPresent

    $tryRoots = {
        foreach ($root in $SearchRoots) {
            if (-not (Test-Path -LiteralPath $root)) { continue }
            if ($NoRecurse) {
                $candidate = Join-Path $root $ExeName
                if (Test-Path -LiteralPath $candidate) {
                    $full = [System.IO.Path]::GetFullPath($candidate)
                    return $full
                }
            } else {
                $found = Get-ChildItem -Path $root -Filter $ExeName -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($found) {
                    $full = [System.IO.Path]::GetFullPath($found.FullName)
                    return $full
                }
            }
        }
        return $null
    }

    if ($searchRootsFirst) {
        $rootHit = & $tryRoots
        if ($rootHit) { return $rootHit }
    }

    # Ưu tiên tìm từ biến môi trường PATH.
    $cmd = Get-Command -Name $ExeName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd -and $cmd.Source) {
        return [System.IO.Path]::GetFullPath($cmd.Source)
    }

    if (-not $searchRootsFirst) {
        $rootHit = & $tryRoots
        if ($rootHit) { return $rootHit }
    }

    return $null
}


# =====================================
# HÀM HỖ TRỢ ENCODE / XỬ LÝ THAM SỐ
# =====================================

# Ghi log vào StatusLog (hoạt động cả UI lẫn runspace).
function Add-StatusLogLine {
    param([string]$Text)
    if ($script:StatusLog -and $script:StatusLog.Dispatcher.CheckAccess()) {
        $script:StatusLog.AppendText($Text + "`r`n")
        $script:StatusLog.ScrollToEnd()
        return
    }
    if ($script:EncodeLogQueue) {
        $script:EncodeLogQueue.Enqueue($Text) | Out-Null
        return
    }
    if (-not $script:StatusLog) { return }
    # Dùng closure để tránh lỗi overload BeginInvoke trong runspace.
    $action = {
        $tb = $script:StatusLog
        if (-not $tb) { return }
        $tb.AppendText($Text + "`r`n")
        $tb.ScrollToEnd()
    }
    $null = $script:StatusLog.Dispatcher.BeginInvoke([action]$action)
}

# Ghi lỗi encode ra StatusLog, có thể bật thêm console khi cần kiểm thử.
function Write-EncodeErrorLog {
    param([Parameter(Mandatory = $true)][string]$Message)

    if ($global:log_writer) {
        try { & $global:log_writer $Message } catch {}
    } else {
        Add-StatusLogLine -Text $Message
    }
}

# Build argument string an toàn cho Start-Process (quote đúng khi có dấu cách).
function ConvertTo-ProcessArgumentString {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    # Start-Process nhận argument dạng chuỗi; cần quote đúng để không vỡ tham số khi path có khoảng trắng.
    $quoted = foreach ($arg in $Arguments) {
        $text = [string]$arg

        if ($text.Length -eq 0) {
            '""'
            continue
        }

        if ($text -notmatch '[\s"]') {
            $text
            continue
        }

        # Quote theo quy tắc CommandLineToArgvW.
        $escaped = $text -replace '(\\*)"', '$1$1\"'
        $escaped = $escaped -replace '(\\+)$', '$1$1'
        '"' + $escaped + '"'
    }

    return ($quoted -join ' ')
}

# Tạo file ảnh tạm để tránh ghi đè trực tiếp khi input và output trùng đuôi.
function New-UniqueTempImagePath {
    param(
        [Parameter(Mandatory = $true)][string]$OutputDir,
        [Parameter(Mandatory = $true)][string]$BaseName,
        [Parameter(Mandatory = $true)][string]$Extension
    )

    do {
        $tempName = "{0}.__tmp__{1}.{2}" -f $BaseName, ([guid]::NewGuid().ToString("N")), $Extension
        $normalPath = Join-Path $OutputDir $tempName
    } while (Test-Path -LiteralPath $normalPath)

    return [PSCustomObject]@{
        NormalPath = $normalPath
        LongPath   = Convert-ToLongPath $normalPath
    }
}


# =====================================
# HÀM HỖ TRỢ ĐO TẢI HỆ THỐNG / CHẠY SONG SONG ĐỘNG
# =====================================

# Nhóm hàm này đo tải hệ thống và tính số worker phù hợp cho chế độ adaptive.
function Limit-Percent {
    param([double]$Value)
    if ($Value -lt 0) { return 0.0 }
    if ($Value -gt 100) { return 100.0 }
    return [double]$Value
}

# Lấy chỉ số tải CPU/RAM/IO hiện tại để điều chỉnh worker.
function Get-SystemLoadMetrics {
    [double]$cpu = 0
    [double]$io = 0
    [double]$ram = 0

    try {
        # Truy vấn PerfFormattedData để tránh Get-Counter -SampleInterval 1 (gây chặn ~1 giây mỗi vòng).
        $cpuObj = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'" -ErrorAction Stop
        if ($cpuObj) { $cpu = [double]$cpuObj.PercentProcessorTime }
    }
    catch {
        $cpu = 0
    }

    try {
        $disk = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'" -ErrorAction Stop
        if ($disk) { $io = [double]$disk.PercentDiskTime }
    }
    catch {
        $io = 0
    }

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $freeKb = [double]$os.FreePhysicalMemory
        $totalKb = [double]$os.TotalVisibleMemorySize
        if ($totalKb -gt 0) {
            $ram = (1 - ($freeKb / $totalKb)) * 100
        }
    }
    catch {
        $ram = 0
    }

    return [PSCustomObject]@{
        CPU = [math]::Round((Limit-Percent $cpu), 2)
        RAM = [math]::Round((Limit-Percent $ram), 2)
        IO  = [math]::Round((Limit-Percent $io), 2)
    }
}

# Ước tính số luồng CPU đang bận dựa trên %CPU.
function Get-EstimatedBusyWorkersFromCpu {
    param(
        [Parameter(Mandatory = $true)][double]$CpuPercent,
        [Parameter(Mandatory = $true)][int]$LogicalThreads
    )

    # Quy đổi %CPU hiện tại thành số luồng CPU đang bận (xấp xỉ).
    if ($LogicalThreads -le 0) { return 0 }

    $busy = [int][Math]::Round(
        (($CpuPercent / 100.0) * $LogicalThreads),
        0,
        [System.MidpointRounding]::AwayFromZero
    )

    return [Math]::Max(0, [Math]::Min($LogicalThreads, $busy))
}

# Lập kế hoạch số worker ban đầu cho chế độ adaptive.
function Get-AdaptiveStartupPlan {
    param(
        [Parameter(Mandatory = $true)][double]$CpuPercent,
        [Parameter(Mandatory = $true)][int]$LogicalThreads,
        [Parameter(Mandatory = $true)][double]$StartFreeRatio,
        [Parameter(Mandatory = $true)][int]$PendingCount
    )

    $busyWorkers = Get-EstimatedBusyWorkersFromCpu -CpuPercent $CpuPercent -LogicalThreads $LogicalThreads
    $freeWorkers = [Math]::Max(0, $LogicalThreads - $busyWorkers)
    $ratio = [Math]::Max(0.0, [Math]::Min(1.0, $StartFreeRatio))

    # Luật khởi động:
    # - workers bắt đầu = ratio * freeWorkers (làm tròn).
    # - Luôn đảm bảo tối thiểu 1 worker.
    $target = [int][Math]::Round($freeWorkers * $ratio, 0, [System.MidpointRounding]::AwayFromZero)
    $target = [Math]::Max(1, $target)

    if ($PendingCount -gt 0) {
        $target = [Math]::Min($target, $PendingCount)
    }

    return [PSCustomObject]@{
        BusyWorkers       = $busyWorkers
        FreeWorkers       = $freeWorkers
        RecommendedWorkers = $target
    }
}

# Chuẩn hóa lý do lỗi encode để log rõ ràng, không mơ hồ.
function Get-EncodeFailureDetail {
    param([Parameter(Mandatory = $true)][string]$Reason)

    if ([string]::IsNullOrWhiteSpace($Reason)) { return $strings[$script:currentLang].LogOutputWriteFailed }

    if ($Reason -match '^encode_exit_code_(\d+)$') {
        return $strings[$script:currentLang].LogOutputWriteFailed
    }

    switch ($Reason) {
        "output_missing" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "output_empty" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "output_probe_failed" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "start_process_failed" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "magick_missing" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "heic_magick_failed" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "heic_temp_missing" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "probe_failed" { return $strings[$script:currentLang].LogOutputWriteFailed }
        "unsupported_format" { return $strings[$script:currentLang].LogOutputWriteFailed }
        default { return $strings[$script:currentLang].LogOutputWriteFailed }
    }
}

# Lấy cấu hình chất lượng theo thang 100/95/90/85/80/75/70/65/60/50/40/30/20/10.
function Get-QualityPlan {
    param([Parameter(Mandatory = $true)][int]$EncodeQuality)

    switch ($EncodeQuality) {
        100 {
            return @{
                JpgQuality = 100
                WebpLossless = $true
                WebpQuality = 100
                JxlQuality = 100
                AvifMagickQuality = 100
                AvifencLossless = $true
                AvifencMin = $null
                AvifencMax = $null
            }
        }
        95 {
            return @{
                JpgQuality = 95
                WebpLossless = $false
                WebpQuality = 95
                JxlQuality = 95
                AvifMagickQuality = 95
                AvifencLossless = $false
                AvifencMin = 13
                AvifencMax = 13
            }
        }
        90 {
            return @{
                JpgQuality = 90
                WebpLossless = $false
                WebpQuality = 90
                JxlQuality = 90
                AvifMagickQuality = 90
                AvifencLossless = $false
                AvifencMin = 16
                AvifencMax = 16
            }
        }
        85 {
            return @{
                JpgQuality = 85
                WebpLossless = $false
                WebpQuality = 85
                JxlQuality = 85
                AvifMagickQuality = 80
                AvifencLossless = $false
                AvifencMin = 24
                AvifencMax = 24
            }
        }
        80 {
            return @{
                JpgQuality = 80
                WebpLossless = $false
                WebpQuality = 80
                JxlQuality = 80
                AvifMagickQuality = 70
                AvifencLossless = $false
                AvifencMin = 28
                AvifencMax = 28
            }
        }
        75 {
            return @{
                JpgQuality = 75
                WebpLossless = $false
                WebpQuality = 75
                JxlQuality = 75
                AvifMagickQuality = 60
                AvifencLossless = $false
                AvifencMin = 31
                AvifencMax = 31
            }
        }
        70 {
            return @{
                JpgQuality = 70
                WebpLossless = $false
                WebpQuality = 70
                JxlQuality = 70
                AvifMagickQuality = 50
                AvifencLossless = $false
                AvifencMin = 36
                AvifencMax = 36
            }
        }
        65 {
            return @{
                JpgQuality = 65
                WebpLossless = $false
                WebpQuality = 65
                JxlQuality = 65
                AvifMagickQuality = 48
                AvifencLossless = $false
                AvifencMin = 38
                AvifencMax = 38
            }
        }
        60 {
            return @{
                JpgQuality = 60
                WebpLossless = $false
                WebpQuality = 60
                JxlQuality = 60
                AvifMagickQuality = 44
                AvifencLossless = $false
                AvifencMin = 40
                AvifencMax = 40
            }
        }
        50 {
            return @{
                JpgQuality = 50
                WebpLossless = $false
                WebpQuality = 50
                JxlQuality = 50
                AvifMagickQuality = 40
                AvifencLossless = $false
                AvifencMin = 42
                AvifencMax = 42
            }
        }
        40 {
            return @{
                JpgQuality = 40
                WebpLossless = $false
                WebpQuality = 40
                JxlQuality = 40
                AvifMagickQuality = 36
                AvifencLossless = $false
                AvifencMin = 44
                AvifencMax = 44
            }
        }
        30 {
            return @{
                JpgQuality = 30
                WebpLossless = $false
                WebpQuality = 30
                JxlQuality = 30
                AvifMagickQuality = 30
                AvifencLossless = $false
                AvifencMin = 47
                AvifencMax = 47
            }
        }
        20 {
            return @{
                JpgQuality = 20
                WebpLossless = $false
                WebpQuality = 20
                JxlQuality = 20
                AvifMagickQuality = 20
                AvifencLossless = $false
                AvifencMin = 51
                AvifencMax = 51
            }
        }
        10 {
            return @{
                JpgQuality = 10
                WebpLossless = $false
                WebpQuality = 10
                JxlQuality = 10
                AvifMagickQuality = 10
                AvifencLossless = $false
                AvifencMin = 58
                AvifencMax = 58
            }
        }
        default {
            throw "Giá trị chất lượng không hợp lệ: $EncodeQuality"
        }
    }
}

# Tạo tham số magick theo định dạng và chất lượng.
function New-MagickArgsForFormat {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$FormatKey,
        [Parameter(Mandatory = $true)][string]$InputExt,
        [Parameter(Mandatory = $true)][hashtable]$QualityPlan,
        [Parameter()][string]$ResizeArg,
        [Parameter(Mandatory = $true)][bool]$StripMetadata,
        [Parameter()][bool]$IsAnimated = $false
    )

    $magickArgs = @($InputPath)
    $isAnimatedGif = $IsAnimated -and ($FormatKey -eq "gif")
    $isAnimatedWebp = $IsAnimated -and ($FormatKey -eq "webp")
    $useAnimationPipeline = $isAnimatedGif -or $isAnimatedWebp
    $needsCoalesce = $useAnimationPipeline -or ($IsAnimated -and ($FormatKey -eq "jxl") -and ($InputExt -in @("avif", "webp", "gif", "heic", "heif", "tif", "tiff")))
    if ($needsCoalesce) { $magickArgs += "-coalesce" }
    if ($isAnimatedWebp) { $magickArgs += @("-alpha", "on", "-background", "none") }
    $magickArgs += "-auto-orient"
    if ($StripMetadata) { $magickArgs += "-strip" }
    if ($ResizeArg) { $magickArgs += @("-resize", $ResizeArg) }

    switch ($FormatKey) {
        "jpg" {
            if ($null -ne $QualityPlan.JpgQuality) {
                $magickArgs += @("-quality", $QualityPlan.JpgQuality)
            }
        }
        "png" {
            if ($null -ne $QualityPlan.JpgQuality) {
                $magickArgs += @("-quality", "95")
            }
        }
        "tiff" {
            $magickArgs += @("-compress", "Zip")
        }
        "webp" {
            $magickArgs += @("-define", "webp:method=6")
            if ($QualityPlan.WebpLossless) {
                $magickArgs += @("-define", "webp:lossless=true")
            }
            else {
                if ($null -ne $QualityPlan.WebpQuality) {
                    $magickArgs += @("-quality", $QualityPlan.WebpQuality)
                }
            }
        }
        "jxl" {
            $magickArgs += @("-define", "jxl:effort=7")
            if ($null -ne $QualityPlan.JxlQuality) {
                $magickArgs += @("-quality", $QualityPlan.JxlQuality)
            }
        }
        "avif" {
            $magickArgs += @("-define", "avif:speed=5")
            if ($null -ne $QualityPlan.AvifMagickQuality) {
                $magickArgs += @("-quality", $QualityPlan.AvifMagickQuality)
            }
        }
        default {
            # Định dạng khác: chỉ resize/strip nếu có.
        }
    }

    if ($isAnimatedGif) {
        $magickArgs += @("-colors", "256", "-dither", "FloydSteinberg", "-layers", "OptimizeTransparency")
    }
    elseif ($isAnimatedWebp) {
        $magickArgs += @("-layers", "OptimizePlus")
    }
    elseif ($useAnimationPipeline) {
        $magickArgs += @("-layers", "Optimize")
    }
    $magickArgs += $OutputPath
    return $magickArgs
}

# Tạo tham số avifenc theo chất lượng.
function New-AvifencArgs {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][hashtable]$QualityPlan,
        [Parameter(Mandatory = $true)][bool]$StripMetadata
    )

    $avifArgs = @("--speed", "5", "--jobs", "all")

    if ($QualityPlan.AvifencLossless) {
        $avifArgs += "--lossless"
    }
    else {
        if ($null -ne $QualityPlan.AvifencMin -and $null -ne $QualityPlan.AvifencMax) {
            $avifArgs += @("--min", [string]$QualityPlan.AvifencMin, "--max", [string]$QualityPlan.AvifencMax)
        }
    }

    if ($StripMetadata) {
        $avifArgs += @("--ignore-exif", "--ignore-xmp", "--ignore-icc")
    }

    $avifArgs += $InputPath
    $avifArgs += $OutputPath
    return $avifArgs
}

# Chuẩn bị file PNG tạm khi cần tiền xử lý bằng magick. Trả về kết quả + lý do nếu thất bại.
function Invoke-HeicPreprocess {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Task,
        [Parameter(Mandatory = $true)][string]$MagickPath
    )

    if (-not $Task.HeicStep1Args) {
        return [PSCustomObject]@{ Success = $true; Reason = $null }
    }
    if (-not $MagickPath) {
        return [PSCustomObject]@{ Success = $false; Reason = "magick_missing" }
    }

    try {
        & $MagickPath @($Task.HeicStep1Args) 2>$null
        $exitCode = $LASTEXITCODE
    }
    catch {
        $exitCode = 1
    }

    if ($exitCode -ne 0) {
        return [PSCustomObject]@{ Success = $false; Reason = "heic_magick_failed" }
    }

    if (-not (Test-Path -LiteralPath $Task.HeicTempLong)) {
        return [PSCustomObject]@{ Success = $false; Reason = "heic_temp_missing" }
    }

    return [PSCustomObject]@{ Success = $true; Reason = $null }
}

# Dọn file tạm heic/heif nếu có.
function Clear-HeicTempIfExists {
    param([Parameter(Mandatory = $true)][pscustomobject]$Task)
    if ($Task.HeicTempLong -and (Test-Path -LiteralPath $Task.HeicTempLong)) {
        Remove-Item -LiteralPath $Task.HeicTempLong -Force -ErrorAction SilentlyContinue
    }
}

# Ghi nhận trạng thái để UI đọc lại sau khi chạy.
function Add-TaskReport {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Task,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter()][string]$Reason
    )

    $script:TaskReport.Add([PSCustomObject]@{
            InputPath  = $Task.InputPathNormal
            OutputPath = $Task.FinalOutputNormal
            OutputFramesGlob = if ($Task.FrameRenameGlob) { $Task.FrameRenameGlob } else { $Task.FrameOutputGlob }
            OutputFramesGlobAlt = $Task.FrameOutputGlob
            Status     = $Status
            Reason     = $Reason
        }) | Out-Null

    # Đẩy tiến độ về UI theo đường dẫn đầy đủ và theo ngôn ngữ.
    if ($global:log_writer -and $Status -eq "success" -and -not [string]::IsNullOrWhiteSpace($Task.InputPathNormal)) {
        $label = $strings[$script:currentLang].LogConvertSuccess
        $progressMessage = ("{0} {1}" -f $Task.InputPathNormal, $label)
        try { & $global:log_writer $progressMessage } catch {}
    }
}

# Lấy kích thước ảnh bằng magick identify; ưu tiên -ping để tăng tốc.
function Get-ImageDimensionsFromMagick {
    param(
        [Parameter(Mandatory = $true)][string]$FilePathLong,
        [bool]$UsePing = $true
    )

    if (-not $magickPath) { return $null }
    try {
        $argsm = @("identify")
        if ($UsePing) { $argsm += "-ping" }
        $argsm += @("-format", "%wx%h", $FilePathLong)
        $probe = & $magickPath @argsm 2>$null
        if (-not $probe) { return $null }
        $text = ($probe -join "`n") -split "`r?`n" | Select-Object -First 1
        $match = [regex]::Match($text, '(\d+)\D+(\d+)')
        if ($match.Success) {
            return [PSCustomObject]@{
                Width  = [int]$match.Groups[1].Value
                Height = [int]$match.Groups[2].Value
            }
        }
        return $null
    }
    catch { return $null }
}

# Lấy kích thước ảnh bằng magick: thử -ping trước, thất bại thì bỏ -ping.
function Get-ImageDimensions {
    param([Parameter(Mandatory = $true)][string]$FilePathLong)

    $probe = Get-ImageDimensionsFromMagick -FilePathLong $FilePathLong -UsePing $true
    if ($probe) { return $probe }
    return Get-ImageDimensionsFromMagick -FilePathLong $FilePathLong -UsePing $false
}

# Lấy số frame tổng quát cho các định dạng đa khung hình khác.
function Get-FrameCountGeneric {
    param(
        [Parameter(Mandatory = $true)][string]$FilePathLong,
        [Parameter(Mandatory = $true)][string]$MagickPath
    )

    if (-not $MagickPath) { return $null }
    try {
        $out = & $MagickPath identify -ping -format "%n`n" $FilePathLong 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
        if (-not $out) { return $null }
        $nums = New-Object System.Collections.Generic.List[int]
        foreach ($line in (($out -join "`n") -split "`r?`n")) {
            $text = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $val = 0
            if ([int]::TryParse($text, [ref]$val)) {
                $nums.Add($val) | Out-Null
            }
        }
        if ($nums.Count -eq 0) { return $null }
        return ($nums | Measure-Object -Maximum).Maximum
    }
    catch { return $null }
}

# Ghi kích thước vào map và tùy chọn ưu tiên kích thước lớn hơn.
function Set-DimensionMapEntry {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Map,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][int]$Width,
        [Parameter(Mandatory = $true)][int]$Height,
        [switch]$PreferLarger
    )

    $mapKey = ([string]$Key).ToLowerInvariant()
    if (-not $Map.ContainsKey($mapKey)) {
        $Map[$mapKey] = [PSCustomObject]@{ Width = $Width; Height = $Height }
        return
    }

    if ($PreferLarger) {
        $old = $Map[$mapKey]
        $oldArea = [long]$old.Width * [long]$old.Height
        $newArea = [long]$Width * [long]$Height
        if ($newArea -gt $oldArea) {
            $Map[$mapKey] = [PSCustomObject]@{ Width = $Width; Height = $Height }
        }
    }
    else {
        $Map[$mapKey] = [PSCustomObject]@{ Width = $Width; Height = $Height }
    }
}

# Xác định đuôi output thực tế theo chế độ converter (original chuẩn hóa theo họ đuôi).
function Resolve-OutputExtension {
    param(
        [Parameter(Mandatory = $true)][string]$InputExt,
        [Parameter(Mandatory = $true)][string]$Converter
    )

    $ext = $InputExt.ToLowerInvariant()
    if ($Converter -eq "original") {
        if ($script:jpg_family_exts -contains $ext) { return "jpg" }
        if ($script:tiff_family_exts -contains $ext) { return "tiff" }
        if ($script:png_family_exts -contains $ext) { return "png" }
        return $ext
    }
    return $Converter
}


# Đọc kích thước ảnh bằng magick theo batch và đa luồng; ưu tiên -ping để tăng tốc.
function Get-ImageDimensionsMagickBatch {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory = $true)][int]$ThreadCount,
        [Parameter(Mandatory = $true)][string]$MagickPath,
        [bool]$UsePing = $true
    )

    $result = @{}
    if (-not $Files -or $Files.Count -eq 0) { return $result }
    if (-not $MagickPath) { return $result }

    $threads = [Math]::Max(1, $ThreadCount)
    $totalFiles = $Files.Count

    # Tính batch theo file dài nhất để giới hạn ~32k ký tự trên dòng lệnh.
    $longestPath = 0
    foreach ($f in $Files) {
        $len = $f.FullName.Length
        if ($len -gt $longestPath) { $longestPath = $len }
    }
    $safeLength = [Math]::Max(1, $longestPath)
    $batchSize = [Math]::Floor(32000 / $safeLength)
    $batchSize = [Math]::Max(1, $batchSize)

    $batches = New-Object System.Collections.Generic.List[object[]]
    for ($i = 0; $i -lt $totalFiles; $i += $batchSize) {
        $end = [Math]::Min($i + $batchSize - 1, $totalFiles - 1)
        $batches.Add($Files[$i..$end])
    }

    $pool = [RunspaceFactory]::CreateRunspacePool(1, $threads)
    $pool.Open()
    $jobs = New-Object System.Collections.Generic.List[object]

    foreach ($batch in $batches) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        $null = $ps.AddScript({
            param($magickExe, $batchFiles, $usePing)
            $results = New-Object System.Collections.Generic.List[object]
            try {
                if ($usePing) {
                    $infoLines = & $magickExe "identify" "-ping" "-format" "%i|%w|%h`n" ($batchFiles.FullName) 2>$null
                } else {
                    $infoLines = & $magickExe "identify" "-format" "%i|%w|%h`n" ($batchFiles.FullName) 2>$null
                }
                foreach ($line in ($infoLines -split "`r?`n")) {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    $parts = $line.Split('|')
                    if ($parts.Count -lt 3) { continue }
                    # Gom multi-frame (tiff/ico/webp động) về cùng đường dẫn gốc bằng cách bỏ hậu tố [n].
                    $pathRaw = $parts[0].Trim()
                    if ([string]::IsNullOrWhiteSpace($pathRaw)) { continue }
                    $basePath = $pathRaw -replace '\[\d+\]$', ''
                    $results.Add([PSCustomObject]@{
                        Path   = $basePath
                        Width  = [int]$parts[1].Trim()
                        Height = [int]$parts[2].Trim()
                    }) | Out-Null
                }
            }
            catch {
                # Bỏ qua batch lỗi.
            }
            return $results
        }).AddArgument($MagickPath).AddArgument($batch).AddArgument($UsePing)

        $jobs.Add([PSCustomObject]@{
            Pipe = $ps
            Handle = $ps.BeginInvoke()
        }) | Out-Null
    }

    foreach ($job in $jobs) {
        $items = $job.Pipe.EndInvoke($job.Handle)
        foreach ($item in $items) {
            if (-not $item) { continue }
            Set-DimensionMapEntry -Map $result -Key $item.Path -Width ([int]$item.Width) -Height ([int]$item.Height) -PreferLarger
        }
        $job.Pipe.Dispose()
    }

    $pool.Close()
    $pool.Dispose()

    return $result
}

# Lấy kích thước ảnh bằng magick cho toàn bộ file:
# - Chạy batch với -ping trước để nhanh.
# - Những file không có kết quả sẽ được chạy lại không -ping.
# - File vẫn lỗi sẽ được trả về qua danh sách lỗi để báo statuslog.
function Get-ImageDimensionsMagickAllMap {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory = $true)][string]$MagickPath,
        [Parameter(Mandatory = $true)][int]$MagickThreadCount,
        [hashtable]$InputSizeMap,
        [ref]$FailedKeys
    )

    $dimensionMap = @{}
    if (-not $Files -or $Files.Count -eq 0) { return $dimensionMap }

    $pending = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($file in $Files) {
        $full = [System.IO.Path]::GetFullPath($file.FullName)
        $key = $full.ToLowerInvariant()

        # Nếu đã có kích thước từ listview/nguồn khác thì bỏ qua đo lại.
        if ($InputSizeMap -and $InputSizeMap.ContainsKey($key)) {
            $dimensionMap[$key] = $InputSizeMap[$key]
            continue
        }
        $pending.Add($file) | Out-Null
    }

    $failSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (-not $MagickPath) {
        foreach ($file in $pending) {
            $failSet.Add([System.IO.Path]::GetFullPath($file.FullName)) | Out-Null
        }
        if ($FailedKeys) { $FailedKeys.Value = $failSet }
        return $dimensionMap
    }

    if ($pending.Count -gt 0) {
        # Pass 1: -ping
        $pingMap = Get-ImageDimensionsMagickBatch -Files $pending -ThreadCount $MagickThreadCount -MagickPath $MagickPath -UsePing $true
        foreach ($k in $pingMap.Keys) {
            $dim = $pingMap[$k]
            Set-DimensionMapEntry -Map $dimensionMap -Key $k -Width ([int]$dim.Width) -Height ([int]$dim.Height)
        }

        # Gom file chưa có kết quả để chạy lại không -ping.
        $retry = New-Object System.Collections.Generic.List[System.IO.FileInfo]
        foreach ($file in $pending) {
            $full = [System.IO.Path]::GetFullPath($file.FullName)
            $key = $full.ToLowerInvariant()
            if (-not $dimensionMap.ContainsKey($key)) {
                $retry.Add($file) | Out-Null
            }
        }

        if ($retry.Count -gt 0) {
            # Pass 2: không -ping cho các file lỗi.
            $retryMap = Get-ImageDimensionsMagickBatch -Files $retry -ThreadCount $MagickThreadCount -MagickPath $MagickPath -UsePing $false
            foreach ($k in $retryMap.Keys) {
                $dim = $retryMap[$k]
                Set-DimensionMapEntry -Map $dimensionMap -Key $k -Width ([int]$dim.Width) -Height ([int]$dim.Height)
            }
        }

        # Những file vẫn không có kết quả sẽ được đo lại từng file để tránh batch bị ảnh hưởng bởi file lỗi.
        foreach ($file in $pending) {
            $full = [System.IO.Path]::GetFullPath($file.FullName)
            $key = $full.ToLowerInvariant()
            if (-not $dimensionMap.ContainsKey($key)) {
                $singleProbe = Get-ImageDimensions -FilePathLong $full
                if ($singleProbe) {
                    $dimensionMap[$key] = $singleProbe
                } else {
                    $failSet.Add($full) | Out-Null
                }
            }
        }
    }

    if ($FailedKeys) { $FailedKeys.Value = $failSet }
    return $dimensionMap
}

# Tính toán resize nếu cần và trả về kế hoạch resize.
function Get-ResizePlan {
    param(
        [Parameter(Mandatory = $true)][int]$OldSize,
        [Parameter(Mandatory = $true)][int]$MaxSize,
        [Parameter(Mandatory = $true)][double]$Minimize
    )

    $didResize = $false
    $newSize = $null
    $resizeArg = $null

    if ($Minimize -gt 1 -and $OldSize -gt $MaxSize) {
        $candidate = $OldSize / $Minimize
        if ($MaxSize -lt $candidate) {
            $newSize = [int][Math]::Round($candidate)
        }
        else {
            $newSize = $MaxSize
        }
        $resizeArg = "${newSize}x${newSize}>"
        $didResize = $true
    }

    return [PSCustomObject]@{
        DidResize    = $didResize
        NewSize      = $newSize
        ResizeArg    = $resizeArg
    }
}

# Lấy danh sách kích thước ICO theo cấu hình bật/tắt.
function Get-IcoSizes {
    $sizes = @()
    if ($ico16)  { $sizes += 16 }
    if ($ico32)  { $sizes += 32 }
    if ($ico48)  { $sizes += 48 }
    if ($ico64)  { $sizes += 64 }
    if ($ico128) { $sizes += 128 }
    if ($ico256) { $sizes += 256 }

    if ($sizes.Count -eq 0) {
        $sizes = @(256)
    }

    return ($sizes | Sort-Object -Unique)
}

function Initialize-LogicFromSettings {
    param([Parameter(Mandatory = $true)][hashtable]$Settings)

    # Kiểm tra tối thiểu các khóa bắt buộc từ UI.
    $requiredKeys = @(
        "converter",
        "encode_quality",
        "max_size",
        "minimize",
        "prefix",
        "suffix",
        "create_folder",
        "metadata_keep",
        "enable_adaptive_parallel",
        "ico_option",
        "ico16",
        "ico32",
        "ico48",
        "ico64",
        "ico128",
        "ico256"
    )
    foreach ($key in $requiredKeys) {
        if (-not $Settings.ContainsKey($key)) {
            throw "Thiếu cấu hình bắt buộc từ UI: $key"
        }
    }

    # Nhận cấu hình trực tiếp từ UI, không dùng mặc định toàn cục.
    $script:converter = [string]$Settings.converter
    $script:encode_quality = [int]$Settings.encode_quality
    $script:max_size = [int]$Settings.max_size
    $script:minimize = [double]$Settings.minimize
    $script:prefix = [string]$Settings.prefix
    $script:suffix = [string]$Settings.suffix
    $script:create_folder = [bool]$Settings.create_folder
    $script:metadata_keep = [bool]$Settings.metadata_keep
    $script:enable_adaptive_parallel = [bool]$Settings.enable_adaptive_parallel
    $script:ico_option = [string]$Settings.ico_option
    if ($Settings.ContainsKey("input_size_map")) {
        $script:input_size_map = $Settings.input_size_map
    } else {
        $script:input_size_map = $null
    }

    # Cấu hình ICO từ UI.
    $script:ico16  = [bool]$Settings.ico16
    $script:ico32  = [bool]$Settings.ico32
    $script:ico48  = [bool]$Settings.ico48
    $script:ico64  = [bool]$Settings.ico64
    $script:ico128 = [bool]$Settings.ico128
    $script:ico256 = [bool]$Settings.ico256

    # Các thông số nội bộ của logic (không lấy từ UI).
    $script:low_threshold = 60
    $script:high_threshold = 75
    $script:monitor_interval_seconds = 0.5
    $script:adaptive_poll_milliseconds = 100
    $script:adaptive_start_free_ratio = 0.4
    $script:adaptive_min_batch_threads_5_9 = 15
    $script:adaptive_scaleup_step = 1
    $script:adaptive_scaledown_step = 1
    $script:retry_count = 3

    # Họ đuôi file để chuẩn hóa khi converter = original.
    $script:jpg_family_exts = @("jpg", "jpeg", "jpe", "jfif", "jif", "pjpeg", "pjp", "heic", "heif")
    $script:tiff_family_exts = @("tif", "tiff")
    $script:png_family_exts = @("png", "bmp", "apng", "ico", "icon")
}

function Invoke-EncodeLogic {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][string[]]$InputFiles
    )

    Initialize-LogicFromSettings -Settings $Settings
    $InformationPreference = 'Continue'
    $script:InformationPreference = 'Continue'
    $logicalThreads = [Math]::Max(1, [Environment]::ProcessorCount)

# =====================================
# KIỂM TRA CÔNG CỤ (YÊU CẦU A)
# =====================================

$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$currentDirectory = (Get-Location).Path
$searchRoots = @($scriptDirectory, $currentDirectory) | Where-Object { $_ } | Select-Object -Unique

$magickPath = Resolve-ExecutablePath -ExeName "magick.exe" -SearchRoots $searchRoots -PreferSearchRoots

# Tìm avifenc.exe trong thư mục script và thư mục con.
$avifencPath = $null
$avifencFile = Get-ChildItem -Path $scriptDirectory -Filter "avifenc.exe" -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
if ($avifencFile) {
    $avifencPath = [System.IO.Path]::GetFullPath($avifencFile.FullName)
    $script:avifenc_available = $true
}
else {
    $script:avifenc_available = $false
}
$script:avifencPath = $avifencPath

if (-not $magickPath) {
    Write-EncodeErrorLog -Message $strings[$script:currentLang].LogMagickMissing
    exit 1
}


# =====================================
# TẠO THƯ MỤC ĐẦU RA (YÊU CẦU B)
# =====================================

$converter = $converter.ToLower()
$ico_option = $ico_option.ToLower()
# Những định dạng đầu ra được chấp nhận 
$supportedConverters = @("jpg", "png", "webp", "jxl", "avif", "tiff", "ico", "original")

$qualityPlan = Get-QualityPlan -EncodeQuality $encode_quality
$timestampLabel = (Get-Date).ToString("yyyyMMdd_HHmmss")
$converterLabel = if ($converter -eq "original") { "Original" } else { $converter.ToUpperInvariant() }
$qualityLabel = "quality_{0}" -f (ConvertTo-InvariantString $encode_quality)
if ($converter -eq "original") {
    $outputFolderNameBase = "{0}_Original_{1}" -f $timestampLabel, $qualityLabel
}
elseif ($minimize -gt 1) {
    $outputFolderNameBase = "{0}_max_{1}_limit_{2}_{3}_{4}" -f $timestampLabel, $max_size, (ConvertTo-InvariantString $minimize), $converterLabel, $qualityLabel
}
else {
    $outputFolderNameBase = "{0}_{1}_{2}" -f $timestampLabel, $converterLabel, $qualityLabel
}


# =====================================
# TẠO DANH SÁCH TÁC VỤ
# =====================================

$files = @()
# Chỉ chấp nhận danh sách file từ UI, không tự quét thư mục để tránh chạy ngoài ý muốn.
if (-not $InputFiles -or $InputFiles.Count -eq 0) {
    exit 1
}

$files = $InputFiles | ForEach-Object {
    try { Get-Item -LiteralPath $_ -ErrorAction Stop } catch { $null }
} | Where-Object {
    # Chỉ dùng một danh sách chuẩn ($script:supportedExtensions) để lọc input hợp lệ.
    $_ -and -not $_.PSIsContainer -and ($script:supportedExtensions -contains $_.Extension.ToLower())
}
if (-not $files -or $files.Count -eq 0) {
    Write-EncodeErrorLog -Message $strings[$script:currentLang].MissingFileMsg
    exit 0
}

$magickThreads = [Math]::Max(1, $logicalThreads)

# Đo kích thước toàn bộ bằng magick (ping trước, fallback không ping).
$failedKeys = $null
$dimensionMap = Get-ImageDimensionsMagickAllMap `
    -Files $files `
    -MagickPath $magickPath `
    -MagickThreadCount $magickThreads `
    -InputSizeMap $script:input_size_map `
    -FailedKeys ([ref]$failedKeys)

$baseOutputCounts = @{}
foreach ($f in $files) {
    if (-not $f) { continue }
    $ext = $f.Extension.TrimStart('.').ToLowerInvariant()
    $outExt = Resolve-OutputExtension -InputExt $ext -Converter $converter
    $base = $f.BaseName.ToLowerInvariant()
    $key = "$base|$outExt"
    if ($baseOutputCounts.ContainsKey($key)) {
        $baseOutputCounts[$key]++
    }
    else {
        $baseOutputCounts[$key] = 1
    }
}

$seenInputs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$createdOutputDirs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$tasks = New-Object System.Collections.Generic.List[object]
$script:TaskReport = New-Object System.Collections.Generic.List[object]

foreach ($file in $files) {
    $inputNormal = [System.IO.Path]::GetFullPath($file.FullName)
    if (-not $seenInputs.Add($inputNormal)) {
        continue
    }

    $inputLong = Convert-ToLongPath $inputNormal
    $inputLongOriginal = $inputLong

    # Không gọi magick theo từng file; nếu thiếu kích thước trong batch thì coi là lỗi.
    $probeInfo = $null
    $key = $inputNormal.ToLowerInvariant()
    if ($dimensionMap.ContainsKey($key)) {
        $probeInfo = $dimensionMap[$key]
    }
    if (-not $probeInfo) {
        Add-TaskReport -Task ([PSCustomObject]@{
                InputPathNormal = $inputNormal
                FinalOutputNormal = $inputNormal
            }) -Status "failed" -Reason "probe_failed"
        continue
    }

    $width = $probeInfo.Width
    $height = $probeInfo.Height
    $oldSize = [Math]::Max($width, $height)

    # Xác định định dạng đích thực tế cho từng file (original chuẩn hóa theo họ đuôi).
    $inputExt = $file.Extension.TrimStart('.').ToLower()
    $outputExt = Resolve-OutputExtension -InputExt $inputExt -Converter $converter

    $skipOutputValidation = $false
    $isAnimated = $false
    $isMultiFrame = $false
    $frameCount = Get-FrameCountGeneric -FilePathLong $inputLongOriginal -MagickPath $magickPath
    if ($frameCount -gt 1) {
        $isMultiFrame = $true
        $isAnimated = $true
        if ($inputExt -in @("gif", "webp")) {
            $skipOutputValidation = $true
        }
    }

    $formatKey = $outputExt
    if ($jpg_family_exts -contains $formatKey) {
        $formatKey = "jpg"
    }
    elseif ($tiff_family_exts -contains $formatKey) {
        $formatKey = "tiff"
    }

    if ($converter -ne "original" -and -not ($supportedConverters -contains $formatKey)) {
        Add-TaskReport -Task ([PSCustomObject]@{
                InputPathNormal = $inputNormal
                FinalOutputNormal = $inputNormal
            }) -Status "failed" -Reason "unsupported_format"
        continue
    }

    $skipReason = $null
    if (($inputExt -in @("ico", "icon")) -and -not ($converter -in @("png", "tiff", "original"))) {
        $skipReason = $strings[$script:currentLang].LogIcoOnlyPngOrTiffSupported
    }
    elseif ($converter -eq "ico" -and $isAnimated) {
        $skipReason = $strings[$script:currentLang].LogMultiFrameToIcoNotSupported
    }

    if ($skipReason) {
        $inputBytes = Get-FileSizeBytes $inputLongOriginal
        $tasks.Add([PSCustomObject]@{
                Name               = $file.Name
                InputPathNormal    = $inputNormal
                InputPathLong      = $inputLongOriginal
                InputSizeBytes     = $inputBytes
                FinalOutputNormal  = $inputNormal
                FinalOutputLong    = $inputLongOriginal
                SkipReason         = $skipReason
                RetryCount         = 0
            })
        continue
    }

    $resizePlan = Get-ResizePlan -OldSize $oldSize -MaxSize $max_size -Minimize $minimize
    $stripMetadata = -not $metadata_keep

    $inputDir = Split-Path -Parent $inputNormal
    if ($create_folder) {
        $outputDir = Join-Path $inputDir $outputFolderNameBase
        if (-not (Test-Path -LiteralPath $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir | Out-Null
        }
        $createdOutputDirs.Add($outputDir) | Out-Null
    }
    else {
        $outputDir = $inputDir
    }

    $dupKey = ("{0}|{1}" -f $file.BaseName.ToLowerInvariant(), $outputExt)
    $useFullName = $false
    if ($baseOutputCounts.ContainsKey($dupKey) -and $baseOutputCounts[$dupKey] -gt 1) {
        $useFullName = $true
    }
    $baseNameForOutput = if ($useFullName) { $file.Name } else { $file.BaseName }
    $outputBaseName = "${prefix}${baseNameForOutput}${suffix}"
    $finalName = "$outputBaseName.$outputExt"
    $finalOutputNormal = Join-Path $outputDir $finalName
    $finalOutputLong = Convert-ToLongPath $finalOutputNormal

    $canOverwrite = (-not $create_folder) -and [string]::IsNullOrEmpty($prefix) -and [string]::IsNullOrEmpty($suffix)
    $useTmpOverwrite = $false
    $shouldCheckOutputSize = $false
    $deleteInputAfterSuccess = $false
    $renameInputOnOutputTooLarge = $false
    $frameSuffixToken = "__frame__number__"
    $useFrameSuffix = $isMultiFrame -and ($inputExt -in @("ico", "icon", "gif", "webp", "jxl", "avif", "heic", "heif", "tif", "tiff")) -and ($formatKey -in @("jpg", "png"))
    $frameOutputGlob = $null
    $frameRenameGlob = $null

    if ($canOverwrite -and ($inputExt -eq $outputExt)) {
        # Trường hợp A: output trùng đuôi input -> dùng file tạm.
        $useTmpOverwrite = $true
        $shouldCheckOutputSize = $true
    }
    elseif ($canOverwrite -and ($converter -eq "original") -and ($jpg_family_exts -contains $inputExt) -and ($formatKey -eq "jpg")) {
        # Trường hợp B: họ JPG -> output .jpg, không dùng file tạm.
        $shouldCheckOutputSize = $true
        $deleteInputAfterSuccess = $true
        $renameInputOnOutputTooLarge = $true
    }
    elseif ($canOverwrite -and ($converter -eq "original") -and ($tiff_family_exts -contains $inputExt) -and ($formatKey -eq "tiff")) {
        # Trường hợp C: họ TIFF -> output .tiff, không dùng file tạm.
        $shouldCheckOutputSize = $true
        $deleteInputAfterSuccess = $true
        $renameInputOnOutputTooLarge = $true
    }
    $encodeTargetNormal = $finalOutputNormal
    $encodeTargetLong = $finalOutputLong

    if ($useTmpOverwrite) {
        $tempPathInfo = New-UniqueTempImagePath -OutputDir $outputDir -BaseName $outputBaseName -Extension $outputExt
        $encodeTargetNormal = $tempPathInfo.NormalPath
        $encodeTargetLong = $tempPathInfo.LongPath
    }
    elseif ($useFrameSuffix) {
        $frameOutputName = "$outputBaseName$frameSuffixToken%d.$outputExt"
        $encodeTargetNormal = Join-Path $outputDir $frameOutputName
        $encodeTargetLong = Convert-ToLongPath $encodeTargetNormal
        $frameOutputGlob = Join-Path $outputDir ("$outputBaseName$frameSuffixToken*.$outputExt")
        $frameRenameGlob = Join-Path $outputDir ("$outputBaseName-*.$outputExt")
        $skipOutputValidation = $true
        $shouldCheckOutputSize = $false
    }

    $runnerExe = $null
    $runnerArgs = $null
    $heicTempLong = $null
    $heicTempNormal = $null
    $heicStep1Args = $null

    if ($formatKey -eq "ico") {
        $icoSizes = Get-IcoSizes

        if ($icoSizes.Count -eq 1) {
            $size = $icoSizes[0]
            $resizeArg = "${size}x${size}!"

            if ($ico_option -eq "best") {
                $runnerExe = $magickPath
                $runnerArgs = @($inputNormal, "-auto-orient")
                if ($stripMetadata) { $runnerArgs += "-strip" }
                $runnerArgs += @("-resize", $resizeArg, "-background", "none", $encodeTargetNormal)
            }
            else {
                $runnerExe = $magickPath
                $runnerArgs = @(
                    $inputNormal,
                    "-auto-orient",
                    "-resize", $resizeArg,
                    "-background", "white",
                    "-alpha", "remove",
                    "-quantize", "transparent",
                    "+dither",
                    "-colors", "256"
                )
                if ($stripMetadata) { $runnerArgs += "-strip" }
                $runnerArgs += $encodeTargetNormal
            }
        }
        else {
            $maxSize = ($icoSizes | Measure-Object -Maximum).Maximum
            $allSizes = ($icoSizes | Sort-Object -Descending) -join ","
            $resizeArg = "${maxSize}x${maxSize}!"

            if ($ico_option -eq "best") {
                $runnerExe = $magickPath
                $runnerArgs = @($inputNormal, "-auto-orient")
                if ($stripMetadata) { $runnerArgs += "-strip" }
                $runnerArgs += @(
                    "-resize", $resizeArg,
                    "-background", "none",
                    "-define", "icon:auto-resize=$allSizes",
                    $encodeTargetNormal
                )
            }
            else {
                $runnerExe = $magickPath
                $runnerArgs = @($inputNormal, "-auto-orient")
                foreach ($size in ($icoSizes | Sort-Object -Descending)) {
                    $runnerArgs += @("(", "-clone", "0", "-resize", "${size}x${size}!", "-background", "white", "-alpha", "remove", "-quantize", "transparent", "+dither", "-colors", "256", ")")
                }
                if ($stripMetadata) { $runnerArgs += "-strip" }
                $runnerArgs += @("-delete", "0", $encodeTargetNormal)
            }
        }
    }
    elseif ($formatKey -eq "avif" -and $script:avifenc_available -and $script:avifencPath -and -not $isAnimated) {
        $useAvifencDirect = ($jpg_family_exts -contains $inputExt) -or ($inputExt -eq "png")
        $shouldPreprocessForAvifenc = (-not $useAvifencDirect) -or ($minimize -gt 1)
        if (-not $shouldPreprocessForAvifenc) {
            $runnerExe = $script:avifencPath
            $runnerArgs = New-AvifencArgs `
                -InputPath $inputNormal `
                -OutputPath $encodeTargetNormal `
                -QualityPlan $qualityPlan `
                -StripMetadata $stripMetadata
        }
        else {
            $tempPathInfo = New-UniqueTempImagePath -OutputDir $outputDir -BaseName $outputBaseName -Extension "png"
            $heicTempNormal = $tempPathInfo.NormalPath
            $heicTempLong = $tempPathInfo.LongPath

            $heicStep1Args = New-MagickArgsForFormat `
                -InputPath $inputNormal `
                -OutputPath $heicTempNormal `
                -FormatKey "png" `
                -InputExt $inputExt `
                -QualityPlan $qualityPlan `
                -ResizeArg $resizePlan.ResizeArg `
                -StripMetadata $stripMetadata `
                -IsAnimated $false

            $runnerExe = $script:avifencPath
            $runnerArgs = New-AvifencArgs `
                -InputPath $heicTempNormal `
                -OutputPath $encodeTargetNormal `
                -QualityPlan $qualityPlan `
                -StripMetadata $stripMetadata
        }
    }
    else {
        $runnerExe = $magickPath
            $runnerArgs = New-MagickArgsForFormat `
                -InputPath $inputNormal `
                -OutputPath $encodeTargetNormal `
                -FormatKey $formatKey `
                -InputExt $inputExt `
                -QualityPlan $qualityPlan `
                -ResizeArg $resizePlan.ResizeArg `
                -StripMetadata $stripMetadata `
                -IsAnimated $isAnimated
    }

    $inputBytes = Get-FileSizeBytes $inputLongOriginal

    $tasks.Add([PSCustomObject]@{
            Name               = $file.Name
            InputPathNormal    = $inputNormal
            InputPathLong      = $inputLongOriginal
            InputSizeBytes     = $inputBytes
            InputWidth         = $width
            InputHeight        = $height
            DidResize          = $resizePlan.DidResize
            NewSize            = $resizePlan.NewSize
            FinalOutputNormal  = $finalOutputNormal
            FinalOutputLong    = $finalOutputLong
            EncodeTargetNormal = $encodeTargetNormal
            EncodeTargetLong   = $encodeTargetLong
            UseTmpOverwrite    = $useTmpOverwrite
            ShouldCheckOutputSize = $shouldCheckOutputSize
            SkipOutputValidation = $skipOutputValidation
            DeleteInputAfterSuccess = $deleteInputAfterSuccess
            RenameInputOnOutputTooLarge = $renameInputOnOutputTooLarge
            FrameOutputGlob    = $frameOutputGlob
            FrameRenameGlob    = $frameRenameGlob
            FrameSuffixToken   = $frameSuffixToken
            UseFrameSuffix     = $useFrameSuffix
            RunnerExe          = $runnerExe
            RunnerArgs         = $runnerArgs
            HeicTempNormal     = $heicTempNormal
            HeicTempLong       = $heicTempLong
            HeicStep1Args      = $heicStep1Args
            SkipReason         = $skipReason
            RetryCount         = 0
        })
}

if ($tasks.Count -eq 0) {
    Write-EncodeErrorLog -Message $strings[$script:currentLang].LogNoValidTasks
    exit 0
}


# =====================================
# HẬU XỬ LÝ CHO TỪNG TÁC VỤ ĐÃ ENCODE
# =====================================

# Lấy danh sách file output theo pattern frame nếu có.
function Get-FrameOutputFiles {
    param([Parameter(Mandatory = $true)][pscustomobject]$Task)

    if (-not $Task.FrameOutputGlob) { return @() }
    $dir = Split-Path -Parent $Task.FrameOutputGlob
    $pattern = Split-Path -Leaf $Task.FrameOutputGlob
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return @() }
    return Get-ChildItem -Path $dir -Filter $pattern -File -ErrorAction SilentlyContinue
}

function Get-TotalSizeByGlob {
    param([Parameter(Mandatory = $true)][string]$Glob)

    if ([string]::IsNullOrWhiteSpace($Glob)) { return $null }
    $dir = Split-Path -Parent $Glob
    $pattern = Split-Path -Leaf $Glob
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return $null }
    $files = Get-ChildItem -Path $dir -Filter $pattern -File -ErrorAction SilentlyContinue
    if (-not $files -or $files.Count -eq 0) { return $null }
    $sum = 0L
    foreach ($f in $files) {
        if (-not $f) { continue }
        try { $sum += [long]$f.Length } catch { }
    }
    return $sum
}

# Hoàn tất tác vụ: kiểm tra output, xử lý overwrite, và ghi encode info nếu cần.
function Complete-TaskOutput {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Task,
        [Parameter(Mandatory = $true)][int]$ExitCode
    )

    $failureReason = $null
    $outputBytes = $null
    $skipValidation = ($Task.SkipOutputValidation -eq $true)
    $frameFiles = @()
    $hasFrameOutput = $false
    if ($Task.FrameOutputGlob) {
        $frameFiles = Get-FrameOutputFiles -Task $Task
        if ($frameFiles -and $frameFiles.Count -gt 0) { $hasFrameOutput = $true }
    }

    if ($ExitCode -ne 0) {
        $failureReason = "encode_exit_code_$ExitCode"
    }
    elseif ($Task.FrameOutputGlob) {
        if (-not $hasFrameOutput) {
            $failureReason = "output_missing"
        }
        else {
            foreach ($f in $frameFiles) {
                if (-not $f) { continue }
                try { $outputBytes += [long]$f.Length } catch { }
            }
            if ($outputBytes -le 0) {
                $failureReason = "output_empty"
            }
        }
    }
    elseif (-not $skipValidation) {
        if (-not (Test-Path -LiteralPath $Task.EncodeTargetLong)) {
            $failureReason = "output_missing"
        }
        else {
            $outputBytes = Get-FileSizeBytes $Task.EncodeTargetLong
            if ($outputBytes -le 0) {
                $failureReason = "output_empty"
            }
        }
    }

    if (-not $failureReason -and -not $skipValidation -and -not $Task.FrameOutputGlob) {
        # Mức 2: output validation để đảm bảo file hợp lệ (không bị rỗng/giả hợp lệ).
        $outputProbe = Get-ImageDimensions -FilePathLong $Task.EncodeTargetLong

        if (-not $outputProbe) {
            $failureReason = "output_probe_failed"
        }
    }

    if ($failureReason) {
        $failureDetail = Get-EncodeFailureDetail -Reason $failureReason
        if ($failureReason -like "encode_exit_code_*") {
            $failureDetail = $strings[$script:currentLang].LogOutputWriteFailed
        }
        Write-EncodeErrorLog -Message ($strings[$script:currentLang].LogEncodeFailed -f $Task.InputPathNormal, $failureDetail)

        Clear-HeicTempIfExists -Task $Task

        Add-TaskReport -Task $Task -Status "failed" -Reason $failureReason

        return [PSCustomObject]@{
            Success       = $false
            FailureReason = $failureReason
            Result        = $null
            OutputTooLarge = $false
        }
    }

    $outputTooLarge = $false
    # Nếu đang ở chế độ ghi đè và output lớn hơn input -> xóa output, giữ nguyên input.
    if (-not $skipValidation -and $Task.ShouldCheckOutputSize -and $outputBytes -gt $Task.InputSizeBytes) {
        $outputTooLarge = $true
        if (Test-Path -LiteralPath $Task.EncodeTargetLong) {
            Remove-Item -LiteralPath $Task.EncodeTargetLong -Force -ErrorAction SilentlyContinue
        }

        if ($Task.RenameInputOnOutputTooLarge) {
            # Đổi đuôi input sang đuôi output khi output lớn hơn input.
            try {
                Move-Item -LiteralPath $Task.InputPathLong -Destination $Task.FinalOutputLong -Force
            }
            catch {
                # Nếu không đổi được thì giữ nguyên input.
            }
        }

        $resultOutputPath = $null
        $resultOutputBytes = $Task.InputSizeBytes
        if ($Task.FinalOutputLong -and (Test-Path -LiteralPath $Task.FinalOutputLong)) {
            $resultOutputPath = $Task.FinalOutputNormal
            try { $resultOutputBytes = Get-FileSizeBytes $Task.FinalOutputLong } catch {}
        }
        elseif ($Task.InputPathLong -and (Test-Path -LiteralPath $Task.InputPathLong)) {
            $resultOutputPath = $Task.InputPathNormal
            try { $resultOutputBytes = Get-FileSizeBytes $Task.InputPathLong } catch {}
        }
        if (-not $resultOutputPath) { $resultOutputPath = $Task.FinalOutputNormal }

        Clear-HeicTempIfExists -Task $Task

        Add-TaskReport -Task $Task -Status "success" -Reason "output_too_large"

        return [PSCustomObject]@{
            Success        = $true
            FailureReason  = $null
            OutputTooLarge = $true
            Result         = [PSCustomObject]@{
                Name           = $Task.Name
                InputSizeBytes  = $Task.InputSizeBytes
                OutputSizeBytes = $resultOutputBytes
                OutputPath     = $resultOutputPath
                OutputTooLarge = $true
            }
        }
    }

    if ($Task.UseTmpOverwrite) {
        if (-not $skipValidation -or (Test-Path -LiteralPath $Task.EncodeTargetLong)) {
            Remove-Item -LiteralPath $Task.InputPathLong -Force
            Move-Item -LiteralPath $Task.EncodeTargetLong -Destination $Task.FinalOutputLong -Force
        }
    }

    if ($Task.FrameOutputGlob -and $hasFrameOutput) {
        foreach ($f in $frameFiles) {
            if (-not $f) { continue }
            $newName = $f.Name.Replace($Task.FrameSuffixToken, "-")
            if ($newName -eq $f.Name) { continue }
            $newPath = Join-Path $f.DirectoryName $newName
            if (Test-Path -LiteralPath $newPath) { continue }
            try { Move-Item -LiteralPath $f.FullName -Destination $newPath } catch { }
        }
    }

    $finalBytes = $Task.InputSizeBytes
    if ($Task.FrameOutputGlob -and $hasFrameOutput -and $null -ne $outputBytes) {
        $finalBytes = $outputBytes
    }
    elseif (Test-Path -LiteralPath $Task.FinalOutputLong) {
        try { $finalBytes = Get-FileSizeBytes $Task.FinalOutputLong } catch { $finalBytes = $Task.InputSizeBytes }
    }

    $hasFinalOutput = $false
    if ($Task.FrameOutputGlob -and $hasFrameOutput) {
        $hasFinalOutput = $true
    }
    elseif (Test-Path -LiteralPath $Task.FinalOutputLong) {
        $hasFinalOutput = $true
    }

    if ($Task.DeleteInputAfterSuccess -and (-not $skipValidation -or $hasFinalOutput)) {
        try {
            Remove-Item -LiteralPath $Task.InputPathLong -Force
        }
        catch {
            # Nếu không xóa được thì bỏ qua để tránh dừng batch.
        }
    }

    Clear-HeicTempIfExists -Task $Task

    Add-TaskReport -Task $Task -Status "success" -Reason $null

    return [PSCustomObject]@{
        Success        = $true
        FailureReason  = $null
        OutputTooLarge = $outputTooLarge
        Result         = [PSCustomObject]@{
            Name           = $Task.Name
            InputSizeBytes  = $Task.InputSizeBytes
            OutputSizeBytes = $finalBytes
            OutputPath     = $Task.FinalOutputNormal
            OutputFramesGlob = if ($Task.FrameRenameGlob) { $Task.FrameRenameGlob } else { $Task.FrameOutputGlob }
            OutputFramesGlobAlt = $Task.FrameOutputGlob
            OutputTooLarge = $outputTooLarge
        }
    }
}

# Tạo entry retry chuẩn hóa.
function New-RetryEntry {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Task,
        [Parameter(Mandatory = $true)][int]$Attempt,
        [Parameter(Mandatory = $true)][string]$Reason
    )
    return [PSCustomObject]@{
        Task    = $Task
        Attempt = $Attempt
        Reason  = $Reason
    }
}

# Đưa task vào hàng đợi retry nếu còn lượt.
function Add-RetryIfPossible {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Task,
        [Parameter(Mandatory = $true)][string]$Reason,
        [Parameter(Mandatory = $true)][int]$MaxRetry,
        [AllowEmptyCollection()]
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[object]]$RetryQueue
    )

    if ($Task.RetryCount -lt $MaxRetry) {
        $Task.RetryCount++
        $RetryQueue.Add((New-RetryEntry -Task $Task -Attempt $Task.RetryCount -Reason $Reason)) | Out-Null
        return $true
    }
    return $false
}

# Chạy một task trong cùng process (tuần tự).
function Invoke-TaskLocal {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Task,
        [AllowEmptyCollection()]
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[object]]$Results
    )

    $pre = Invoke-HeicPreprocess -Task $Task -MagickPath $magickPath
    if (-not $pre.Success) {
        Clear-HeicTempIfExists -Task $Task
        return [PSCustomObject]@{
            Success         = $false
            FailureReason   = $pre.Reason
            PreprocessFailed = $true
        }
    }

    & $Task.RunnerExe @($Task.RunnerArgs) 2>$null
    $exitCode = $LASTEXITCODE

    $outcome = Complete-TaskOutput -Task $Task -ExitCode $exitCode
    if ($outcome.Success) {
        if ($outcome.Result) { $Results.Add($outcome.Result) }
    }

    return $outcome
}

# Khởi chạy task bằng process (song song).
function Start-TaskProcess {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Task
    )

    $pre = Invoke-HeicPreprocess -Task $Task -MagickPath $magickPath
    if (-not $pre.Success) {
        Clear-HeicTempIfExists -Task $Task
        return [PSCustomObject]@{
            Started       = $false
            FailureReason = $pre.Reason
        }
    }

    try {
        $argLine = ConvertTo-ProcessArgumentString -Arguments $Task.RunnerArgs
        $proc = Start-Process `
            -FilePath $Task.RunnerExe `
            -ArgumentList $argLine `
            -PassThru `
            -WindowStyle Hidden
    }
    catch {
        $failReason = "start_process_failed"
        return [PSCustomObject]@{
            Started       = $false
            FailureReason = $failReason
        }
    }

    return [PSCustomObject]@{
        Started = $true
        Process = $proc
    }
}

# Chạy một batch retry với số worker cố định.
function Invoke-RetryBatch {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[object]]$RetryEntries,
        [Parameter(Mandatory = $true)][int]$MaxWorkers
    )

    $failedEntries = New-Object System.Collections.Generic.List[object]
    if ($RetryEntries.Count -eq 0) { return $failedEntries }

    if ($MaxWorkers -le 1) {
        foreach ($entry in $RetryEntries) {
            $task = $entry.Task
            $detail = Get-EncodeFailureDetail -Reason $entry.Reason
            Write-EncodeErrorLog -Message ($strings[$script:currentLang].LogRetryAttempt -f $entry.Attempt, $task.Name, $detail)
            $outcome = Invoke-TaskLocal -Task $task -Results $results
            if (-not $outcome.Success) {
                $failedEntries.Add((New-RetryEntry -Task $task -Attempt ($entry.Attempt + 1) -Reason $outcome.FailureReason)) | Out-Null
            }
        }

        return $failedEntries
    }

    $pendingQueue = [System.Collections.Generic.Queue[object]]::new()
    foreach ($entry in $RetryEntries) { $pendingQueue.Enqueue($entry) }
    $running = @{}

    while ($pendingQueue.Count -gt 0 -or $running.Count -gt 0) {
        while ($running.Count -lt $MaxWorkers -and $pendingQueue.Count -gt 0) {
            $entry = $pendingQueue.Dequeue()
            $task = $entry.Task
            $detail = Get-EncodeFailureDetail -Reason $entry.Reason
            Write-EncodeErrorLog -Message ($strings[$script:currentLang].LogRetryAttempt -f $entry.Attempt, $task.Name, $detail)
            $start = Start-TaskProcess -Task $task
            if (-not $start.Started) {
                $failedEntries.Add((New-RetryEntry -Task $task -Attempt ($entry.Attempt + 1) -Reason $start.FailureReason)) | Out-Null
                continue
            }

            $running[$start.Process.Id] = [PSCustomObject]@{
                Process = $start.Process
                Task    = $task
                Attempt = $entry.Attempt
                Reason  = $entry.Reason
            }
        }

        if ($running.Count -eq 0) { continue }

        Start-Sleep -Milliseconds 100

        foreach ($id in @($running.Keys)) {
            $entry = $running[$id]
            $proc = $entry.Process
            if ($proc.HasExited) {
                $outcome = Complete-TaskOutput -Task $entry.Task -ExitCode $proc.ExitCode
                if ($outcome.Success) {
                    if ($outcome.Result) { $results.Add($outcome.Result) }
                }
                else {
                    $failedEntries.Add((New-RetryEntry -Task $entry.Task -Attempt ($entry.Attempt + 1) -Reason $outcome.FailureReason)) | Out-Null
                }
                $running.Remove($id)
            }
        }
    }

    return $failedEntries
}


# =====================================
# CHẾ ĐỘ CHẠY: TUẦN TỰ HOẶC SONG SONG ĐỘNG
# =====================================

# Ghi lỗi cho task bị skip (theo yêu cầu đặc biệt).
function Write-SkipTaskLog {
    param([Parameter(Mandatory = $true)][pscustomobject]$Task)

    $reasonText = $Task.SkipReason
    if ([string]::IsNullOrWhiteSpace($reasonText)) {
        $reasonText = $strings[$script:currentLang].LogOutputWriteFailed
    }
    Write-EncodeErrorLog -Message ($strings[$script:currentLang].LogEncodeFailed -f $Task.InputPathNormal, $reasonText)
}

# Danh sách kết quả để tổng hợp thống kê sau khi encode.
$results = New-Object System.Collections.Generic.List[object]
$retryQueue = New-Object System.Collections.Generic.List[object]
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$useAdaptiveRuntime = $enable_adaptive_parallel

# AVIF luôn chạy tuần tự để tránh CPU 100% dù chỉ 1 worker.
if ($converter -eq "avif") {
    $useAdaptiveRuntime = $false
}

if ($enable_adaptive_parallel) {
    if ($logicalThreads -le 4) {
        $useAdaptiveRuntime = $false
    }
    elseif ($logicalThreads -ge 5 -and $logicalThreads -le 9 -and $tasks.Count -lt $adaptive_min_batch_threads_5_9) {
        $useAdaptiveRuntime = $false
    }
}

if (-not $useAdaptiveRuntime) {
    foreach ($task in $tasks) {
        if ($task.SkipReason) {
            Write-SkipTaskLog -Task $task
            Add-TaskReport -Task $task -Status "failed" -Reason "skip"
            continue
        }

        $outcome = Invoke-TaskLocal -Task $task -Results $results
        if (-not $outcome.Success) {
            Add-RetryIfPossible -Task $task -Reason $outcome.FailureReason -MaxRetry $retry_count -RetryQueue $retryQueue | Out-Null
        }
    }
}
else {
    $pendingQueue = [System.Collections.Generic.Queue[object]]::new()
    foreach ($task in $tasks) { $pendingQueue.Enqueue($task) }

    # Ước tính tải nền để chọn số worker khởi tạo hợp lý ngay từ đầu.
    $maxWorkers = [Math]::Max(1, $logicalThreads)
    $startupMetrics = Get-SystemLoadMetrics
    $startupPlan = Get-AdaptiveStartupPlan `
        -CpuPercent $startupMetrics.CPU `
        -LogicalThreads $logicalThreads `
        -StartFreeRatio $adaptive_start_free_ratio `
        -PendingCount $pendingQueue.Count

    $targetWorkers = [Math]::Max(1, [Math]::Min($maxWorkers, $startupPlan.RecommendedWorkers))

    $pollMilliseconds = [Math]::Max(50, $adaptive_poll_milliseconds)
    $scaleUpStep = [Math]::Max(0, $adaptive_scaleup_step)
    $scaleDownStep = [Math]::Max(1, $adaptive_scaledown_step)
    $monitorInterval = [TimeSpan]::FromSeconds([Math]::Max(0.2, [double]$monitor_interval_seconds))
    $nextMonitorUtc = [DateTime]::UtcNow

    $running = @{}

    while ($pendingQueue.Count -gt 0 -or $running.Count -gt 0) {

        while ($running.Count -lt $targetWorkers -and $pendingQueue.Count -gt 0) {
            $task = $pendingQueue.Dequeue()

            if ($task.SkipReason) {
                Write-SkipTaskLog -Task $task
                Add-TaskReport -Task $task -Status "failed" -Reason "skip"
                continue
            }

            if ($task.DidResize) {
            }
            else {
            }
            $start = Start-TaskProcess -Task $task
            if (-not $start.Started) {
                Add-RetryIfPossible -Task $task -Reason $start.FailureReason -MaxRetry $retry_count -RetryQueue $retryQueue | Out-Null
                continue
            }

    $running[$start.Process.Id] = [PSCustomObject]@{
        Process = $start.Process
        Task    = $task
    }
        }

        if ($running.Count -eq 0) {
            continue
        }

        Start-Sleep -Milliseconds $pollMilliseconds

        foreach ($id in @($running.Keys)) {
            $entry = $running[$id]
            $proc = $entry.Process
            if ($proc.HasExited) {
                $outcome = Complete-TaskOutput -Task $entry.Task -ExitCode $proc.ExitCode
                if ($outcome.Success) {
                    if ($outcome.Result) { $results.Add($outcome.Result) }
                }
                else {
                    Add-RetryIfPossible -Task $entry.Task -Reason $outcome.FailureReason -MaxRetry $retry_count -RetryQueue $retryQueue | Out-Null
                }
                $running.Remove($id)
            }
        }

        if ([DateTime]::UtcNow -ge $nextMonitorUtc) {
            $metrics = Get-SystemLoadMetrics
            if ($metrics.CPU -gt $high_threshold -or $metrics.RAM -gt $high_threshold -or $metrics.IO -gt $high_threshold) {
                # Quá ngưỡng: giảm nhanh để nhường tài nguyên.
                $targetWorkers = [Math]::Max(1, $targetWorkers - $scaleDownStep)
            }
            elseif ($metrics.CPU -lt $low_threshold -and $metrics.RAM -lt $low_threshold -and $metrics.IO -lt $low_threshold) {
                # Còn headroom: tăng đều mỗi chu kỳ.
                if ($scaleUpStep -gt 0 -and $pendingQueue.Count -gt 0 -and $targetWorkers -lt $maxWorkers) {
                    $targetWorkers = [Math]::Min($maxWorkers, $targetWorkers + $scaleUpStep)
                }
            }

            $nextMonitorUtc = [DateTime]::UtcNow.Add($monitorInterval)
        }
    }
}

# =====================================
# RETRY SAU BATCH CHÍNH
# =====================================

if ($retryQueue.Count -gt 0) {
    # Chạy lại các file lỗi theo từng lượt: lần 1 bình thường, lần 2 giảm tải, lần 3 tuần tự.
    $currentRetries = $retryQueue
    $retryQueue = New-Object System.Collections.Generic.List[object]

    for ($attempt = 1; $attempt -le $retry_count; $attempt++) {
        $batch = New-Object System.Collections.Generic.List[object]
        foreach ($entry in $currentRetries) {
            if ($entry.Attempt -eq $attempt) {
                $batch.Add($entry)
            }
            else {
                $retryQueue.Add($entry)
            }
        }

        if ($batch.Count -eq 0) {
            $currentRetries = $retryQueue
            $retryQueue = New-Object System.Collections.Generic.List[object]
            continue
        }

        $maxWorkers = 1
        if ($enable_adaptive_parallel) {
            $startupMetrics = Get-SystemLoadMetrics
            $startupPlan = Get-AdaptiveStartupPlan `
                -CpuPercent $startupMetrics.CPU `
                -LogicalThreads $logicalThreads `
                -StartFreeRatio $adaptive_start_free_ratio `
                -PendingCount $batch.Count

            $baseWorkers = [Math]::Max(1, [Math]::Min($logicalThreads, $startupPlan.RecommendedWorkers))

            if ($attempt -eq 1) {
                $maxWorkers = $baseWorkers
            }
            elseif ($attempt -eq 2) {
                $maxWorkers = [Math]::Max(1, $baseWorkers - $adaptive_scaledown_step)
            }
            else {
                $maxWorkers = 1
            }
        }

        $failed = Invoke-RetryBatch -RetryEntries $batch -MaxWorkers $maxWorkers
        $currentRetries = $failed
        $retryQueue = New-Object System.Collections.Generic.List[object]
    }

}

# Nếu metadata làm thay đổi kích thước file (ví dụ ghi XMP), cập nhật lại số liệu báo cáo cho chính xác.
foreach ($item in $results) {
    try {
        if ($item.PSObject.Properties.Name -contains "OutputFramesGlob" -and $item.OutputFramesGlob) {
            $sum = Get-TotalSizeByGlob -Glob $item.OutputFramesGlob
            if ($null -eq $sum -and ($item.PSObject.Properties.Name -contains "OutputFramesGlobAlt") -and $item.OutputFramesGlobAlt) {
                $sum = Get-TotalSizeByGlob -Glob $item.OutputFramesGlobAlt
            }
            if ($null -ne $sum) {
                $item.OutputSizeBytes = $sum
                continue
            }
        }
        $item.OutputSizeBytes = Get-FileSizeBytes (Convert-ToLongPath $item.OutputPath)
    }
    catch {
        # Bỏ qua nếu file không còn tồn tại.
    }
}

$stopwatch.Stop()

# =====================================
# BÁO CÁO CUỐI CÙNG (YÊU CẦU G)
# =====================================

$totalInputBytes = 0.0
$totalOutputBytes = 0.0

foreach ($item in $results) {
    $totalInputBytes += [double]$item.InputSizeBytes
    $totalOutputBytes += [double]$item.OutputSizeBytes
}

if ($create_folder -and $createdOutputDirs.Count -gt 0) {
    foreach ($dir in $createdOutputDirs) {
        try {
            Invoke-Item $dir
        }
        catch {
            # Bỏ qua nếu không thể mở thư mục.
        }
    }
}

# Trả về summary trực tiếp từ trong hàm (đúng scope của $script:TaskReport và $results).
$successItems = @($script:TaskReport | Where-Object { $_.Status -eq "success" })
$successCount = $successItems.Count
$summaryInputBytes = 0.0
$summaryOutputBytes = 0.0
$successInputPaths = New-Object System.Collections.Generic.List[string]
foreach ($entry in $successItems) {
    $inPath = [string]$entry.InputPath
    $outPath = [string]$entry.OutputPath
    if (-not [string]::IsNullOrWhiteSpace($inPath)) {
        $successInputPaths.Add($inPath) | Out-Null
        if (Test-Path -LiteralPath $inPath) {
            try { $summaryInputBytes += (Get-Item -LiteralPath $inPath).Length } catch {}
        }
    }
    $outputAdded = $false
    if ($entry.PSObject.Properties.Name -contains "OutputFramesGlob") {
        $sum = $null
        if ($entry.OutputFramesGlob) {
            $sum = Get-TotalSizeByGlob -Glob $entry.OutputFramesGlob
        }
        if (($null -eq $sum) -and ($entry.PSObject.Properties.Name -contains "OutputFramesGlobAlt") -and $entry.OutputFramesGlobAlt) {
            $sum = Get-TotalSizeByGlob -Glob $entry.OutputFramesGlobAlt
        }
        if ($null -ne $sum -and $sum -gt 0) {
            $summaryOutputBytes += $sum
            $outputAdded = $true
        }
    }
    if (-not $outputAdded -and -not [string]::IsNullOrWhiteSpace($outPath) -and (Test-Path -LiteralPath $outPath)) {
        try { $summaryOutputBytes += (Get-Item -LiteralPath $outPath).Length } catch {}
    }
}
$summaryPercentChange = 0.0
if ($summaryInputBytes -gt 0) {
    $summaryPercentChange = (($summaryOutputBytes - $summaryInputBytes) / $summaryInputBytes) * 100
}
# Trả summary về qua $script: scope để runspace đọc được (tránh lẫn vào output stream).
$script:__QHSummary = [PSCustomObject]@{
    SuccessCount     = $successCount
    TotalInputBytes  = $summaryInputBytes
    TotalOutputBytes = $summaryOutputBytes
    PercentChange    = $summaryPercentChange
    SuccessInputs    = $successInputPaths.ToArray()
}

} # Kết thúc hàm Invoke-EncodeLogic.

if (-not $script:LogicBootstrap) {


function Format-Bytes {
    param(
        [double]$Bytes,
        [switch]$LimitToMb
    )
    if (-not $LimitToMb -and $Bytes -ge 1GB) { return "{0:0.00}GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:0.00}MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:0.00}KB" -f ($Bytes / 1KB) }
    return "{0:0}B" -f $Bytes
}

function ConvertTo-Number {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $clean = $Text.Trim().Replace(",", ".")
    $value = 0.0
    $ok = [double]::TryParse(
        $clean,
        [System.Globalization.NumberStyles]::AllowDecimalPoint,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$value
    )
    if ($ok) { return $value }
    return $null
}

# Convert size text like "3024x4032" (from listview) to Width/Height.
function ConvertFrom-SizeText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $m = [regex]::Match($Text, "(\d+)\s*[xX]\s*(\d+)")
    if (-not $m.Success) { return $null }
    return [PSCustomObject]@{
        Width  = [int]$m.Groups[1].Value
        Height = [int]$m.Groups[2].Value
    }
}

function Get-SizeSortKey {
    param([string]$Text)
    $sz = ConvertFrom-SizeText $Text
    if (-not $sz) { return $null }
    $w = [double]$sz.Width
    $h = [double]$sz.Height
    return [PSCustomObject]@{
        Area = $w * $h
        Max  = [Math]::Max($w, $h)
        Min  = [Math]::Min($w, $h)
    }
}

function Get-BytesSortKey {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $m = [regex]::Match($Text.Trim(), "^([\d\.,]+)\s*(B|KB|MB|GB|TB)$", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) { return $null }
    $numText = $m.Groups[1].Value -replace ",", "."
    $unit = $m.Groups[2].Value.ToUpperInvariant()
    $val = 0.0
    try { $val = [double]::Parse($numText, [System.Globalization.CultureInfo]::InvariantCulture) } catch { return $null }
    $factor = switch ($unit) {
        "B"  { 1.0 }
        "KB" { 1024.0 }
        "MB" { 1024.0 * 1024.0 }
        "GB" { 1024.0 * 1024.0 * 1024.0 }
        "TB" { 1024.0 * 1024.0 * 1024.0 * 1024.0 }
        default { 1.0 }
    }
    return ($val * $factor)
}

function New-KeyComparer {
    param(
        [Parameter(Mandatory = $true)][int]$Direction,
        [Parameter(Mandatory = $true)][scriptblock]$KeySelector,
        [Parameter(Mandatory = $true)][scriptblock]$CompareKeys
    )
    return [System.Collections.Generic.Comparer[object]]::Create({
            param($x, $y)
            $vx = & $KeySelector $x
            $vy = & $KeySelector $y
            if ($null -eq $vx -and $null -eq $vy) { return 0 }
            if ($null -eq $vx) { return 1 }
            if ($null -eq $vy) { return -1 }
            return (& $CompareKeys $vx $vy $Direction)
        })
}

function New-SizeComparer {
    param([Parameter(Mandatory = $true)][int]$Direction)
    return New-KeyComparer -Direction $Direction `
        -KeySelector { param($x) Get-SizeSortKey $x.Size } `
        -CompareKeys {
            param($vx, $vy, $dir)
            if ($vx.Area -ne $vy.Area) {
                if ($vx.Area -lt $vy.Area) { return -1 * $dir }
                return 1 * $dir
            }
            if ($vx.Max -ne $vy.Max) {
                if ($vx.Max -lt $vy.Max) { return -1 * $dir }
                return 1 * $dir
            }
            if ($vx.Min -ne $vy.Min) {
                if ($vx.Min -lt $vy.Min) { return -1 * $dir }
                return 1 * $dir
            }
            return 0
        }
}

function New-BytesComparer {
    param([Parameter(Mandatory = $true)][int]$Direction)
    return New-KeyComparer -Direction $Direction `
        -KeySelector { param($x) Get-BytesSortKey $x.Bytes } `
        -CompareKeys {
            param($vx, $vy, $dir)
            if ($vx -eq $vy) { return 0 }
            if ($vx -lt $vy) { return -1 * $dir }
            return 1 * $dir
        }
}

function Test-InvalidFileName {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return $false }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    return ($Value.IndexOfAny($invalid) -ge 0)
}

function Update-ListIndexMap {
    $script:listIndexMap = @{}
    if (-not $script:listItems) { return }
    for ($i = 0; $i -lt $script:listItems.Count; $i++) {
        $item = $script:listItems[$i]
        if ($item -and $item.Path) {
            $key = ([string]$item.Path).ToLowerInvariant()
            $script:listIndexMap[$key] = $i
        }
    }
}

function Reset-StatusLog {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Strings,
        [Parameter(Mandatory = $true)][int]$TotalCount
    )
    $null = $Strings

    if ($script:StatusLog) {
        try { $script:StatusLog.Clear() } catch {}
    }
    $script:StatusTotalCount = [int]$TotalCount
}

function Add-StatusSummaryLine {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Summary,
        [Parameter(Mandatory = $true)][hashtable]$Strings
    )

    $successCount = [int]$Summary.SuccessCount
    $totalCount = [int]$script:StatusTotalCount
    $failCount = $totalCount - $successCount
    if ($failCount -lt 0) { $failCount = 0 }

$initialText = Format-Bytes -Bytes ([double]$Summary.TotalInputBytes) -LimitToMb
$currentText = Format-Bytes -Bytes ([double]$Summary.TotalOutputBytes) -LimitToMb
    $percentText = ("{0:+0.00;-0.00;0.00}%" -f ([double]$Summary.PercentChange))

    $totalSeconds = 0
    if ($script:EncodeStartTime) {
        $elapsed = (Get-Date) - $script:EncodeStartTime
        $totalSeconds = [int][Math]::Round($elapsed.TotalSeconds)
    }
    $elapsedMinutes = [int]($totalSeconds / 60)
    $elapsedSeconds = $totalSeconds % 60

    $template = $Strings.AfterEncodeTemplate
    $isValidTemplate = $true
    if ([string]::IsNullOrWhiteSpace($template)) {
        $isValidTemplate = $false
    } else {
        foreach ($token in @("{0}", "{1}", "{2}", "{3}", "{4}", "{5}", "{6}")) {
            if ($template -notlike "*$token*") {
                $isValidTemplate = $false
                break
            }
        }
    }
    if (-not $isValidTemplate) {
        $template = $Strings.AfterEncodeTemplate
    }
    $summaryMessage = [string]::Format(
        $template,
        $successCount,
        $failCount,
        $initialText,
        $currentText,
        $percentText,
        $elapsedMinutes,
        $elapsedSeconds
    )
    Add-StatusLogLine $summaryMessage
}


function Add-FilesToList {
    param(
        [Parameter(Mandatory = $true)][object[]]$Files,
        [Parameter(Mandatory = $true)][hashtable]$Strings,
        [bool]$ShowLoading = $true
    )

    $loggedSummary = $false

    if ($window -and $window.Dispatcher -and -not $window.Dispatcher.CheckAccess()) {
        $window.Dispatcher.Invoke([action]{
            Add-FilesToList -Files $Files -Strings $Strings -ShowLoading:$ShowLoading
        })
        return
    }

    try {
        if (-not $script:StatusLog -and $window) {
            $script:StatusLog = $window.FindName("StatusLog")
        }

        $validCount = 0
        $skipUnsupported = 0
        $skipDuplicate = 0
        $skipMissing = 0
        if (-not $script:listIndexMap -or $script:listIndexMap.Count -ne $script:listItems.Count) {
            Update-ListIndexMap
        }

        if ($ShowLoading) {
            Show-LoadingDialog -Owner $window -Message $Strings.LoadingList
        }

        # Đo kích thước theo batch bằng magick (-ping trước, fallback không -ping), tránh gọi magick từng file.
        $dimensionMap = @{}
        $sizeErrorCount = 0
        $fileInfoBatch = New-Object System.Collections.Generic.List[System.IO.FileInfo]
        foreach ($item in $Files) {
            if ($item -is [System.IO.FileInfo]) {
                $fileInfoBatch.Add($item) | Out-Null
                continue
            }
            $path = [string]$item
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            if (-not (Test-Path -LiteralPath $path)) { continue }
            try {
                $fi = Get-Item -LiteralPath $path -ErrorAction Stop
                if ($fi -and -not $fi.PSIsContainer) {
                    $fileInfoBatch.Add($fi) | Out-Null
                }
            }
            catch {
                # Bỏ qua file lỗi khi thu thập batch.
            }
        }
        if ($fileInfoBatch.Count -gt 0) {
            $magickThreads = [Math]::Max(1, $logicalThreads)
            $failedKeys = $null
            $dimensionMap = Get-ImageDimensionsMagickAllMap `
                -Files $fileInfoBatch `
                -MagickPath $magickPath `
                -MagickThreadCount $magickThreads `
                -InputSizeMap $null `
                -FailedKeys ([ref]$failedKeys)
        }

        # Tắt cập nhật per-item: gom kết quả trước, sau đó gán ItemsSource một lần cuối.
        $itemsToAdd = New-Object System.Collections.Generic.List[object]
        $newKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($item in $Files) {
            $fileInfo = $null
            $file = $null
            $ext = $null

            if ($item -is [System.IO.FileInfo]) {
                $fileInfo = $item
                $file = $fileInfo.FullName
                $ext = $fileInfo.Extension.ToLowerInvariant()
            } else {
                $file = [string]$item
                if ([string]::IsNullOrWhiteSpace($file)) {
                    $skipMissing++
                    continue
                }
                if (-not (Test-Path -LiteralPath $file)) {
                    $skipMissing++
                    continue
                }
                $ext = [IO.Path]::GetExtension($file).ToLowerInvariant()
            }

            if ($script:supportedExtensions -notcontains $ext) {
                $skipUnsupported++
                continue
            }

            $key = $file.ToLowerInvariant()
            if ($script:listIndexMap.ContainsKey($key) -or $newKeys.Contains($key)) {
                $skipDuplicate++
                continue
            }

            $sizeText = "?"
            if ($dimensionMap -and $dimensionMap.ContainsKey($key)) {
                $info = $dimensionMap[$key]
                $sizeText = "{0}x{1}" -f $info.Width, $info.Height
            } else {
                $sizeErrorCount++
            }
            if (-not $fileInfo) {
                try { $fileInfo = Get-Item -LiteralPath $file } catch { $fileInfo = $null }
            }
            $bytesText = if ($fileInfo) { Format-Bytes -Bytes $fileInfo.Length } else { "?" }

            $obj = [PSCustomObject]@{
                Name = $file
                Size = $sizeText
                Bytes = $bytesText
                Status = ""
                Path = $file
            }
            $itemsToAdd.Add($obj) | Out-Null
            $null = $newKeys.Add($key)
            $validCount++
        }

        if ($validCount -eq 0) {
            if ($ShowLoading) { Close-LoadingDialog }
            $msg = $Strings.UnsupportedMsg -replace "<EXT>", $extensionsText
            Add-StatusLogLine -Text $msg
        } elseif ($itemsToAdd.Count -gt 0) {
            # Gắn lại ItemsSource một lần để UI chỉ cập nhật sau khi hoàn tất.
            $combined = New-Object System.Collections.ObjectModel.ObservableCollection[object]
            foreach ($old in $script:listItems) { $combined.Add($old) | Out-Null }
            foreach ($newItem in $itemsToAdd) { $combined.Add($newItem) | Out-Null }
            $script:listItems = $combined
            $listItems = $combined
            if ($script:listFiles) { $script:listFiles.ItemsSource = $combined }
            Update-ListIndexMap
        }

        # Ghi log tổng kết sau khi thêm danh sách.
        if ($validCount -gt 0) {
            if ($sizeErrorCount -gt 0) {
                $summaryText = ($Strings.LogAddedFilesWithErrors -f $validCount, $sizeErrorCount)
            } else {
                $summaryText = ($Strings.LogAddedFiles -f $validCount)
            }
            Add-StatusLogLine -Text $summaryText
            $loggedSummary = $true
        }
    } catch {
        $script:LastUiError = $_.Exception
    } finally {
        if (-not $loggedSummary -and $validCount -gt 0) {
            if ($sizeErrorCount -gt 0) {
                Add-StatusLogLine -Text ($Strings.LogAddedFilesWithErrors -f $validCount, $sizeErrorCount)
            } else {
                Add-StatusLogLine -Text ($Strings.LogAddedFiles -f $validCount)
            }
        }
        if ($ShowLoading) { Close-LoadingDialog }
        Update-EncodeButtonState
    }
}

# Hàm liệt kê file trong thư mục
function Add-FolderToList {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][bool]$IncludeSub,
        [Parameter(Mandatory = $true)][hashtable]$Strings
    )

    if ($window -and $window.Dispatcher -and -not $window.Dispatcher.CheckAccess()) {
        $window.Dispatcher.Invoke([action]{
            Add-FolderToList -Folder $Folder -IncludeSub $IncludeSub -Strings $Strings
        })
        return
    }

    try {
        Show-LoadingDialog -Owner $window -Message $Strings.LoadingList

        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return }
        $files = if ($IncludeSub) {
            Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction SilentlyContinue
        } else {
            Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue
        }

        if ($files.Count -eq 0) { return }

        $imageFiles = $files | Where-Object { $script:supportedExtensions -contains $_.Extension.ToLowerInvariant() }
        if ($imageFiles.Count -gt 0) {
            Add-FilesToList -Files $imageFiles -Strings $Strings -ShowLoading:$false
        } else {
            Close-LoadingDialog
            $msg = $Strings.UnsupportedMsg -replace "<EXT>", $extensionsText
            Add-StatusLogLine -Text $msg
        }
    } catch {
        $script:LastUiError = $_.Exception
    } finally {
        Close-LoadingDialog
    }
}

# ==============================
# HÀM HỖ TRỢ UI
# ==============================
function Get-Shell32Path {
    $systemDir = [Environment]::GetFolderPath('System')
    return (Join-Path $systemDir 'shell32.dll')
}

function Get-ShellIconImage {
    param(
        [Parameter(Mandatory = $true)][int]$Index,
        [int]$Size = 32
    )

    $shell32 = Get-Shell32Path
    $hModule = [ShellIcon]::LoadLibraryEx($shell32, [IntPtr]::Zero, [ShellIcon]::LOAD_LIBRARY_AS_DATAFILE)
    if ($hModule -eq [IntPtr]::Zero) { return $null }

    try {
        $iconSize = if ($Size -gt 0) { $Size } else { 32 }
        $hIcon = [ShellIcon]::LoadImage($hModule, [IntPtr]$Index, [ShellIcon]::IMAGE_ICON, $iconSize, $iconSize, 0)
        if ($hIcon -eq [IntPtr]::Zero) { return $null }

        $source = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
            $hIcon,
            [System.Windows.Int32Rect]::Empty,
            [System.Windows.Media.Imaging.BitmapSizeOptions]::FromWidthAndHeight($iconSize, $iconSize)
        )

        [ShellIcon]::DestroyIcon($hIcon) | Out-Null
        return $source
    }
    finally {
        [ShellIcon]::FreeLibrary($hModule) | Out-Null
    }
}

function New-Tooltip {
    param([Parameter(Mandatory = $true)][string]$Text)
    $tt = New-Object System.Windows.Controls.ToolTip
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Text
    $tb.TextWrapping = "Wrap"
    $tb.MaxWidth = 280
    $tt.Content = $tb
    $tt.Placement = $tooltipPlacement
    $tt.HorizontalOffset = $tooltipOffsetX
    $tt.VerticalOffset = $tooltipOffsetY
    $tt.Padding = $tooltipPadding
    $tt.Background = $tooltipBackground
    $tt.Foreground = $tooltipForeground
    $tt.BorderBrush = $tooltipBorderBrush
    $tt.BorderThickness = $tooltipBorderThickness
    return $tt
}

# Tooltip hiện ngay lập tức và giữ nguyên khi hover
function Set-TooltipBehavior {
    param([Parameter(Mandatory = $true)][System.Windows.DependencyObject]$Control)
    [System.Windows.Controls.ToolTipService]::SetInitialShowDelay($Control, 0)
    [System.Windows.Controls.ToolTipService]::SetBetweenShowDelay($Control, 0)
    [System.Windows.Controls.ToolTipService]::SetShowDuration($Control, [int]::MaxValue)
}

function Show-IconDialog {
    param(
        # Cho phép Owner null để tránh lỗi khi UI đã đóng nhưng timer vẫn chạy.
        [AllowNull()][System.Windows.Window]$Owner,
        # Cho phép title rỗng để tránh lỗi binding, sẽ tự fallback bên dưới.
        [AllowNull()][AllowEmptyString()][string]$Title,
        # Cho phép message rỗng để tránh lỗi binding, sẽ tự fallback bên dưới.
        [AllowNull()][AllowEmptyString()][string]$Message,
        [Parameter(Mandatory = $true)][int]$IconIndex,
        [AllowNull()][string[]]$Buttons
    )

    $s = $strings[$script:currentLang]
    if (-not $Buttons -or $Buttons.Count -eq 0) {
        $Buttons = @($s.Ok)
    }
    if ([string]::IsNullOrWhiteSpace($Title)) {
        $Title = $s.DialogDefaultTitle
    }

    $dialog = New-Object System.Windows.Window
    $dialog.Title = $Title
    $ownerVisible = $false
    if ($Owner) {
        try { $ownerVisible = $Owner.IsVisible } catch { $ownerVisible = $false }
    }
    $dialog.WindowStartupLocation = if ($ownerVisible) { "CenterOwner" } else { "CenterScreen" }
    $dialog.ResizeMode = "NoResize"
    $dialog.SizeToContent = "WidthAndHeight"
    if ($ownerVisible) {
        $dialog.Owner = $Owner
    }
    $dialog.WindowStyle = "SingleBorderWindow"

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = 20
    $row0 = New-Object System.Windows.Controls.RowDefinition
    $row0.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row0) | Out-Null
    $grid.RowDefinitions.Add($row1) | Out-Null

    $stack = New-Object System.Windows.Controls.StackPanel
    $stack.Orientation = "Horizontal"
    $stack.Margin = "0,0,0,15"

    $icon = New-Object System.Windows.Controls.Image
    $icon.Source = Get-ShellIconImage -Index $IconIndex -Size 32
    $icon.Width = 32
    $icon.Height = 32
    $icon.Margin = "0,0,12,0"

    $text = New-Object System.Windows.Controls.TextBlock
    $text.Text = $Message
    $text.TextWrapping = "Wrap"
    $text.MaxWidth = 460

    $stack.Children.Add($icon) | Out-Null
    $stack.Children.Add($text) | Out-Null

    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Center"

    foreach ($label in $Buttons) {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = $label
        $btn.MinWidth = 110
        $btn.Margin = "6,0,6,0"
        $btn.Add_Click({
            $script:dialogResult = $this.Content
            $dialog.Close()
        })
        $buttonPanel.Children.Add($btn) | Out-Null
    }

    [System.Windows.Controls.Grid]::SetRow($stack, 0)
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
    $grid.Children.Add($stack) | Out-Null
    $grid.Children.Add($buttonPanel) | Out-Null

    $dialog.Content = $grid
    $dialog.ShowDialog() | Out-Null
    return $script:dialogResult
}

function Show-GuideWindow {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Window]$Owner,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$ContentText
    )

    $win = New-Object System.Windows.Window
    $win.Title = $Title
    $win.Owner = $Owner
    $win.WindowStartupLocation = "CenterOwner"
    $win.ResizeMode = "CanResize"
    $win.Height = $Owner.ActualHeight
    $win.Width = [Math]::Max(600, [Math]::Floor($Owner.ActualWidth * 0.66))

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = 15
    $row0 = New-Object System.Windows.Controls.RowDefinition
    $row0.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row0) | Out-Null
    $grid.RowDefinitions.Add($row1) | Out-Null

    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Text = $ContentText
    $tb.TextWrapping = "Wrap"
    $tb.IsReadOnly = $true
    $tb.VerticalScrollBarVisibility = "Auto"
    $tb.VerticalAlignment = "Stretch"
    $tb.HorizontalAlignment = "Stretch"

    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = $strings[$script:currentLang].Ok
    $btn.Width = 120
    $btn.Height = 34
    $btn.Margin = "0,10,0,0"
    $btn.HorizontalAlignment = "Center"
    $btn.VerticalAlignment = "Bottom"
    $btn.Add_Click({ $win.Close() })

    [System.Windows.Controls.Grid]::SetRow($tb, 0)
    [System.Windows.Controls.Grid]::SetRow($btn, 1)

    $grid.Children.Add($tb) | Out-Null
    $grid.Children.Add($btn) | Out-Null

    $win.Content = $grid
    $win.Icon = Get-ShellIconImage -Index 24
    $win.ShowDialog() | Out-Null
}

function Show-FolderDialog {
    param([System.Windows.Window]$Owner)
    $null = $Owner
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.ShowNewFolderButton = $false
    $result = $dlg.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.SelectedPath
    }
    return $null
}

function Show-LoadingDialog {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Window]$Owner,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $script:EnableLoadingDialog) {
        return
    }

    if (-not $script:LoadingDialogCount) { $script:LoadingDialogCount = 0 }
    if ($script:LoadingDialogCount -gt 0) {
        $script:LoadingDialogCount++
        return
    }

    $script:LoadingDialogCount = 1

    $dlg = New-Object System.Windows.Window
    $dlg.WindowStyle = "None"
    $dlg.ResizeMode = "NoResize"
    $dlg.SizeToContent = "WidthAndHeight"
    $dlg.ShowInTaskbar = $false
    $dlg.Topmost = $true
    $dlg.Owner = $Owner
    $dlg.WindowStartupLocation = "CenterOwner"

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Message
    $tb.Margin = "20"
    $tb.TextWrapping = "Wrap"
    $tb.MaxWidth = 380
    $dlg.Content = $tb

    $script:LoadingDialog = $dlg
    $dlg.Show() | Out-Null

    try {
        $dlg.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    } catch { }
}

function Close-LoadingDialog {
    if (-not $script:EnableLoadingDialog) {
        return
    }
    if (-not $script:LoadingDialogCount -or $script:LoadingDialogCount -le 0) {
        return
    }
    $script:LoadingDialogCount--
    if ($script:LoadingDialogCount -gt 0) { return }

    if ($script:LoadingDialog) {
        try { $script:LoadingDialog.Close() } catch { }
    }
    $script:LoadingDialog = $null
}

function Update-ListViewColumns {
    if (-not $ListFiles -or -not $ListFiles.Columns) { return }
    if ($ListFiles.Columns.Count -lt 3) { return }

    $ListFiles.Columns[1].Width = $listColNarrowWidth
    $ListFiles.Columns[2].Width = $listColNarrowWidth
    $ListFiles.Columns[0].Width = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
}

# ==============================
# TOP BUTTONS
# ==============================
function New-TopButton {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][int]$IconIndex,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $btn = New-Object System.Windows.Controls.Button
    $btn.Name = $Name
    $btn.Margin = "4,0,4,0"
    $btn.Padding = "6"
    $btn.BorderThickness = 0 # Độ dày đường viền mặc định của button

    $stack = New-Object System.Windows.Controls.StackPanel
    $stack.Orientation = "Vertical"
    $stack.HorizontalAlignment = "Center"

    $img = New-Object System.Windows.Controls.Image
    $img.Source = Get-ShellIconImage -Index $IconIndex -Size 32
    $img.Width = 32
    $img.Height = 32
    $img.Margin = "0,0,0,4"

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Text
    $tb.HorizontalAlignment = "Center"

    $stack.Children.Add($img) | Out-Null
    $stack.Children.Add($tb) | Out-Null
    $btn.Content = $stack
    return @{ Button = $btn; TextBlock = $tb; Stack = $stack }
}

function New-SettingRow {
    param(
        [string]$Key,
        [string]$Label,
        [System.Windows.Controls.Control]$Control,
        [string]$Tooltip
    )

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "0,0,0,8"
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = 125 })) | Out-Null
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = 125 })) | Out-Null

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Label
    $tb.VerticalAlignment = "Center"

    [System.Windows.Controls.Grid]::SetColumn($tb, 0)
    [System.Windows.Controls.Grid]::SetColumn($Control, 1)

    if ($Tooltip) {
        $grid.ToolTip = New-Tooltip $Tooltip
        Set-TooltipBehavior -Control $grid
    }

    $grid.Children.Add($tb) | Out-Null
    $grid.Children.Add($Control) | Out-Null

    $rightRows[$Key] = @{ Label = $tb; Control = $Control; Container = $grid }
    return $grid
}

function New-IcoRow {
    param(
        [string]$Left,
        [string]$Right,
        [int]$LeftValue,
        [int]$RightValue
    )
    $grid = New-Object System.Windows.Controls.Grid
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
    $grid.Margin = "0,0,0,4"

    $cb1 = New-Object System.Windows.Controls.CheckBox
    $cb1.Content = $Left
    $cb2 = New-Object System.Windows.Controls.CheckBox
    $cb2.Content = $Right
    if ($LeftValue -gt 0) { $cb1.Tag = $LeftValue }
    if ($RightValue -gt 0) { $cb2.Tag = $RightValue }

    [System.Windows.Controls.Grid]::SetColumn($cb1, 0)
    [System.Windows.Controls.Grid]::SetColumn($cb2, 1)
    $grid.Children.Add($cb1) | Out-Null
    $grid.Children.Add($cb2) | Out-Null

    return @{ Grid = $grid; Left = $cb1; Right = $cb2 }
}

function Test-ValidMaxSizeText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    if ($Text -notmatch '^[0-9]+$') { return $false }
    $value = [int]$Text
    return ($value -gt 0)
}

function Test-ValidMinimizeText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $value = ConvertTo-Number $Text
    if ($null -eq $value) { return $false }
    return ($value -ge 1)
}

function Update-EncodeButtonState {
    if (-not $script:ToolsReady) {
        $btnEncode.Button.IsEnabled = $false
        return
    }

    $isValidMax = Test-ValidMaxSizeText $tbMax.Text
    $isValidMin = Test-ValidMinimizeText $tbMin.Text
    $hasFiles = $script:listItems -and $script:listItems.Count -gt 0
    $btnEncode.Button.IsEnabled = ($isValidMax -and $isValidMin -and $hasFiles)
}

# Cập nhật danh sách Quality theo định dạng output (ICO có 2 lựa chọn).
function Update-QualityOptions {
    param(
        [Parameter(Mandatory = $true)][string]$ConverterKey,
        [Parameter(Mandatory = $true)][hashtable]$Strings
    )

    if (-not $cbQuality) { return }
    $isIco = $ConverterKey -eq 'ico'
    $isLosslessOnly = $ConverterKey -in @('png', 'tiff')

    if ($isIco) {
        if ($cbQuality.ItemsSource -and $cbQuality.ItemsSource.Count -ge 5) {
            $script:LastNonIcoQualityIndex = $cbQuality.SelectedIndex
        }
        $cbQuality.ItemsSource = @($Strings.IcoQualityNormal, $Strings.IcoQualityBest)
        $sel = if ($script:LastIcoQualityIndex -in 0,1) { $script:LastIcoQualityIndex } else { 0 }
        $cbQuality.SelectedIndex = $sel
    } elseif ($isLosslessOnly) {
        $cbQuality.ItemsSource = @($Strings.Lossless)
        $cbQuality.SelectedIndex = 0
    } else {
        if ($cbQuality.ItemsSource -and $cbQuality.ItemsSource.Count -eq 2 -and ($cbQuality.ItemsSource -contains $Strings.IcoQualityNormal)) {
            $script:LastIcoQualityIndex = $cbQuality.SelectedIndex
        }
        $cbQuality.ItemsSource = @($Strings.Q100, $Strings.Q95, $Strings.Q90, $Strings.Q85, $Strings.Q80, $Strings.Q75, $Strings.Q70, $Strings.Q65, $Strings.Q60, $Strings.Q50, $Strings.Q40, $Strings.Q30, $Strings.Q20, $Strings.Q10)
        $sel = if ($script:LastNonIcoQualityIndex -ge 0 -and $script:LastNonIcoQualityIndex -lt 14) { $script:LastNonIcoQualityIndex } else { 4 }
        $cbQuality.SelectedIndex = $sel
    }
}

# ==============================
# CẬP NHẬT NGÔN NGỮ
# ==============================
function Update-Language {
    $s = $strings[$script:currentLang]
    $window.Title = $s.Title

    $btnAddFiles.TextBlock.Text = $s.AddFiles
    $btnAddFolder.TextBlock.Text = $s.AddFolder
    $btnClear.TextBlock.Text = $s.ClearList
    $btnGuide.TextBlock.Text = $s.Guide
    $btnEncode.TextBlock.Text = $s.Encode

    if ($cbLang) {
        $script:LangChanging = $true
        $cbLang.ItemsSource = @($s.LangVi, $s.LangEn)
        $cbLang.SelectedIndex = if ($script:currentLang -eq "en") { 1 } else { 0 }
        $script:LangChanging = $false
    }

    $btnAddFiles.Button.ToolTip = New-Tooltip $s.TooltipAddFiles
    $btnAddFolder.Button.ToolTip = New-Tooltip $s.TooltipAddFolder
    $btnClear.Button.ToolTip = New-Tooltip $s.TooltipClear
    $btnGuide.Button.ToolTip = New-Tooltip $s.TooltipGuide
    $btnEncode.Button.ToolTip = New-Tooltip $s.TooltipEncode

    $rightRows["converter"].Label.Text = $s.EncTo
    $rightRows["quality"].Label.Text = $s.Quality
    $rightRows["maxsize"].Label.Text = $s.MaxSize
    $rightRows["minimize"].Label.Text = $s.Minimize
    $rightRows["prefix"].Label.Text = $s.Prefix
    $rightRows["suffix"].Label.Text = $s.Suffix
    $rightRows["save"].Label.Text = $s.SaveWhere
    $rightRows["meta"].Label.Text = $s.KeepMeta
    $rightRows["run"].Label.Text = $s.RunMode

    $rightRows["converter"].Container.ToolTip = New-Tooltip $s.TooltipEncTo
    $rightRows["quality"].Container.ToolTip = New-Tooltip $s.TooltipQuality
    $rightRows["maxsize"].Container.ToolTip = New-Tooltip $s.TooltipMaxSize
    $rightRows["minimize"].Container.ToolTip = New-Tooltip $s.TooltipMinimize
    $rightRows["prefix"].Container.ToolTip = New-Tooltip $s.TooltipPrefix
    $rightRows["suffix"].Container.ToolTip = New-Tooltip $s.TooltipSuffix
    $rightRows["save"].Container.ToolTip = New-Tooltip $s.TooltipSaveWhere
    $rightRows["meta"].Container.ToolTip = New-Tooltip $s.TooltipMeta
    $rightRows["run"].Container.ToolTip = New-Tooltip $s.TooltipRun

    $icoTitle.Text = $s.IcoTitle

    # Cập nhật menu chuột phải
    if ($miRemove) { $miRemove.Header = $s.Remove }
    if ($miOpen) { $miOpen.Header = $s.OpenFolder }

    # Cập nhật tiêu đề cột ListView (DataGrid)
    if ($listFiles -and $listFiles.Columns.Count -ge 3) {
        $listFiles.Columns[0].Header = $s.StatusName
        $listFiles.Columns[1].Header = $s.StatusSize
        $listFiles.Columns[2].Header = $s.StatusBytes
    }

    # Cập nhật danh sách combobox theo ngôn ngữ
    $prevConv = $cbConverter.SelectedItem
    $prevIsOriginal = $false
    if ($prevConv) {
        $prevText = $prevConv.ToString()
        if ($prevText -eq $strings["vi"].Original -or $prevText -eq $strings["en"].Original) {
            $prevIsOriginal = $true
        }
    }
    $cbConverter.ItemsSource = @("JPG", "PNG", "WEBP", "JXL", "AVIF", "TIFF", "ICO", $s.Original)
    if ($prevIsOriginal) {
        $cbConverter.SelectedItem = $s.Original
    } elseif ($prevConv) {
        $cbConverter.SelectedItem = $prevConv
    } else {
        $cbConverter.SelectedIndex = 0
    }

    $convSelLang = $cbConverter.SelectedItem
    $convKey = if ($convSelLang) { $convSelLang.ToString().ToLowerInvariant() } else { "" }
    if ($convKey -eq $strings["vi"].Original.ToLowerInvariant() -or $convKey -eq $strings["en"].Original.ToLowerInvariant()) {
        $convKey = "original"
    }
    Update-QualityOptions -ConverterKey $convKey -Strings $s

    $saveIndex = $cbSave.SelectedIndex
    $cbSave.ItemsSource = @($s.SaveSame, $s.SaveNew)
    if ($saveIndex -ge 0) { $cbSave.SelectedIndex = $saveIndex }

    $metaIndex = $cbMeta.SelectedIndex
    $cbMeta.ItemsSource = @($s.MetaAll, $s.MetaNone)
    if ($metaIndex -ge 0) { $cbMeta.SelectedIndex = $metaIndex }

    $runIndex = $cbRun.SelectedIndex
    $cbRun.ItemsSource = @($s.Parallel, $s.Sequential)
    if ($runIndex -ge 0) { $cbRun.SelectedIndex = $runIndex }

}


function Get-EncodeFunctionMap {
    if ($script:EncodeFunctionMap) { return $script:EncodeFunctionMap }
    $map = @{}
    $scriptPath = $script:SourceScriptPath
    foreach ($cmd in (Get-Command -Type Function)) {
        if (-not $cmd.ScriptBlock) { continue }
        $file = $cmd.ScriptBlock.File
        if ($scriptPath) {
            if ($file -and ([System.IO.Path]::GetFullPath($file) -eq [System.IO.Path]::GetFullPath($scriptPath))) {
                $map[$cmd.Name] = $cmd.Definition
            }
        }
        elseif (-not $file) {
            $map[$cmd.Name] = $cmd.Definition
        }
    }
    $script:EncodeFunctionMap = $map
    return $map
}

function Start-EncodeRunspace {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][string[]]$InputFiles,
        [Parameter(Mandatory = $true)][string]$WorkingDirectory,
        [Parameter(Mandatory = $true)][System.Windows.Window]$Owner,
        [Parameter(Mandatory = $true)][System.Windows.Controls.Button]$EncodeButton,
        [Parameter(Mandatory = $true)][System.Windows.Controls.TextBox]$StatusLog,
        [Parameter(Mandatory = $true)][hashtable]$Strings,
        [Parameter(Mandatory = $true)][string]$CurrentLang
    )

    # Không cho chạy chồng nếu đang có job.
    if ($script:EncodeHandle -and -not $script:EncodeHandle.IsCompleted) {
        return $false
    }

    $EncodeButton.IsEnabled = $false
    $script:EncodeStartTime = Get-Date

    try {
        $functionMap = Get-EncodeFunctionMap
        $supportedExts = $script:supportedExtensions
        $ps = [PowerShell]::Create()

        # Nạp function nội bộ vào runspace và chạy logic trực tiếp (không dot-source file).
        $logQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $script:EncodeLogQueue = $logQueue

        [void]$ps.AddScript({
            param($innerFunctionMap, $innerSettings, $innerInputFiles, $innerWorkingDir, $innerStatusLog, $innerLogQueue, $innerLang, $innerStrings, $innerSupportedExts)

            if ($innerWorkingDir) {
                Set-Location -LiteralPath $innerWorkingDir
            }

            # Bật Information stream ở cả local lẫn script scope để Write-Information
            # trong hàm dot-sourced cũng được đẩy ra Information stream.
            $InformationPreference = 'Continue'
            $script:InformationPreference = 'Continue'

            if ($innerFunctionMap) {
                foreach ($entry in $innerFunctionMap.GetEnumerator()) {
                    Set-Item -Path ("function:{0}" -f $entry.Key) -Value ([ScriptBlock]::Create($entry.Value))
                }
            }

            if ($innerStrings) {
                $script:strings = $innerStrings
                $strings = $innerStrings
            }

            if ($innerSupportedExts) {
                $script:supportedExtensions = $innerSupportedExts
            }

            if ($innerLogQueue) {
                $script:EncodeLogQueue = $innerLogQueue
                $global:log_writer = {
                    param([string]$message)
                    if ([string]::IsNullOrWhiteSpace($message)) { return }
                    $innerLogQueue.Enqueue($message) | Out-Null
                }.GetNewClosure()
            }

            $exitCode = 0
            $script:__QHSummary = $null
            try {
                if (-not [string]::IsNullOrWhiteSpace($innerLang)) {
                    $script:currentLang = $innerLang
                }
                # Gọi logic - summary được gán vào $script:__QHSummary bên trong hàm.
                Invoke-EncodeLogic -Settings $innerSettings -InputFiles $innerInputFiles | Out-Null
            }
            catch [System.Management.Automation.ExitException] {
                $exitCode = $_.ExitCode
            }
            catch {
                $exitCode = 1
            }

            [PSCustomObject]@{
                __QHEncodeResult = $true
                ExitCode          = $exitCode
                Summary           = $script:__QHSummary
            }
        })
        [void]$ps.AddArgument($functionMap)
        [void]$ps.AddArgument($Settings)
        [void]$ps.AddArgument($InputFiles)
        [void]$ps.AddArgument($WorkingDirectory)
        [void]$ps.AddArgument($StatusLog)
        [void]$ps.AddArgument($logQueue)
        [void]$ps.AddArgument($CurrentLang)
        [void]$ps.AddArgument($Strings)
        [void]$ps.AddArgument($supportedExts)

        $script:EncodePs = $ps

        $script:EncodeHandle = $ps.BeginInvoke()

        if ($Strings -and $CurrentLang -and $Strings.ContainsKey($CurrentLang)) {
            $script:EncodeStrings = $Strings[$CurrentLang]
        }
        else {
            $script:EncodeStrings = $Strings
        }

        # Dùng DispatcherTimer để check trạng thái mà không block UI.
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(300)
        $timer.Add_Tick({
            try {
                if ($script:EncodeLogQueue) {
                    $logText = $null
                    while ($script:EncodeLogQueue.TryDequeue([ref]$logText)) {
                        if (-not [string]::IsNullOrWhiteSpace($logText)) {
                            Add-StatusLogLine -Text $logText
                        }
                    }
                }
                $uiStrings = $script:EncodeStrings
                if (-not $uiStrings) { $uiStrings = $strings[$script:currentLang] }

                # Bảo vệ khỏi race-condition: handle/ps có thể bị giải phóng trước khi tick chạy.
                if (-not $script:EncodeHandle) {
                    if ($script:EncodeTimer) { $script:EncodeTimer.Stop() }
                    if ($timer) { $timer.Stop() }
                    $script:EncodeTimer = $null
                    $timer = $null
                    return
                }
                if (-not $script:EncodeHandle.IsCompleted) {
                    return
                }
                if (-not $script:EncodePs) {
                    # Nếu PS đã bị dispose ở nơi khác, dừng timer để tránh gọi method trên null.
                    $script:EncodeHandle = $null
                    if ($script:EncodeTimer) { $script:EncodeTimer.Stop() }
                    if ($timer) { $timer.Stop() }
                    $script:EncodeTimer = $null
                    $timer = $null
                    return
                }

                if ($script:EncodeTimer) { $script:EncodeTimer.Stop() }
                if ($timer) { $timer.Stop() }
                $script:EncodeTimer = $null
                $timer = $null

                $exitCode = 0
                $summary = $null
                try {
                    $psInstance = $script:EncodePs
                    $handle = $script:EncodeHandle
                    if ($psInstance -and $handle) {
                        $result = $psInstance.EndInvoke($handle)
                        if ($result -and $result.Count -gt 0) {
                            $summaryResult = $result | Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties.Name -contains "__QHEncodeResult" } | Select-Object -Last 1
                            if ($summaryResult) {
                                $exitCode = [int]$summaryResult.ExitCode
                                $summary = $summaryResult.Summary
                            } else {
                                $exitCode = 0
                            }
                        } elseif ($psInstance.HadErrors) {
                            $exitCode = 1
                            foreach ($err in $psInstance.Streams.Error) {
                                try { Write-Host ("[ENCODE-ERR] " + $err.ToString()) } catch { }
                            }
                        }
                    }
                    else {
                        $exitCode = 1
                    }
                }
                catch {
                    $exitCode = 1
                    try { Write-Host ("[ENCODE-ERR] EndInvoke=" + $_.Exception.Message) } catch { }
                }
                finally {
                    if ($script:EncodePs) {
                        $script:EncodePs.Dispose()
                        $script:EncodePs = $null
                    }
                    $script:EncodeHandle = $null
                    if ($script:EncodeLogQueue) {
                        $logText = $null
                        while ($script:EncodeLogQueue.TryDequeue([ref]$logText)) {
                            if (-not [string]::IsNullOrWhiteSpace($logText)) {
                                Add-StatusLogLine -Text $logText
                            }
                        }
                    }
                    $script:EncodeLogQueue = $null
                }
                # Luôn bật lại nút Encode trước khi hiện thông báo.
                if (Get-Command -Name Update-EncodeButtonState -ErrorAction SilentlyContinue) {
                    Update-EncodeButtonState
                } elseif ($EncodeButton) {
                    $EncodeButton.IsEnabled = $true
                }

                if ($exitCode -ne 0) {
                    Add-StatusLogLine -Text $uiStrings.LogRunspaceError
                    return
                }

                if ($summary -and $null -ne $summary.TotalInputBytes) {
                    Add-StatusSummaryLine -Summary $summary -Strings $uiStrings
                }
            }
            catch {
                # Chặn lỗi văng lên UI thread; vẫn bật lại nút để người dùng thao tác tiếp.
                if (Get-Command -Name Update-EncodeButtonState -ErrorAction SilentlyContinue) {
                    Update-EncodeButtonState
                } elseif ($EncodeButton) {
                    $EncodeButton.IsEnabled = $true
                }
            }
        })

        $script:EncodeTimer = $timer
        $timer.Start()
        return $true
    }
    catch {
        $EncodeButton.IsEnabled = $true
        Add-StatusLogLine -Text $Strings.LogRunspaceError
        $script:EncodeLogQueue = $null
        return $false
    }

    return $true
}# UI (WPF)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Luôn giữ console để xem lỗi khi UI gặp sự cố
$script:LastUiError = $null 

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ShellIcon {
    public const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
    public const uint IMAGE_ICON = 1;
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern IntPtr LoadImage(IntPtr hInst, IntPtr lpszName, uint uType, int cxDesired, int cyDesired, uint fuLoad);
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool FreeLibrary(IntPtr hModule);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
"@

# ==============================
# THIẾT LẬP TOOLTIP CHUNG
# ==============================
$tooltipPlacement = "Mouse"
$tooltipOffsetX = 10
$tooltipOffsetY = 10
$tooltipPadding = 6
$tooltipBackground = "#FF2B2B2B"
$tooltipForeground = "#FFFFFFFF"
$tooltipBorderBrush = "#FF7A7A7A"
$tooltipBorderThickness = 1

# ==============================
# HỖ TRỢ CHẠY LOGIC NỀN (RUNSPACE)
# ==============================

# Lưu trạng thái chạy để UI không bị bấm trùng khi đang encode.
$script:EncodePs = $null
$script:EncodeHandle = $null
$script:EncodeTimer = $null


# Cache path -> index to avoid O(n) scans per update.
$script:listIndexMap = $null

$script:EnableLoadingDialog = $true
$script:StatusLog = $null
$script:StatusTotalCount = 0

# ==============================
# KHỞI TẠO TOOLS + DATA
# ==============================
$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$currentDirectory = (Get-Location).Path
$searchRoots = @($scriptDirectory, $currentDirectory) | Where-Object { $_ } | Select-Object -Unique
$magickPath = Resolve-ExecutablePath -ExeName "magick.exe" -SearchRoots $searchRoots -PreferSearchRoots

# Danh sách input hợp lệ chuẩn, dùng chung cho UI và encode.
$extensionsText = ($script:supportedExtensions -join ", ")
$invalidFileChars = ([System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object { [string]$_ }) -join " "

# ==============================
# XAML
# ==============================
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($strings[$script:currentLang].Title)" MinWidth="800" MinHeight="600"
        WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>

        <StackPanel Name="TopBar" Orientation="Horizontal" Margin="10" Grid.Row="0" HorizontalAlignment="Left" />

        <Grid Grid.Row="1" Margin="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="250" />
            </Grid.ColumnDefinitions>

            <Grid Grid.Column="0" Margin="0,0,10,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="2*" />
                    <RowDefinition Height="*" />
                </Grid.RowDefinitions>

                <DataGrid Name="ListFiles" Grid.Row="0" Margin="0,0,0,6" AllowDrop="True"
                          AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended"
                          HeadersVisibility="Column" RowHeaderWidth="0"
                          CanUserResizeColumns="False" CanUserReorderColumns="False"
                          GridLinesVisibility="All" HorizontalGridLinesBrush="#ffd6d6d6" VerticalGridLinesBrush="#FFD6D6D6"
                          ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                          EnableRowVirtualization="True" EnableColumnVirtualization="True"
                          ScrollViewer.CanContentScroll="True" MinHeight="200"
                          VirtualizingPanel.IsVirtualizing="True"
                          VirtualizingPanel.VirtualizationMode="Standard">
                    <DataGrid.Resources>
                        <Style TargetType="DataGridCell">
                            <Setter Property="BorderBrush" Value="#FFD6D6D6" />
                            <Setter Property="BorderThickness" Value="0,0,1,1" />
                            <Setter Property="Padding" Value="4,0,4,0" />
                        </Style>
                    </DataGrid.Resources>
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="$($strings[$script:currentLang].StatusName)" Binding="{Binding Name}" Width="2*">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="TextWrapping" Value="Wrap" />
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>
                        <DataGridTextColumn Header="$($strings[$script:currentLang].StatusSize)" Binding="{Binding Size}" SortMemberPath="Size" Width="50" />
                        <DataGridTextColumn Header="$($strings[$script:currentLang].StatusBytes)" Binding="{Binding Bytes}" SortMemberPath="Bytes" Width="50" />
                    </DataGrid.Columns>
                </DataGrid>

                <TextBox Name="StatusLog" Grid.Row="1"
                         TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Disabled"
                         AcceptsReturn="True"
                         IsReadOnly="True"
                         MinHeight="80" />
            </Grid>

            <ScrollViewer Grid.Column="1" VerticalScrollBarVisibility="Auto">
                <StackPanel Name="RightPanel" />
            </ScrollViewer>
        </Grid>
    </Grid>
</Window>
"@

$xml = [xml]$xaml
$reader = New-Object System.Xml.XmlNodeReader $xml
$window = [Windows.Markup.XamlReader]::Load($reader)

# ==============================
# CÁC THÀNH PHẦN UI
# ==============================
$TopBar = $window.FindName("TopBar")
$ListFiles = $window.FindName("ListFiles")
$StatusLog = $window.FindName("StatusLog")
$script:listFiles = $ListFiles
$RightPanel = $window.FindName("RightPanel")

# ==============================
# STATUS LOG (TEXTBOX)
# ==============================
$script:StatusLog = $StatusLog

# ==============================
# THIẾT LẬP CỘT DATAGRID
# ==============================
# Điều chỉnh kích thước cố định cho cột "Kích thước" và "Dung lượng", phần còn lại cho cột "Tên"
$listColNarrowWidth = 80

# Icon title
$window.Icon = Get-ShellIconImage -Index 63001

# Kích thước cửa sổ ngang = 1/2 màn hình, dọc = 2/3 màn hình
$screen = [System.Windows.SystemParameters]::WorkArea
$window.Width = [Math]::Floor($screen.Width * 0.5)
$window.Height = [Math]::Floor($screen.Height * 0.66)
$window.Left = [Math]::Floor(($screen.Width - $window.Width) / 2)
$window.Top = [Math]::Floor(($screen.Height - $window.Height) / 2)

# ==============================
# LISTVIEW DATA
# ==============================
$items = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$ListFiles.ItemsSource = $items
$listItems = $items
$script:listItems = $items
$script:listIndexMap = @{}

# Ctrl+A chọn tất cả
$ListFiles.Add_KeyDown({
    if (($_.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -and ($_.Key -eq [System.Windows.Input.Key]::A)) {
        $ListFiles.SelectAll()
        $_.Handled = $true
    }
})

# Sort số cho cột Kích thước và Dung lượng
$ListFiles.Add_Sorting({
    param($gridSender, $e)
    $sortKey = $e.Column.SortMemberPath
    if ($sortKey -ne "Size" -and $sortKey -ne "Bytes") {
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSender.ItemsSource)
        if ($view) { $view.CustomSort = $null }
        return
    }
    $e.Handled = $true

    if (-not $script:SortDirectionMap) { $script:SortDirectionMap = @{} }
    $dir = [System.ComponentModel.ListSortDirection]::Ascending
    if ($script:SortDirectionMap.ContainsKey($sortKey) -and $script:SortDirectionMap[$sortKey] -eq [System.ComponentModel.ListSortDirection]::Ascending) {
        $dir = [System.ComponentModel.ListSortDirection]::Descending
    }
    $script:SortDirectionMap[$sortKey] = $dir

    foreach ($c in $gridSender.Columns) {
        if ($c -ne $e.Column) { $c.SortDirection = $null }
    }
    $e.Column.SortDirection = $dir

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSender.ItemsSource)
    if ($view) {
        $view.SortDescriptions.Clear()
        $view.CustomSort = $null
        $directionSign = if ($dir -eq [System.ComponentModel.ListSortDirection]::Ascending) { 1 } else { -1 }
        if ($sortKey -eq "Size") {
            $view.CustomSort = New-SizeComparer -Direction $directionSign
        }
        else {
            $view.CustomSort = New-BytesComparer -Direction $directionSign
        }
    }
})

# Cập nhật chiều rộng khi đổi kích thước
$ListFiles.Add_SizeChanged({ Update-ListViewColumns })
$window.Add_ContentRendered({ Update-ListViewColumns })

$btnAddFiles = New-TopButton -Name "BtnAddFiles" -IconIndex 63009 -Text $strings[$script:currentLang].AddFiles
$btnAddFolder = New-TopButton -Name "BtnAddFolder" -IconIndex 4 -Text $strings[$script:currentLang].AddFolder
$btnClear = New-TopButton -Name "BtnClear" -IconIndex 261 -Text $strings[$script:currentLang].ClearList
$btnLang = New-TopButton -Name "BtnLang" -IconIndex 244 -Text " "
$btnGuide = New-TopButton -Name "BtnGuide" -IconIndex 24 -Text $strings[$script:currentLang].Guide
$btnEncode = New-TopButton -Name "BtnEncode" -IconIndex 16739 -Text $strings[$script:currentLang].Encode

$cbLang = New-Object System.Windows.Controls.ComboBox
$cbLang.ItemsSource = @($strings[$script:currentLang].LangVi, $strings[$script:currentLang].LangEn)
$cbLang.SelectedIndex = if ($script:currentLang -eq "en") { 1 } else { 0 }
$cbLang.HorizontalAlignment = "Center"
$cbLang.MinWidth = 90
$btnLang.Stack.Children.Remove($btnLang.TextBlock) | Out-Null
$btnLang.Stack.Children.Add($cbLang) | Out-Null

$TopBar.Children.Add($btnAddFiles.Button) | Out-Null
$TopBar.Children.Add($btnAddFolder.Button) | Out-Null
$TopBar.Children.Add($btnClear.Button) | Out-Null
$TopBar.Children.Add($btnLang.Button) | Out-Null
$TopBar.Children.Add($btnGuide.Button) | Out-Null
$TopBar.Children.Add($btnEncode.Button) | Out-Null

$btnAddFiles.Button.ToolTip = New-Tooltip $strings[$script:currentLang].TooltipAddFiles
$btnAddFolder.Button.ToolTip = New-Tooltip $strings[$script:currentLang].TooltipAddFolder
$btnClear.Button.ToolTip = New-Tooltip $strings[$script:currentLang].TooltipClear
$btnGuide.Button.ToolTip = New-Tooltip $strings[$script:currentLang].TooltipGuide
$btnEncode.Button.ToolTip = New-Tooltip $strings[$script:currentLang].TooltipEncode
Set-TooltipBehavior -Control $btnAddFiles.Button
Set-TooltipBehavior -Control $btnAddFolder.Button
Set-TooltipBehavior -Control $btnClear.Button
Set-TooltipBehavior -Control $btnGuide.Button
Set-TooltipBehavior -Control $btnEncode.Button

$cbLang.Add_SelectionChanged({
    if ($script:LangChanging) { return }
    if (-not $cbLang.SelectedItem) { return }
    $selectedText = $cbLang.SelectedItem.ToString()
    $newLang = if ($selectedText -eq "Tiếng Việt") { "vi" } else { "en" }
    if ($newLang -ne $script:currentLang) {
        $script:currentLang = $newLang
        Update-Language
    }
})

# ==============================
# RIGHT PANEL CONTROLS
# ==============================
$rightRows = @{}

# ICO options panel (ẩn)
$icoPanel = New-Object System.Windows.Controls.StackPanel
$icoPanel.Visibility = "Collapsed"
$icoPanel.Margin = "0,0,0,10"

$icoTitle = New-Object System.Windows.Controls.TextBlock
$icoTitle.Text = $strings[$script:currentLang].IcoTitle
$icoTitle.Margin = "0,0,0,6"
$icoPanel.Children.Add($icoTitle) | Out-Null

$icoRow1 = New-IcoRow -Left "16px" -Right "32px" -LeftValue 16 -RightValue 32
$icoRow2 = New-IcoRow -Left "48px" -Right "64px" -LeftValue 48 -RightValue 64
$icoRow3 = New-IcoRow -Left "128px" -Right "256px" -LeftValue 128 -RightValue 256

$icoCheckBoxes = @(
    $icoRow1.Left, $icoRow1.Right,
    $icoRow2.Left, $icoRow2.Right,
    $icoRow3.Left, $icoRow3.Right
)

$icoPanel.Children.Add($icoRow1.Grid) | Out-Null
$icoPanel.Children.Add($icoRow2.Grid) | Out-Null
$icoPanel.Children.Add($icoRow3.Grid) | Out-Null

# Row 1: converter
$cbConverter = New-Object System.Windows.Controls.ComboBox
$cbConverter.ItemsSource = @("JPG", "PNG", "WEBP", "JXL", "AVIF", "TIFF", "ICO", $strings[$script:currentLang].Original)
$cbConverter.SelectedIndex = 0
$RightPanel.Children.Add((New-SettingRow -Key "converter" -Label $strings[$script:currentLang].EncTo -Control $cbConverter -Tooltip $strings[$script:currentLang].TooltipEncTo)) | Out-Null

# Khối Ico Option (ẩn, chỉ hiện khi chọn ICO)
$RightPanel.Children.Add($icoPanel) | Out-Null

# Row 2: quality
$cbQuality = New-Object System.Windows.Controls.ComboBox
$cbQuality.ItemsSource = @(
    $strings[$script:currentLang].Q100,
    $strings[$script:currentLang].Q95,
    $strings[$script:currentLang].Q90,
    $strings[$script:currentLang].Q85,
    $strings[$script:currentLang].Q80,
    $strings[$script:currentLang].Q75,
    $strings[$script:currentLang].Q70,
    $strings[$script:currentLang].Q65,
    $strings[$script:currentLang].Q60,
    $strings[$script:currentLang].Q50,
    $strings[$script:currentLang].Q40,
    $strings[$script:currentLang].Q30,
    $strings[$script:currentLang].Q20,
    $strings[$script:currentLang].Q10
)
# Chất lượng mặc định là index 4 (80%)
$cbQuality.SelectedIndex = 4
$script:LastNonIcoQualityIndex = $cbQuality.SelectedIndex
$RightPanel.Children.Add((New-SettingRow -Key "quality" -Label $strings[$script:currentLang].Quality -Control $cbQuality -Tooltip $strings[$script:currentLang].TooltipQuality)) | Out-Null

# Row 3: Max Resolution
$tbMax = New-Object System.Windows.Controls.TextBox
$tbMax.Text = "3840"
$RightPanel.Children.Add((New-SettingRow -Key "maxsize" -Label $strings[$script:currentLang].MaxSize -Control $tbMax -Tooltip $strings[$script:currentLang].TooltipMaxSize)) | Out-Null

# Row 4: minimize
$tbMin = New-Object System.Windows.Controls.TextBox
$tbMin.Text = "1.5"
$RightPanel.Children.Add((New-SettingRow -Key "minimize" -Label $strings[$script:currentLang].Minimize -Control $tbMin -Tooltip $strings[$script:currentLang].TooltipMinimize)) | Out-Null

# Chặn nhập ký tự không hợp lệ
$tbMax.Add_PreviewTextInput({
    if ($_.Text -notmatch '^[0-9]+$') { $_.Handled = $true }
})
[System.Windows.DataObject]::AddPastingHandler($tbMax, {
    $text = $_.DataObject.GetData([System.Windows.DataFormats]::Text)
    if ($text -notmatch '^[0-9]+$') { $_.CancelCommand() }
})

$tbMin.Add_PreviewTextInput({
    if ($_.Text -notmatch '^[0-9\\.,]+$') { $_.Handled = $true }
})
[System.Windows.DataObject]::AddPastingHandler($tbMin, {
    $text = $_.DataObject.GetData([System.Windows.DataFormats]::Text)
    if ($text -notmatch '^[0-9\\.,]+$') { $_.CancelCommand() }
})

# ==============================
# KIỂM TRA GIÁ TRỊ UI (MAX/MIN)
# ==============================

# Lưu trạng thái tool để không vô tình bật Encode khi thiếu magick.
$script:ToolsReady = $true

$tbMax.Add_TextChanged({ Update-EncodeButtonState })
$tbMin.Add_TextChanged({ Update-EncodeButtonState })

# Row 5: prefix
$tbPrefix = New-Object System.Windows.Controls.TextBox
$RightPanel.Children.Add((New-SettingRow -Key "prefix" -Label $strings[$script:currentLang].Prefix -Control $tbPrefix -Tooltip $strings[$script:currentLang].TooltipPrefix)) | Out-Null

# Row 6: suffix
$tbSuffix = New-Object System.Windows.Controls.TextBox
$RightPanel.Children.Add((New-SettingRow -Key "suffix" -Label $strings[$script:currentLang].Suffix -Control $tbSuffix -Tooltip $strings[$script:currentLang].TooltipSuffix)) | Out-Null

# Row 7: create folder
$cbSave = New-Object System.Windows.Controls.ComboBox
$cbSave.ItemsSource = @($strings[$script:currentLang].SaveSame, $strings[$script:currentLang].SaveNew)
$cbSave.SelectedIndex = 1
$RightPanel.Children.Add((New-SettingRow -Key "save" -Label $strings[$script:currentLang].SaveWhere -Control $cbSave -Tooltip $strings[$script:currentLang].TooltipSaveWhere)) | Out-Null

# Row 8: metadata
$cbMeta = New-Object System.Windows.Controls.ComboBox
$cbMeta.ItemsSource = @($strings[$script:currentLang].MetaAll, $strings[$script:currentLang].MetaNone)
$cbMeta.SelectedIndex = 0
$RightPanel.Children.Add((New-SettingRow -Key "meta" -Label $strings[$script:currentLang].KeepMeta -Control $cbMeta -Tooltip $strings[$script:currentLang].TooltipMeta)) | Out-Null

# Row 9: run mode
$cbRun = New-Object System.Windows.Controls.ComboBox
$cbRun.ItemsSource = @($strings[$script:currentLang].Parallel, $strings[$script:currentLang].Sequential)
$cbRun.SelectedIndex = 0
$RightPanel.Children.Add((New-SettingRow -Key "run" -Label $strings[$script:currentLang].RunMode -Control $cbRun -Tooltip $strings[$script:currentLang].TooltipRun)) | Out-Null

# ==============================
# SỰ KIỆN & XỬ LÝ UI
# ==============================

# Hướng dẫn
$btnGuide.Button.Add_Click({
    $s = $strings[$script:currentLang]
    $guideText = $guide_content
    if ([string]::IsNullOrWhiteSpace($guideText)) {
        $guideText = $s.GuideContentUpdating
    }
    Show-GuideWindow -Owner $window -Title $s.GuideTitle -ContentText $guideText
})

# Làm trống danh sách
$btnClear.Button.Add_Click({
    $listItems.Clear()
    if ($script:listFiles) { $script:listFiles.Items.Refresh() }
    if ($script:StatusLog) {
        try { $script:StatusLog.Clear() } catch {}
    }
    $script:StatusTotalCount = 0
    Update-ListIndexMap
    Update-EncodeButtonState
})

# Menu chuột phải
$ctxMenu = New-Object System.Windows.Controls.ContextMenu
$miRemove = New-Object System.Windows.Controls.MenuItem
$miOpen = New-Object System.Windows.Controls.MenuItem
$ctxMenu.Items.Add($miRemove) | Out-Null
$ctxMenu.Items.Add($miOpen) | Out-Null
$listFiles.ContextMenu = $ctxMenu

$miRemove.Add_Click({
    $selected = @($listFiles.SelectedItems)
    if ($selected.Count -gt 0) {
        foreach ($it in $selected) {
            $listItems.Remove($it) | Out-Null
        }
        if ($script:listFiles) { $script:listFiles.Items.Refresh() }
        Update-ListIndexMap
        Update-EncodeButtonState
    }
})
$miOpen.Add_Click({
    if ($listFiles.SelectedItem) {
        $item = $listFiles.SelectedItem
        if (Test-Path -LiteralPath $item.Path) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$($item.Path)`"" | Out-Null
        }
        else {
            Add-StatusLogLine -Text $strings[$script:currentLang].MissingFileMsg
        }
    }
})

# Cập nhật nhãn menu theo ngôn ngữ
$window.Add_ContentRendered({
    $s = $strings[$script:currentLang]
    $miRemove.Header = $s.Remove
    $miOpen.Header = $s.OpenFolder
})

# Dọn worker đo kích thước khi đóng cửa sổ

# Thêm file ảnh
$btnAddFiles.Button.Add_Click({
    $s = $strings[$script:currentLang]
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Multiselect = $true
    $filterExt = ($script:supportedExtensions | ForEach-Object { "*$($_)" }) -join ';'
    $dlg.Filter = "Image Files|$filterExt|All Files|*.*"
    $ok = $dlg.ShowDialog()
    if ($ok) {
        Add-FilesToList -Files $dlg.FileNames -Strings $s
    }
})

# Thêm thư mục ảnh
$btnAddFolder.Button.Add_Click({
    $s = $strings[$script:currentLang]
    $folder = Show-FolderDialog -Owner $window
    if ($folder) {
        $choice = Show-IconDialog -Owner $window -Title $s.AddFolder -Message $s.PickScope -IconIndex 4 -Buttons @($s.OnlyParent, $s.IncludeSub)
        $includeSub = $choice -eq $s.IncludeSub
        Add-FolderToList -Folder $folder -IncludeSub $includeSub -Strings $s
    }
})

# Kéo thả file / thư mục
$listFiles.AllowDrop = $true
$listFiles.Add_PreviewDragOver({ $_.Effects = [System.Windows.DragDropEffects]::Copy; $_.Handled = $true })
$listFiles.Add_Drop({
    $s = $strings[$script:currentLang]
    $data = $_.Data.GetData([System.Windows.DataFormats]::FileDrop)
    if (-not $data) { return }

    $files = @()
    $folders = @()
    foreach ($p in $data) {
        if (Test-Path -LiteralPath $p -PathType Container) { $folders += $p } else { $files += $p }
    }

    if ($folders.Count -gt 0) {
        $choice = Show-IconDialog -Owner $window -Title $s.AddFolder -Message $s.PickScope -IconIndex 4 -Buttons @($s.OnlyParent, $s.IncludeSub)
        $includeSub = $choice -eq $s.IncludeSub
        foreach ($f in $folders) {
            Add-FolderToList -Folder $f -IncludeSub $includeSub -Strings $s
        }
    }

    if ($files.Count -gt 0) {
        Add-FilesToList -Files $files -Strings $s
    }
})

# Hiện/ẩn khối Ico Option
$cbConverter.Add_SelectionChanged({
    $selItem = $cbConverter.SelectedItem
    if (-not $selItem) { return }
    $sel = $selItem.ToString().ToLowerInvariant()
    if ($sel -eq $strings[$script:currentLang].Original -or $sel -eq $strings[$script:currentLang].Original.ToLowerInvariant()) {
        $sel = 'original'
    }
    $isIco = $sel -eq 'ico'
    $icoPanel.Visibility = if ($isIco) { 'Visible' } else { 'Collapsed' }

    # Cập nhật danh sách Quality theo converter
    Update-QualityOptions -ConverterKey $sel -Strings $strings[$script:currentLang]

})


# Encode
$btnEncode.Button.Add_Click({
    $s = $strings[$script:currentLang]
    if ($listItems.Count -eq 0) {
        return
    }

    # Thu thập settings
    $converterSel = $cbConverter.SelectedItem.ToString().ToLowerInvariant()
    if ($converterSel -eq $strings[$script:currentLang].Original.ToLowerInvariant()) { $converterSel = 'original' }

    $icoOption = "normal"
    if ($converterSel -eq 'ico') {
        $qualitySel = 100
        if ($cbQuality.SelectedIndex -eq 1) { $icoOption = "best" }
    } else {
        $qualityMap = @{
            $strings[$script:currentLang].Q100 = 100
            $strings[$script:currentLang].Q95  = 95
            $strings[$script:currentLang].Q90  = 90
            $strings[$script:currentLang].Q85  = 85
            $strings[$script:currentLang].Q80  = 80
            $strings[$script:currentLang].Q75  = 75
            $strings[$script:currentLang].Q70  = 70
            $strings[$script:currentLang].Q65  = 65
            $strings[$script:currentLang].Q60  = 60
            $strings[$script:currentLang].Q50  = 50
            $strings[$script:currentLang].Q40  = 40
            $strings[$script:currentLang].Q30  = 30
            $strings[$script:currentLang].Q20  = 20
            $strings[$script:currentLang].Q10  = 10
            $strings[$script:currentLang].Lossless = 100
        }
        $qualitySel = $qualityMap[$cbQuality.SelectedItem.ToString()]
    }

    $maxSize = [int]$tbMax.Text
    $minimize = ConvertTo-Number($tbMin.Text)

    $prefix = $tbPrefix.Text
    $suffix = $tbSuffix.Text

    if (Test-InvalidFileName -Value $prefix) {
        Add-StatusLogLine -Text ($s.InvalidChar -f $invalidFileChars)
        return
    }
    if (Test-InvalidFileName -Value $suffix) {
        Add-StatusLogLine -Text ($s.InvalidChar -f $invalidFileChars)
        return
    }

    $createFolder = $cbSave.SelectedIndex -eq 1

    $metaSel = $cbMeta.SelectedItem.ToString()
    $metadataKeep = $metaSel -eq $strings[$script:currentLang].MetaAll

    $adaptive = $cbRun.SelectedIndex -eq 0

    $icoSizes = @()
    foreach ($cb in $icoCheckBoxes) {
        if ($cb.IsChecked) { $icoSizes += $cb.Tag }
    }

    $inputFiles = $listItems | ForEach-Object { $_.Path }
    Reset-StatusLog -Strings $s -TotalCount $inputFiles.Count

    # Map kích thước đã đo ở listview để tránh probe lại.
    $inputSizeMap = @{}
    foreach ($it in $listItems) {
        if (-not $it -or -not $it.Path) { continue }
        $sz = ConvertFrom-SizeText $it.Size
        if ($sz) {
            $key = ([string]$it.Path).ToLowerInvariant()
            $inputSizeMap[$key] = $sz
        }
    }

    $settings = [ordered]@{
        converter = $converterSel
        encode_quality = $qualitySel
        max_size = $maxSize
        minimize = $minimize
        prefix = $prefix
        suffix = $suffix
        create_folder = $createFolder
        metadata_keep = $metadataKeep
        enable_adaptive_parallel = $adaptive
        ico_option = $icoOption
        input_size_map = $inputSizeMap
        ico16 = $icoSizes -contains 16
        ico32 = $icoSizes -contains 32
        ico48 = $icoSizes -contains 48
        ico64 = $icoSizes -contains 64
        ico128 = $icoSizes -contains 128
        ico256 = $icoSizes -contains 256
    }

    $workingDir = (Get-Location).Path

    # Chạy logic nền trong cùng process để tiện đóng gói 1 file.
    Start-EncodeRunspace `
        -Settings $settings `
        -InputFiles $inputFiles `
        -WorkingDirectory $workingDir `
        -Owner $window `
        -EncodeButton $btnEncode.Button `
        -StatusLog $script:StatusLog `
        -Strings $strings `
        -CurrentLang $script:currentLang
})

# Cập nhật text theo ngôn ngữ mặc định
Update-Language
Update-EncodeButtonState

# Kiểm tra công cụ cần thiết
$script:ShowToolsMissingDialog = $false
if (-not $magickPath) {
    $s = $strings[$script:currentLang]
    $script:ShowToolsMissingDialog = $true
    $btnEncode.Button.IsEnabled = $false
    $script:ToolsReady = $false
} else {
    $script:ToolsReady = $true
    Update-EncodeButtonState
}

# Mở UI
try {
    $window.Dispatcher.Add_UnhandledException({
        param($uiSender, $e)
        $null = $uiSender
        $script:LastUiError = $e.Exception
        $e.Handled = $true
    })

    if ($script:ShowToolsMissingDialog) {
        $window.Add_ContentRendered({
            if (-not $script:ShowToolsMissingDialog) { return }
            $script:ShowToolsMissingDialog = $false
            $s = $strings[$script:currentLang]
            Show-IconDialog -Owner $window -Title $s.ToolsMissingTitle -Message $s.ToolsMissingMsg -IconIndex 24 -Buttons @($s.Ok) | Out-Null
        })
    }

    $window.ShowDialog() | Out-Null
} catch {
    $script:LastUiError = $_.Exception
} finally {
}

}
