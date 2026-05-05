using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace MCG.Inventor.Ribbon
{
    /// <summary>
    /// Helper chuyển đổi System.Drawing.Image sang stdole.IPictureDisp
    /// để truyền vào Inventor API khi tạo ButtonDefinition icon.
    ///
    /// Kỹ thuật: kế thừa AxHost để truy cập protected static method
    /// GetIPictureDispFromPicture() của Windows Forms.
    /// </summary>
    public class PictureDispConverter : AxHost
    {
        private PictureDispConverter() : base(string.Empty) { }

        /// <summary>Chuyển Bitmap/Image sang stdole.IPictureDisp. Trả null nếu image null.</summary>
        public static stdole.IPictureDisp ToIPictureDisp(Image image)
        {
            if (image == null) return null;
            try
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PictureDispConverter] LỖI chuyển IPictureDisp: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Helper load Bitmap từ embedded resource của 1 assembly bất kỳ.
        /// Mỗi addin gọi với assembly riêng + full resource name.
        /// Ví dụ:
        ///   LoadBitmapFromResource(
        ///     Assembly.GetExecutingAssembly(),
        ///     "MCGInventorPlugin.Resources.SymbolHandler.ReplaceSymbol_16.png");
        /// </summary>
        public static Bitmap LoadBitmapFromResource(Assembly assembly, string fullResourceName)
        {
            if (assembly == null || string.IsNullOrEmpty(fullResourceName)) return null;

            using (var stream = assembly.GetManifestResourceStream(fullResourceName))
            {
                if (stream == null)
                {
                    var available = string.Join(", ", assembly.GetManifestResourceNames());
                    System.Diagnostics.Debug.WriteLine($"[PictureDispConverter] CẢNH BÁO: Không tìm thấy '{fullResourceName}'. Resources có: {available}");
                    return null;
                }
                return new Bitmap(stream);
            }
        }
    }
}
