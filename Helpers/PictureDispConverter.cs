using System.Drawing;
using System.Windows.Forms;

namespace InventorDrawingPlugin.Helpers
{
    public class PictureDispConverter : AxHost
    {
        private PictureDispConverter() : base("59EE46BA-677D-4d20-BF10-8D8067CB8B33") { }

        // Mẹo: Trả về thẳng 'object' thay vì 'IPictureDisp' để không cần thư viện stdole
        public static object ToIPictureDisp(Image image)
        {
            return GetIPictureDispFromPicture(image);
        }
    }
}