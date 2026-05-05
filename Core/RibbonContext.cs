using System;

namespace MCG.Inventor.Ribbon
{
    /// <summary>
    /// Ribbon Inventor mà tool sẽ xuất hiện. Dùng [Flags] để 1 tool có thể
    /// đăng ký ở nhiều ribbon (ví dụ utility dùng ở cả Part + Assembly).
    /// </summary>
    [Flags]
    public enum RibbonContext
    {
        None     = 0,
        Part     = 1 << 0,
        Assembly = 1 << 1,
        Drawing  = 1 << 2,

        /// <summary>Part + Assembly — môi trường 3D.</summary>
        Model3D  = Part | Assembly,

        /// <summary>Tất cả ribbon Inventor hỗ trợ.</summary>
        All      = Part | Assembly | Drawing
    }
}
