namespace MCG.Inventor.Ribbon
{
    /// <summary>
    /// Vị trí panel trên tab "MCG TOOLS".
    /// Mỗi ribbon Inventor hiện panels khác nhau:
    ///   - Part / Assembly ribbon: Model + Utility
    ///   - Drawing ribbon        : Drawing + Utility
    /// </summary>
    public enum PanelLocation
    {
        /// <summary>Tool làm việc với Part/Assembly (3D).</summary>
        Model,

        /// <summary>Tool làm việc với Drawing (2D).</summary>
        Drawing,

        /// <summary>Tool dùng chung — có thể thuộc nhóm 3D hoặc Drawing tùy Contexts.</summary>
        Utility
    }
}
