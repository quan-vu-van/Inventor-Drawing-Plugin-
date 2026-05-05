using UserControl = System.Windows.Controls.UserControl;
using Inventor;

namespace MCG.Inventor.Ribbon
{
    /// <summary>
    /// Mô tả một DockableWindow gắn với tool. Chỉ những tool cần palette
    /// (ví dụ Symbol Handler) mới trả về descriptor này từ IToolDescriptor.
    /// </summary>
    public interface IDockablePanelDescriptor
    {
        /// <summary>Unique id cho DockableWindow (ví dụ "SymbolHandler.DockableWindow").</summary>
        string Id { get; }

        /// <summary>Tiêu đề hiển thị trên title bar của palette.</summary>
        string Title { get; }

        /// <summary>Vị trí neo mặc định khi tạo lần đầu.</summary>
        DockingStateEnum DefaultDockingState { get; }

        /// <summary>Kích thước tối thiểu (width, height) theo pixel.</summary>
        int MinWidth { get; }
        int MinHeight { get; }

        /// <summary>
        /// Factory tạo WPF UserControl nhúng vào DockableWindow.
        /// Gọi 1 lần khi palette được khởi tạo lần đầu.
        /// </summary>
        UserControl CreateContent();

        /// <summary>
        /// Callback sau khi content được nhúng. Dùng để wire up controllers
        /// với WPF panel (ví dụ PaletteController.SetPanel).
        /// </summary>
        void OnContentEmbedded(UserControl content);
    }
}
