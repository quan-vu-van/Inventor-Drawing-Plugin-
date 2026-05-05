using System.Drawing;
using Inventor;

namespace MCG.Inventor.Ribbon
{
    /// <summary>
    /// Mô tả 1 tool xuất hiện trên tab "MCG TOOLS".
    /// Mỗi addin implement interface này cho từng tool, sau đó đăng ký qua
    /// <see cref="MCGRibbonManager.RegisterTool"/> rồi gọi <see cref="MCGRibbonManager.Build"/>.
    /// </summary>
    public interface IToolDescriptor
    {
        // ─── Metadata ─────────────────────────────────────────────────────────

        /// <summary>Unique internal name — bắt buộc không trùng tool khác (ví dụ "id.Button.SymbolHandler").</summary>
        string Id { get; }

        /// <summary>Text hiển thị trên button. Dùng "\n" để xuống dòng.</summary>
        string DisplayName { get; }

        /// <summary>Tooltip ngắn khi hover.</summary>
        string Tooltip { get; }

        /// <summary>Mô tả đầy đủ hiển thị trong enhanced tooltip.</summary>
        string Description { get; }

        // ─── Icon ─────────────────────────────────────────────────────────────
        // Tool tự cung cấp Bitmap (thường load từ embedded resource của addin tương ứng).
        // MCGRibbonManager sẽ chuyển sang IPictureDisp rồi dispose bitmap sau khi tạo button.

        /// <summary>Icon 16×16 cho small button.</summary>
        Bitmap Icon16 { get; }

        /// <summary>Icon 32×32 cho large button.</summary>
        Bitmap Icon32 { get; }

        // ─── Placement ────────────────────────────────────────────────────────

        /// <summary>Panel nào chứa button trên tab MCG TOOLS.</summary>
        PanelLocation Panel { get; }

        /// <summary>Ribbon nào hiển thị button (Part/Assembly/Drawing, có thể combine).</summary>
        RibbonContext Contexts { get; }

        // ─── Button style ─────────────────────────────────────────────────────

        /// <summary>kAlwaysDisplayText / kLargeIconOnly / ...</summary>
        ButtonDisplayEnum ButtonDisplay { get; }

        /// <summary>kNonShapeEditCmdType / kShapeEditCmdType — liên quan Undo.</summary>
        CommandTypesEnum CommandType { get; }

        /// <summary>true = dùng icon 32px; false = 16px.</summary>
        bool UseLargeIcon { get; }

        // ─── Behavior ─────────────────────────────────────────────────────────

        /// <summary>
        /// Handler khi user click button. Nếu tool có <see cref="DockablePanel"/>,
        /// MCGRibbonManager tự toggle visibility của palette — <see cref="OnExecute"/>
        /// có thể để trống hoặc thực hiện hành động khác (vd bật/tắt mode).
        /// </summary>
        void OnExecute(NameValueMap context);

        // ─── Optional: Dockable palette ───────────────────────────────────────

        /// <summary>
        /// Descriptor cho DockableWindow nếu tool cần palette.
        /// Trả về null nếu tool chỉ là command không có palette.
        /// Khi có palette, MCGRibbonManager tự xử lý toggle visibility qua click button
        /// và đồng bộ Pressed state.
        /// </summary>
        IDockablePanelDescriptor DockablePanel { get; }
    }
}
