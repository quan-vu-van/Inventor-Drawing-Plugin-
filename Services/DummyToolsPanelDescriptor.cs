using System.Windows.Controls;
using Inventor;
using InventorDrawingPlugin.Views;
using MCG.Inventor.Ribbon;

namespace InventorDrawingPlugin.Services
{
    /// <summary>
    /// DockableWindow descriptor cho Dummy Tools palette.
    /// Tao MainPalette WPF UserControl khi DockableWindow load.
    /// </summary>
    internal class DummyToolsPanelDescriptor : IDockablePanelDescriptor
    {
        private readonly Inventor.Application _app;

        public DummyToolsPanelDescriptor(Inventor.Application app)
        {
            _app = app;
        }

        public string Id                            => "MCGDummyTools.DummyTools.DockableWindow";
        public string Title                         => "Drawing Center";
        public DockingStateEnum DefaultDockingState => DockingStateEnum.kDockRight;
        public int MinWidth  => 320;
        public int MinHeight => 400;

        public UserControl CreateContent()
        {
            return new MainPalette(_app);
        }

        public void OnContentEmbedded(UserControl content)
        {
            // MainPalette tu wire tat ca handlers trong constructor + XAML.
            // Khong can lam them gi.
        }
    }
}
