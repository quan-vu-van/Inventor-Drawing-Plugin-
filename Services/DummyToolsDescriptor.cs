using System.Drawing;
using System.Reflection;
using Inventor;
using MCG.Inventor.Ribbon;

namespace InventorDrawingPlugin.Services
{
    /// <summary>
    /// Tool descriptor cho "Dummy Tools" — button mo palette chua cac task
    /// (Smart Replace, Smart Rotate, Dummy Section, Dummy Detail).
    /// Xuat hien tren tab "MCG TOOLS" → panel "Drawing".
    /// </summary>
    internal class DummyToolsDescriptor : IToolDescriptor
    {
        private readonly DummyToolsPanelDescriptor _panel;

        public DummyToolsDescriptor(Inventor.Application app)
        {
            _panel = new DummyToolsPanelDescriptor(app);
        }

        public string Id          => "id.Button.MCGDummyTools.DummyTools";
        public string DisplayName => "Dummy\nTools";
        public string Tooltip     => "Dummy Tools";
        public string Description => "Create dummy section markers and detail boundaries on Inventor drawings";

        public Bitmap Icon16 => LoadIcon("icondata16.png");
        public Bitmap Icon32 => LoadIcon("icondata32.png");

        public PanelLocation Panel    => PanelLocation.Drawing;
        public RibbonContext Contexts => RibbonContext.Drawing;

        public ButtonDisplayEnum ButtonDisplay => ButtonDisplayEnum.kAlwaysDisplayText;
        public CommandTypesEnum  CommandType   => CommandTypesEnum.kShapeEditCmdType;
        public bool              UseLargeIcon  => true;

        public void OnExecute(NameValueMap context)
        {
            // Khi tool co DockablePanel, MCGRibbonManager tu toggle visibility.
            // Khong can xu ly o day.
        }

        public IDockablePanelDescriptor DockablePanel => _panel;

        // Lay rootNs tu typeof de match RootNamespace cua project (= "InventorPlugin").
        // KHONG dung asm.GetName().Name (= "MCG_InventorCreateDummyDetailSection").
        private static Bitmap LoadIcon(string fileName)
        {
            var asm = Assembly.GetExecutingAssembly();
            string rootNs = typeof(DummyToolsDescriptor).Namespace.Split('.')[0]; // "InventorDrawingPlugin"
            string resourceName = $"{rootNs}.Resources.{fileName}";

            // Note: csproj dung RootNamespace = "InventorPlugin", khong phai "InventorDrawingPlugin"
            // -> override truc tiep
            resourceName = $"InventorPlugin.Resources.{fileName}";
            return PictureDispConverter.LoadBitmapFromResource(asm, resourceName);
        }
    }
}
