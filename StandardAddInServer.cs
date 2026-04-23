using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.IO;
using System.Reflection;
using Inventor;
using InventorDrawingPlugin.Views;     // Đã đổi namespace
using InventorDrawingPlugin.Services;  // Đã đổi namespace
using System.Windows.Forms.Integration;
using InventorDrawingPlugin.Helpers;   // Đã đổi namespace

// ĐỔI NAMESPACE TỔNG
namespace InventorDrawingPlugin 
{
    // GUID MỚI KHỚP VỚI FILE .ADDIN
    [Guid("9D0B4403-518D-40C6-98AC-CD2FF878B309")]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private Inventor.Application _inventorApp;
        private DockableWindow _dockableWindow;
        private MainPalette _paletteControl;
        // Bỏ SelectionService của PickData đi vì Add-in này chỉ tập trung vào Drawing
        private ElementHost _elementHost; 
        private ButtonDefinition _toggleButtonDef; 

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _inventorApp = addInSiteObject.Application;
            UserInterfaceManager uiMan = _inventorApp.UserInterfaceManager;

            _paletteControl = new MainPalette(_inventorApp);
            _elementHost = new ElementHost { Dock = System.Windows.Forms.DockStyle.Fill, Child = _paletteControl };

            // Tên định danh cửa sổ cũng phải đổi để không trùng với PickData
            _dockableWindow = uiMan.DockableWindows.Add(Guid.NewGuid().ToString(), "InternalName_DrawingFastTools", "Drawing Center");
            _dockableWindow.AddChild(_elementHost.Handle);
            // Dock vao canh phai voi chieu rong 320px
            _dockableWindow.DockingState = DockingStateEnum.kDockRight;
            _dockableWindow.Width = 320;
            _dockableWindow.Visible = false;

            System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(new System.Windows.Window());

            CreateRibbonButton(firstTime);
        }

        private void CreateRibbonButton(bool firstTime)
        {
            try
            {
                ControlDefinitions controlDefs = _inventorApp.CommandManager.ControlDefinitions;
                object icondata16 = null; object icondata32 = null;

                try {
                    Assembly asm = Assembly.GetExecutingAssembly();
                    using (Stream s16 = asm.GetManifestResourceStream("InventorPlugin.Resources.icondata16.png"))
                    using (Stream s32 = asm.GetManifestResourceStream("InventorPlugin.Resources.icondata32.png")) {
                        if (s16 != null && s32 != null) {
                            icondata16 = PictureDispConverter.ToIPictureDisp(new Bitmap(s16));
                            icondata32 = PictureDispConverter.ToIPictureDisp(new Bitmap(s32));
                        }
                    }
                } catch { }

                try {
                    _toggleButtonDef = (ButtonDefinition)controlDefs["InventorDrawingPlugin.ToggleButton.v4"];
                } catch {
                    _toggleButtonDef = controlDefs.AddButtonDefinition(
                        "Drawing\nTools", "InventorDrawingPlugin.ToggleButton.v4", 
                        CommandTypesEnum.kShapeEditCmdType, 
                        "{9D0B4403-518D-40C6-98AC-CD2FF878B309}", // GUID MỚI
                        "Công cụ xử lý bản vẽ", "Drawing Tools", icondata16, icondata32);
                }

                _toggleButtonDef.OnExecute += (NameValueMap Context) =>
                {
                    if (_dockableWindow != null) _dockableWindow.Visible = !_dockableWindow.Visible;
                };

                // Chỉ hiện trong môi trường Drawing
                string[] targetEnvironments = { "Drawing" };

                foreach (string envName in targetEnvironments) {
                    try {
                        Ribbon ribbon = _inventorApp.UserInterfaceManager.Ribbons[envName];
                        RibbonTab addinTab = ribbon.RibbonTabs["id_AddInsTab"];
                        RibbonPanel panel;
                        try { panel = addinTab.RibbonPanels["id_Panel_DrawingPlugin"]; }
                        catch { panel = addinTab.RibbonPanels.Add("Drawing Tools", "id_Panel_DrawingPlugin", "{9D0B4403-518D-40C6-98AC-CD2FF878B309}"); }

                        bool exists = false;
                        foreach (CommandControl ctrl in panel.CommandControls) { if (ctrl.InternalName == "InventorDrawingPlugin.ToggleButton.v4") exists = true; }
                        if (!exists) panel.CommandControls.AddButton(_toggleButtonDef, true);
                    } catch { }
                }
            } catch { }
        }

        public void Deactivate()
        {
            if (_dockableWindow != null) _dockableWindow.Delete();
            if (_elementHost != null) _elementHost.Dispose();
            Marshal.ReleaseComObject(_inventorApp);
            _inventorApp = null;
            GC.Collect();
        }

        public void ExecuteCommand(int commandID) { }
        public object Automation => null;
    }
}