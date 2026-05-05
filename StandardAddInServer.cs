using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Inventor;
using InventorDrawingPlugin.Services;
using MCG.Inventor.Ribbon;

namespace InventorDrawingPlugin
{
    /// <summary>
    /// Entry point cho add-in MCG_InventorCreateDummyDetailSection.
    /// Theo pattern JIT-safe: Activate → ActivateInternal [NoInlining] de exception
    /// trong dependency loading khong propagate ra Inventor (gay mat Vault/Content Centre).
    /// </summary>
    [Guid("9D0B4403-518D-40C6-98AC-CD2FF878B309")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private const string ADDIN_GUID = "{9D0B4403-518D-40C6-98AC-CD2FF878B309}";
        private const string LOG_DIR    = @"C:\CustomTools\Inventor\logs";
        private const string LOG_FILE   = @"C:\CustomTools\Inventor\logs\MCG_InventorCreateDummyDetailSection.log";

        private Inventor.Application _app;
        private MCGRibbonManager _ribbon;

        public void Activate(ApplicationAddInSite site, bool firstTime)
        {
            try
            {
                ActivateInternal(site, firstTime);
                DeleteLogFile(); // success → xoa log
            }
            catch (Exception ex)
            {
                LogToFile("Activate", ex);
                // KHONG re-throw — tranh phan via Vault/Content Centre
            }
        }

        // NoInlining BAT BUOC: JIT compile method body khi method duoc goi.
        // Neu thieu DLL phu thuoc, JIT exception van g ra TRUOC try/catch o cung method.
        // Tach sang method khac + NoInlining → exception propagate UP vao try/catch
        // cua Activate, khong thoat ra Inventor.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private void ActivateInternal(ApplicationAddInSite site, bool firstTime)
        {
            _app = site.Application;
            _ribbon = new MCGRibbonManager(_app, ADDIN_GUID);
            _ribbon.RegisterTool(new DummyToolsDescriptor(_app));
            _ribbon.Build(firstTime);
        }

        public void Deactivate()
        {
            try { _ribbon?.Cleanup(); }
            catch (Exception ex) { LogToFile("Deactivate", ex); }
            _ribbon = null;
            _app = null;
        }

        public void ExecuteCommand(int commandID) { }
        public object Automation => null;

        // ─── File logger ─────────────────────────────────────────────
        private static void LogToFile(string phase, Exception ex)
        {
            try
            {
                if (!Directory.Exists(LOG_DIR)) Directory.CreateDirectory(LOG_DIR);
                System.IO.File.AppendAllText(LOG_FILE,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {phase} ERROR: {ex.Message}\n{ex.StackTrace}\n\n");
            }
            catch { }
        }

        private static void DeleteLogFile()
        {
            try { if (System.IO.File.Exists(LOG_FILE)) System.IO.File.Delete(LOG_FILE); }
            catch { }
        }
    }
}
