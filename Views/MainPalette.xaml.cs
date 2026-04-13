using System;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Inventor;

namespace InventorDrawingPlugin.Views
{
    public partial class MainPalette : UserControl
    {
        private Inventor.Application _inventorApp;

        public MainPalette(Inventor.Application app)
        {
            InitializeComponent();
            _inventorApp = app;
        }

        // --- TASK 1: SMART REPLACE ---
        private void BtnSmartReplace_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_inventorApp.ActiveDocument is DrawingDocument dwgDoc)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Title = "Select New Model (Model B) to Replace",
                        Filter = "Inventor Assembly (*.iam)|*.iam|Inventor Part (*.ipt)|*.ipt|All Files (*.*)|*.*"
                    };

                    if (openFileDialog.ShowDialog() == true)
                    {
                        LogMessage($"[Task 1] Bat dau Smart Replace voi file:\n{openFileDialog.FileName}");

                        // Khoi tao Service voi logger truyen thang vao UI
                        var replaceService = new Services.SmartReplaceService(_inventorApp, msg => LogMessage(msg));
                        replaceService.ExecuteSmartReplace(dwgDoc, openFileDialog.FileName);

                        LogMessage($"[Task 1] Hoan tat!");
                        LogMessage($"  Dim:      XANH={replaceService.CountGreen} | DO={replaceService.CountRed} | MAT={replaceService.CountLost}");
                        LogMessage($"  Note:     XANH={replaceService.CountNoteGreen} | MAT={replaceService.CountNoteLost}");
                        LogMessage($"  HoleNote: XANH={replaceService.CountHoleNoteGreen} | DO={replaceService.CountHoleNoteRed} | MAT={replaceService.CountHoleNoteLost}");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng mở một bản vẽ (.idw hoặc .dwg) để sử dụng tính năng này.", 
                                    "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[Lỗi Task 1]: {ex.Message}");
            }
        }

        private void BtnSmartRotate_Click(object sender, RoutedEventArgs e) { LogMessage("Task 2 clicked."); }
        private void BtnTask3_Click(object sender, RoutedEventArgs e) { LogMessage("Task 3 clicked."); }
        private void BtnTask4_Click(object sender, RoutedEventArgs e) { LogMessage("Task 4 clicked."); }

        // Hàm hỗ trợ viết log ra giao diện
        public void LogMessage(string message)
        {
            TxtLog.AppendText($"\n[{DateTime.Now.ToString("HH:mm:ss")}] {message}");
            TxtLog.ScrollToEnd();
        }
    }
}