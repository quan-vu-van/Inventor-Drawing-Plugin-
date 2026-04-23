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
                        LogMessage($"  Sketch:   RESTORED={replaceService.CountSketchRestored} | LOST={replaceService.CountSketchLost}");
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

        // --- TASK 2: SMART ROTATE VIEW (Shelved due to Inventor API limits) ---
        private void BtnSmartRotate_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "Task 2 is PENDING.\n\n" +
                "Smart Rotate View has been shelved due to Inventor API limitations " +
                "(sketch geometry blocks view rotation, proxy system breaks annotations).",
                "Task 2 - Pending", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // --- TASK 3: DUMMY SECTION ---
        private void BtnDummySection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(_inventorApp.ActiveDocument is DrawingDocument dwgDoc))
                {
                    MessageBox.Show("Please open a drawing file.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var service = new Services.DummySectionService(_inventorApp, msg => LogMessage(msg));
                service.StartInteractive(dwgDoc);
            }
            catch (Exception ex)
            {
                LogMessage($"Error: {ex.Message}");
            }
        }
        // --- TASK 4: DUMMY DETAIL ---
        private void BtnDummyDetail_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(_inventorApp.ActiveDocument is DrawingDocument dwgDoc))
                {
                    MessageBox.Show("Please open a drawing file.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Floating dialog: name + circle/rectangle
                string detailName = null;
                bool isCircle = true;
                bool hasConnectionLine = true;

                var thread = new System.Threading.Thread(() =>
                {
                    var win = new Window
                    {
                        Title = "Dummy Detail",
                        Width = 280,
                        Height = 210,
                        WindowStartupLocation = WindowStartupLocation.CenterScreen,
                        ResizeMode = ResizeMode.NoResize,
                        Topmost = true
                    };

                    var sp = new StackPanel { Margin = new Thickness(12) };
                    sp.Children.Add(new TextBlock { Text = "Detail name:" });
                    var txt = new System.Windows.Controls.TextBox { Text = "A", Margin = new Thickness(0, 4, 0, 8) };
                    sp.Children.Add(txt);

                    var rbCircle = new System.Windows.Controls.RadioButton { Content = "Circle", IsChecked = true, Margin = new Thickness(0, 0, 12, 0) };
                    var rbRect = new System.Windows.Controls.RadioButton { Content = "Rectangle" };
                    var rbPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 6) };
                    rbPanel.Children.Add(rbCircle);
                    rbPanel.Children.Add(rbRect);
                    sp.Children.Add(rbPanel);

                    var chkConn = new System.Windows.Controls.CheckBox { Content = "Connection line", IsChecked = true, Margin = new Thickness(0, 0, 0, 10) };
                    sp.Children.Add(chkConn);

                    var btn = new System.Windows.Controls.Button { Content = "OK", Width = 70, HorizontalAlignment = HorizontalAlignment.Right };
                    btn.Click += (s2, e2) => { win.DialogResult = true; win.Close(); };
                    txt.KeyDown += (s2, e2) => { if (e2.Key == System.Windows.Input.Key.Enter) { win.DialogResult = true; win.Close(); } };
                    sp.Children.Add(btn);

                    win.Content = sp;
                    if (win.ShowDialog() == true)
                    {
                        detailName = txt.Text.Trim();
                        isCircle = rbCircle.IsChecked == true;
                        hasConnectionLine = chkConn.IsChecked == true;
                    }
                });
                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                thread.Start();
                thread.Join();

                if (string.IsNullOrEmpty(detailName)) return;

                var service = new Services.DummyDetailService(_inventorApp);
                service.StartInteractive(dwgDoc, detailName, isCircle, hasConnectionLine);
            }
            catch (Exception ex)
            {
                LogMessage($"Error: {ex.Message}");
            }
        }

        // Hàm hỗ trợ viết log ra giao diện
        public void LogMessage(string message)
        {
            TxtLog.AppendText($"\n[{DateTime.Now.ToString("HH:mm:ss")}] {message}");
            TxtLog.ScrollToEnd();
        }
    }
}