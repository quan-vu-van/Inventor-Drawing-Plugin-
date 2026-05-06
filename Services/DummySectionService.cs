using System;
using Inventor;

namespace InventorDrawingPlugin.Services
{
    public class DummySectionService
    {
        private Inventor.Application _inventorApp;
        private Action<string> _log;

        private const string SYM_LR = "SECTION_L-R";
        private const string SYM_RL = "SECTION_R-L";
        private const string SYM_UD = "SECTION_U-D";
        private const string SYM_DU = "SECTION_D-U";
        // Stub length = 4 × scale denominator mm (tren paper)
        // Model space: stubModel_cm = 0.4 / (viewScale × viewScale)

        private InteractionEvents _ie;
        private MouseEvents _me;
        private DrawingDocument _dwgDoc;
        private Sheet _sheet;
        private DrawingView _targetView;
        private string _sectionName;

        private enum PickStep { SelectView, Start, End, Direction }
        private PickStep _step;
        private Point2d _startPt, _endPt;
        private bool _isVertical;

        private GraphicsCoordinateSet _previewCoords;
        private bool _hasPreview;

        private SketchLine _stubLine1, _stubLine2;

        public DummySectionService(Inventor.Application app, Action<string> logger = null)
        {
            _inventorApp = app;
            _log = logger ?? (msg => { });
        }

        public void StartInteractive(DrawingDocument dwgDoc)
        {
            _dwgDoc = dwgDoc;
            _sheet = dwgDoc.ActiveSheet;
            _targetView = GetPreSelectedView(dwgDoc);
            _step = _targetView != null ? PickStep.Start : PickStep.SelectView;
            _startPt = null;
            _endPt = null;
            _hasPreview = false;

            // Neu chua co view duoc chon truoc → hien thong bao truoc khi vao interaction mode
            if (_targetView == null)
            {
                System.Windows.MessageBox.Show(
                    "Please select a VIEW on the drawing first.\n\n" +
                    "Tip: For faster workflow, select a view BEFORE clicking 'Create Dummy Section'.\n\n" +
                    "After clicking OK, click on the desired view to start.",
                    "View Required",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            }

            try
            {
                _ie = _inventorApp.CommandManager.CreateInteractionEvents();
                _ie.InteractionDisabled = false;

                _me = _ie.MouseEvents;
                _me.MouseMoveEnabled = true;
                _me.PointInferenceEnabled = true;
                _me.OnMouseClick += OnMouseClick;
                _me.OnMouseMove += OnMouseMove;

                _ie.OnTerminate += OnTerminate;

                if (_targetView == null)
                {
                    _inventorApp.StatusBarText =
                        ">>> Click on a VIEW to create section on (Right-click or Esc to cancel) <<<";
                }
                else
                {
                    _inventorApp.StatusBarText = $"SECTION ({_targetView.Name}): Click START point of section line";
                }

                _ie.Start();

                // Set cursor: select-view mode khi cho user pick view, neu khong thi mac dinh
                try
                {
                    if (_targetView == null)
                        _ie.SetCursor(CursorTypeEnum.kCursorBuiltInSelectView, null, null);
                }
                catch { }
            }
            catch (Exception ex)
            {
                _log($"Error: {ex.Message}");
                Cleanup();
            }
        }

        // ============================================================
        // MOUSE EVENTS
        // ============================================================
        private void OnMouseClick(MouseButtonEnum Button, ShiftStateEnum ShiftKeys,
                                    Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            if (Button == MouseButtonEnum.kRightMouseButton)
            {
                _inventorApp.StatusBarText = "";
                ClearPreview();
                StopInteraction();
                return;
            }
            if (Button != MouseButtonEnum.kLeftMouseButton) return;

            Point2d pt = _inventorApp.TransientGeometry.CreatePoint2d(ModelPosition.X, ModelPosition.Y);

            switch (_step)
            {
                case PickStep.SelectView:
                    _targetView = FindViewAtPoint(_sheet, pt);
                    if (_targetView == null) return;
                    _step = PickStep.Start;
                    // Doi cursor sang crosshair cho point picking
                    try { _ie.SetCursor(CursorTypeEnum.kCursorBuiltInCrosshair, null, null); } catch { }
                    _inventorApp.StatusBarText = $"SECTION ({_targetView.Name}): Click START point of section line";
                    break;

                case PickStep.Start:
                    _startPt = pt;
                    _step = PickStep.End;
                    InitPreview();
                    _inventorApp.StatusBarText = "SECTION: Click END point (auto-snap H/V)";
                    break;

                case PickStep.End:
                    _endPt = SnapHV(_startPt, pt);
                    _isVertical = Math.Abs(_endPt.Y - _startPt.Y) > Math.Abs(_endPt.X - _startPt.X);
                    _step = PickStep.Direction;
                    ClearPreview();
                    _inventorApp.StatusBarText = "SECTION: Click DIRECTION side (viewing direction)";
                    break;

                case PickStep.Direction:
                    StopInteraction();
                    string name = PromptSectionName();
                    if (!string.IsNullOrEmpty(name))
                    {
                        _sectionName = name;
                        ProcessSection(pt);
                    }
                    else
                    {
                        _inventorApp.StatusBarText = "Section cancelled.";
                    }
                    break;
            }
        }

        private void OnMouseMove(MouseButtonEnum Button, ShiftStateEnum ShiftKeys,
                                  Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            if (_step != PickStep.End || _startPt == null) return;
            double cx = ModelPosition.X, cy = ModelPosition.Y;
            double dx = Math.Abs(cx - _startPt.X), dy = Math.Abs(cy - _startPt.Y);
            UpdatePreview(_startPt.X, _startPt.Y,
                          dy >= dx ? _startPt.X : cx,
                          dy >= dx ? cy : _startPt.Y);
        }

        private void OnTerminate() { Cleanup(); }
        private void StopInteraction() { try { _ie?.Stop(); } catch { } }

        private void Cleanup()
        {
            try
            {
                if (_me != null) { _me.OnMouseClick -= OnMouseClick; _me.OnMouseMove -= OnMouseMove; _me = null; }
                if (_ie != null) { _ie.OnTerminate -= OnTerminate; _ie = null; }
            }
            catch { }
        }

        // ============================================================
        // POPUP INPUT
        // ============================================================
        private string PromptSectionName()
        {
            string result = null;
            var thread = new System.Threading.Thread(() =>
            {
                var win = new System.Windows.Window
                {
                    Title = "Section Name",
                    Width = 280,
                    Height = 130,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                    ResizeMode = System.Windows.ResizeMode.NoResize,
                    Topmost = true
                };

                var sp = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(12) };
                var lbl = new System.Windows.Controls.TextBlock { Text = "Enter section name:" };
                var txt = new System.Windows.Controls.TextBox { Text = "A1", Margin = new System.Windows.Thickness(0, 6, 0, 8) };
                var btn = new System.Windows.Controls.Button { Content = "OK", Width = 70, HorizontalAlignment = System.Windows.HorizontalAlignment.Right };
                btn.Click += (s, e) => { win.DialogResult = true; win.Close(); };
                txt.KeyDown += (s, e) => { if (e.Key == System.Windows.Input.Key.Enter) { win.DialogResult = true; win.Close(); } };

                sp.Children.Add(lbl);
                sp.Children.Add(txt);
                sp.Children.Add(btn);
                win.Content = sp;

                if (win.ShowDialog() == true)
                    result = txt.Text.Trim();
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            return result;
        }

        // ============================================================
        // PREVIEW LINE
        // ============================================================
        private void InitPreview()
        {
            try
            {
                var ig = _ie.InteractionGraphics;
                var node = ig.OverlayClientGraphics.AddNode(1);
                _previewCoords = ig.GraphicsDataSets.CreateCoordinateSet(1);
                double[] pts = { 0, 0, 0, 0, 0, 0 };
                _previewCoords.PutCoordinates(ref pts);
                var ls = node.AddLineStripGraphics();
                ls.CoordinateSet = _previewCoords;
                ls.LineType = LineTypeEnum.kDashDottedLineType;
                _hasPreview = true;
            }
            catch { _hasPreview = false; }
        }

        private void UpdatePreview(double x1, double y1, double x2, double y2)
        {
            if (!_hasPreview) return;
            try
            {
                double[] pts = { x1, y1, 0, x2, y2, 0 };
                _previewCoords.PutCoordinates(ref pts);
                _ie.InteractionGraphics.UpdateOverlayGraphics(_inventorApp.ActiveView);
            }
            catch { }
        }

        private void ClearPreview()
        {
            if (!_hasPreview) return;
            try
            {
                _ie.InteractionGraphics.OverlayClientGraphics.Delete();
                _ie.InteractionGraphics.UpdateOverlayGraphics(_inventorApp.ActiveView);
            }
            catch { }
            _hasPreview = false;
        }

        // ============================================================
        // PROCESSING
        // ============================================================
        private void ProcessSection(Point2d directionPt)
        {
            if (_startPt == null || _endPt == null) return;

            string symbolName = DetermineSymbolName(_startPt, _endPt, directionPt);

            // Auto-import symbols tu "Symbol Sections.idw" (cung thu muc .dll) neu thieu
            EnsureSymbolsImported();

            SketchedSymbolDefinition symDef = null;
            try { symDef = _dwgDoc.SketchedSymbolDefinitions[symbolName]; }
            catch
            {
                _inventorApp.StatusBarText = $"Symbol '{symbolName}' not found (check Symbol Sections.idw)";
                return;
            }

            Point2d mid1, mid2;
            DrawSectionStubs(_startPt, _endPt, out mid1, out mid2);

            bool ok1 = InsertSymbol(symDef, mid1, _stubLine1);
            bool ok2 = InsertSymbol(symDef, mid2, _stubLine2);

            if (ok1 && ok2)
            {
                _inventorApp.StatusBarText = $"Section '{_sectionName}' created successfully!";
            }
            else
            {
                _inventorApp.StatusBarText = "Warning: Some symbols could not be inserted.";
            }
        }

        // ============================================================
        // STUBS + CONSTRAINTS
        // ============================================================
        private void DrawSectionStubs(Point2d start, Point2d end, out Point2d mid1, out Point2d mid2)
        {
            mid1 = start;
            mid2 = end;
            _stubLine1 = null;
            _stubLine2 = null;

            try
            {
                DrawingSketch sk = _targetView != null
                    ? _targetView.Sketches.Add()
                    : _sheet.Sketches.Add();
                sk.Edit();

                Point2d skStart = sk.SheetToSketchSpace(start);
                Point2d skEnd = sk.SheetToSketchSpace(end);
                double dx = skEnd.X - skStart.X, dy = skEnd.Y - skStart.Y;
                double len = Math.Sqrt(dx * dx + dy * dy);
                double ux = dx / len, uy = dy / len;

                // Stub length = 4 × scale denominator mm (trong model space)
                // denominator = 1/viewScale, stub_mm = 4/viewScale, stub_cm = 0.4/viewScale
                double viewScale = _targetView != null ? _targetView.Scale : 1.0;
                double stubLen = 0.4 / viewScale;

                // Stub 1: start -> start + stubLen
                Point2d sk1End = _inventorApp.TransientGeometry.CreatePoint2d(
                    skStart.X + ux * stubLen, skStart.Y + uy * stubLen);
                _stubLine1 = sk.SketchLines.AddByTwoPoints(skStart, sk1End);
                _stubLine1.LineType = LineTypeEnum.kDashDottedLineType;
                _stubLine1.LineWeight = 0.025;

                // Stub 2: end - stubLen -> end
                Point2d sk2Start = _inventorApp.TransientGeometry.CreatePoint2d(
                    skEnd.X - ux * stubLen, skEnd.Y - uy * stubLen);
                _stubLine2 = sk.SketchLines.AddByTwoPoints(sk2Start, skEnd);
                _stubLine2.LineType = LineTypeEnum.kDashDottedLineType;
                _stubLine2.LineWeight = 0.025;

                // Midpoints -> sheet space
                Point2d skMid1 = _inventorApp.TransientGeometry.CreatePoint2d(
                    (skStart.X + sk1End.X) / 2, (skStart.Y + sk1End.Y) / 2);
                Point2d skMid2 = _inventorApp.TransientGeometry.CreatePoint2d(
                    (sk2Start.X + skEnd.X) / 2, (sk2Start.Y + skEnd.Y) / 2);
                mid1 = sk.SketchToSheetSpace(skMid1);
                mid2 = sk.SketchToSheetSpace(skMid2);

                // Dimension 50mm on each stub
                try
                {
                    Point2d dt1 = _inventorApp.TransientGeometry.CreatePoint2d(skMid1.X + uy * 1.5, skMid1.Y - ux * 1.5);
                    sk.DimensionConstraints.AddTwoPointDistance(
                        _stubLine1.StartSketchPoint, _stubLine1.EndSketchPoint,
                        DimensionOrientationEnum.kAlignedDim, dt1, false);
                    Point2d dt2 = _inventorApp.TransientGeometry.CreatePoint2d(skMid2.X + uy * 1.5, skMid2.Y - ux * 1.5);
                    sk.DimensionConstraints.AddTwoPointDistance(
                        _stubLine2.StartSketchPoint, _stubLine2.EndSketchPoint,
                        DimensionOrientationEnum.kAlignedDim, dt2, false);
                }
                catch { }

                // Collinear constraint
                try { sk.GeometricConstraints.AddCollinear((SketchEntity)_stubLine1, (SketchEntity)_stubLine2); }
                catch { }

                // Horizontal or Vertical constraint based on section direction
                try
                {
                    if (_isVertical)
                    {
                        sk.GeometricConstraints.AddVertical((SketchEntity)_stubLine1);
                        sk.GeometricConstraints.AddVertical((SketchEntity)_stubLine2);
                    }
                    else
                    {
                        sk.GeometricConstraints.AddHorizontal((SketchEntity)_stubLine1);
                        sk.GeometricConstraints.AddHorizontal((SketchEntity)_stubLine2);
                    }
                }
                catch { }

                sk.ExitEdit();
            }
            catch (Exception ex)
            {
                _log($"Warning: stub sketch error: {ex.Message}");
            }
        }

        // ============================================================
        // SYMBOL INSERT
        // ============================================================
        private bool InsertSymbol(SketchedSymbolDefinition symDef, Point2d position, SketchLine stubLine)
        {
            string[] prompts = new string[] { "", _sectionName, "" };

            // Try attached (AddWithLeader)
            if (stubLine != null)
            {
                try
                {
                    ObjectCollection lp = _inventorApp.TransientObjects.CreateObjectCollection();
                    lp.Add(_sheet.CreateGeometryIntent(stubLine, position));
                    SketchedSymbol sym = _sheet.SketchedSymbols.AddWithLeader(symDef, lp, 0.0, 1.0, prompts, false, false);
                    return true;
                }
                catch { }
            }

            // Fallback: free placement
            try
            {
                try { _sheet.SketchedSymbols.Add(symDef, position, 0.0, 1.0, prompts); }
                catch { _sheet.SketchedSymbols.Add(symDef, position, 0.0, 1.0); }
                return true;
            }
            catch { return false; }
        }

        // ============================================================
        // HELPERS
        // ============================================================
        private string DetermineSymbolName(Point2d start, Point2d end, Point2d dir)
        {
            double dx = end.X - start.X, dy = end.Y - start.Y;
            bool isVert = Math.Abs(dy) > Math.Abs(dx);
            double midX = (start.X + end.X) / 2, midY = (start.Y + end.Y) / 2;

            if (isVert)
                return dir.X > midX ? SYM_LR : SYM_RL;
            else
                return dir.Y < midY ? SYM_UD : SYM_DU;
        }

        private Point2d SnapHV(Point2d anchor, Point2d cursor)
        {
            double dx = Math.Abs(cursor.X - anchor.X), dy = Math.Abs(cursor.Y - anchor.Y);
            return dy >= dx
                ? _inventorApp.TransientGeometry.CreatePoint2d(anchor.X, cursor.Y)
                : _inventorApp.TransientGeometry.CreatePoint2d(cursor.X, anchor.Y);
        }

        // ============================================================
        // AUTO-IMPORT SYMBOLS TU "Symbol Sections.idw"
        // ============================================================
        private void EnsureSymbolsImported()
        {
            string[] required = { SYM_LR, SYM_RL, SYM_UD, SYM_DU };

            // Iterate collection va so sanh ten — robust hon indexer access
            // (indexer co the throw du symbol exists trong vai truong hop COM)
            var existingNames = new System.Collections.Generic.HashSet<string>(
                System.StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (SketchedSymbolDefinition def in _dwgDoc.SketchedSymbolDefinitions)
                {
                    try { existingNames.Add(def.Name); } catch { }
                }
            }
            catch { }

            bool allExist = true;
            foreach (var name in required)
            {
                if (!existingNames.Contains(name)) { allExist = false; break; }
            }
            if (allExist) return;

            // Tim Symbol Sections.idw o nhieu vi tri:
            // 1. Cung folder voi DLL (sau khi .bat deploy: %APPDATA%\...\Addins\<sub>\)
            // 2. Source folder (.bat khong copy .idw, can fallback ve source)
            string dllDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string[] candidates =
            {
                System.IO.Path.Combine(dllDir, "Symbol Sections.idw"),
                @"C:\CustomTools\Inventor\MCG_InventorCreateDummyDetailSection\Symbol Sections.idw"
            };

            string libFile = null;
            foreach (var path in candidates)
            {
                if (System.IO.File.Exists(path)) { libFile = path; break; }
            }

            if (libFile == null)
            {
                _inventorApp.StatusBarText = "Symbol Sections.idw not found in DLL folder or C:\\CustomTools\\Inventor\\MCG_InventorCreateDummyDetailSection\\";
                return;
            }

            // Mo file library VISIBLE de user tu copy symbols sang drawing cua minh
            // Ly do: CopyContentsTo khong copy hatch/fill - user copy UI se giu nguyen ven
            try
            {
                _inventorApp.Documents.Open(libFile, true);

                System.Windows.MessageBox.Show(
                    "Section symbols are missing from this drawing.\n\n" +
                    "Symbol Sections.idw has been opened.\n\n" +
                    "How to import symbols:\n" +
                    "1. In the opened file's Browser, expand 'Sketched Symbols'\n" +
                    "2. Right-click each symbol (SECTION_L-R, SECTION_R-L, SECTION_U-D, SECTION_D-U)\n" +
                    "3. Choose 'Copy'\n" +
                    "4. Switch back to your drawing\n" +
                    "5. Right-click 'Sketched Symbols' in your Browser → 'Paste'\n\n" +
                    "After copying, click 'Create Dummy Section' again.",
                    "Import Symbols Required",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _inventorApp.StatusBarText = $"Cannot open library: {ex.Message}";
            }
        }

        private DrawingView GetPreSelectedView(DrawingDocument dwgDoc)
        {
            try
            {
                var selSet = dwgDoc.SelectSet;
                for (int i = 1; i <= selSet.Count; i++)
                    if (selSet[i] is DrawingView dv) return dv;
            }
            catch { }
            return null;
        }

        private DrawingView FindViewAtPoint(Sheet sheet, Point2d pt)
        {
            foreach (DrawingView v in sheet.DrawingViews)
            {
                try
                {
                    double cx = v.Center.X, cy = v.Center.Y;
                    double hw = v.Width / 2, hh = v.Height / 2;
                    if (pt.X >= cx - hw && pt.X <= cx + hw && pt.Y >= cy - hh && pt.Y <= cy + hh)
                        return v;
                }
                catch { }
            }
            return null;
        }
    }
}
