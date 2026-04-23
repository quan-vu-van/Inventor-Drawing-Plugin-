using System;
using Inventor;

namespace InventorDrawingPlugin.Services
{
    public class DummyDetailService
    {
        private Inventor.Application _inventorApp;
        private Action<string> _log;

        private InteractionEvents _ie;
        private MouseEvents _me;
        private DrawingDocument _dwgDoc;
        private Sheet _sheet;
        private DrawingView _targetView;

        private string _detailName;
        private bool _isCircle;
        private bool _hasConnectionLine;

        private enum PickStep { SelectView, Point1, Point2, LabelPos }
        private PickStep _step;
        private Point2d _pt1, _pt2, _labelPt;
        private DrawingSketch _boundarySketch;
        private GraphicsCoordinateSet _previewCoords;
        private bool _hasPreview;

        public DummyDetailService(Inventor.Application app, Action<string> logger = null)
        {
            _inventorApp = app;
            _log = logger ?? (msg => { });
        }

        public void StartInteractive(DrawingDocument dwgDoc, string detailName, bool isCircle, bool hasConnectionLine)
        {
            _dwgDoc = dwgDoc;
            _sheet = dwgDoc.ActiveSheet;
            _detailName = detailName;
            _isCircle = isCircle;
            _hasConnectionLine = hasConnectionLine;
            _targetView = GetPreSelectedView(dwgDoc);
            _step = _targetView != null ? PickStep.Point1 : PickStep.SelectView;
            _pt1 = null; _pt2 = null; _labelPt = null;
            _hasPreview = false;

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
                _inventorApp.StatusBarText = _targetView != null
                    ? (_isCircle
                        ? $"DETAIL ({_targetView.Name}): Click CENTER"
                        : $"DETAIL ({_targetView.Name}): Click CORNER 1")
                    : "DETAIL: Click on the VIEW";
                _ie.Start();
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
                    _step = PickStep.Point1;
                    _inventorApp.StatusBarText = _isCircle
                        ? $"DETAIL ({_targetView.Name}): Click CENTER"
                        : $"DETAIL ({_targetView.Name}): Click CORNER 1";
                    break;

                case PickStep.Point1:
                    _pt1 = pt;
                    _step = PickStep.Point2;
                    InitPreview();
                    _inventorApp.StatusBarText = _isCircle
                        ? "DETAIL: Click edge point (RADIUS)"
                        : "DETAIL: Click CORNER 2 (opposite)";
                    break;

                case PickStep.Point2:
                    _pt2 = pt;
                    ClearPreview();
                    CreateBoundarySketch(); // tao boundary that ngay, user thay trong khi pick label
                    _step = PickStep.LabelPos;
                    _inventorApp.StatusBarText = "DETAIL: Click LABEL position";
                    if (_hasConnectionLine)
                        InitLinePreview();
                    break;

                case PickStep.LabelPos:
                    _labelPt = pt;
                    ClearPreview();
                    StopInteraction();
                    ProcessDetail();
                    break;
            }
        }

        private void OnMouseMove(MouseButtonEnum Button, ShiftStateEnum ShiftKeys,
                                  Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            if (_step == PickStep.Point2 && _pt1 != null)
            {
                UpdateShapePreview(_pt1.X, _pt1.Y, ModelPosition.X, ModelPosition.Y);
            }
            else if (_step == PickStep.LabelPos && _hasConnectionLine && _hasPreview)
            {
                // Preview connection line tu boundary edge -> cursor
                Point2d cursor = _inventorApp.TransientGeometry.CreatePoint2d(ModelPosition.X, ModelPosition.Y);
                Point2d edgePt = CalcBoundaryEdgePoint(_pt1, _pt2, cursor);
                UpdateLinePreview(edgePt.X, edgePt.Y, cursor.X, cursor.Y);
            }
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
        // PREVIEW (circle or rectangle, real-time)
        // ============================================================
        private const int CIRCLE_SEGMENTS = 48;

        private void InitPreview()
        {
            try
            {
                var ig = _ie.InteractionGraphics;
                var node = ig.OverlayClientGraphics.AddNode(1);
                _previewCoords = ig.GraphicsDataSets.CreateCoordinateSet(1);

                // Allocate max points: circle needs CIRCLE_SEGMENTS+1, rect needs 5
                int maxPts = _isCircle ? CIRCLE_SEGMENTS + 1 : 5;
                double[] pts = new double[maxPts * 3];
                _previewCoords.PutCoordinates(ref pts);

                var ls = node.AddLineStripGraphics();
                ls.CoordinateSet = _previewCoords;
                ls.LineType = LineTypeEnum.kDashDottedLineType;
                _hasPreview = true;
            }
            catch { _hasPreview = false; }
        }

        private void UpdateShapePreview(double cx, double cy, double mx, double my)
        {
            if (!_hasPreview) return;
            try
            {
                double[] pts;

                if (_isCircle)
                {
                    double radius = Math.Sqrt((mx - cx) * (mx - cx) + (my - cy) * (my - cy));
                    pts = new double[(CIRCLE_SEGMENTS + 1) * 3];
                    for (int i = 0; i <= CIRCLE_SEGMENTS; i++)
                    {
                        double angle = 2.0 * Math.PI * i / CIRCLE_SEGMENTS;
                        pts[i * 3] = cx + radius * Math.Cos(angle);
                        pts[i * 3 + 1] = cy + radius * Math.Sin(angle);
                        pts[i * 3 + 2] = 0;
                    }
                }
                else
                {
                    // Rectangle: 5 points (closed)
                    pts = new double[15] {
                        cx, cy, 0,
                        mx, cy, 0,
                        mx, my, 0,
                        cx, my, 0,
                        cx, cy, 0
                    };
                }

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

        private void InitLinePreview()
        {
            try
            {
                var ig = _ie.InteractionGraphics;
                var node = ig.OverlayClientGraphics.AddNode(2);
                _previewCoords = ig.GraphicsDataSets.CreateCoordinateSet(2);
                double[] pts = { 0, 0, 0, 0, 0, 0 };
                _previewCoords.PutCoordinates(ref pts);
                var ls = node.AddLineStripGraphics();
                ls.CoordinateSet = _previewCoords;
                ls.LineType = LineTypeEnum.kDashDottedLineType;
                _hasPreview = true;
            }
            catch { _hasPreview = false; }
        }

        private void UpdateLinePreview(double x1, double y1, double x2, double y2)
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

        /// <summary>
        /// Tinh diem tren boundary gan label nhat.
        /// Circle: diem tren duong tron huong ve label.
        /// Rectangle: goc gan label nhat.
        /// </summary>
        private Point2d CalcBoundaryEdgePoint(Point2d pt1, Point2d pt2, Point2d labelPt)
        {
            if (_isCircle)
            {
                double radius = Math.Sqrt((pt2.X - pt1.X) * (pt2.X - pt1.X) + (pt2.Y - pt1.Y) * (pt2.Y - pt1.Y));
                double dx = labelPt.X - pt1.X, dy = labelPt.Y - pt1.Y;
                double d = Math.Sqrt(dx * dx + dy * dy);
                if (d > 0.01)
                    return _inventorApp.TransientGeometry.CreatePoint2d(
                        pt1.X + dx / d * radius, pt1.Y + dy / d * radius);
                return _inventorApp.TransientGeometry.CreatePoint2d(pt1.X + radius, pt1.Y);
            }
            else
            {
                double minX = Math.Min(pt1.X, pt2.X), maxX = Math.Max(pt1.X, pt2.X);
                double minY = Math.Min(pt1.Y, pt2.Y), maxY = Math.Max(pt1.Y, pt2.Y);
                return ClosestCorner(labelPt,
                    _inventorApp.TransientGeometry.CreatePoint2d(minX, minY),
                    _inventorApp.TransientGeometry.CreatePoint2d(maxX, minY),
                    _inventorApp.TransientGeometry.CreatePoint2d(maxX, maxY),
                    _inventorApp.TransientGeometry.CreatePoint2d(minX, maxY));
            }
        }

        // ============================================================
        // PROCESSING
        // ============================================================
        private void CreateBoundarySketch()
        {
            try
            {
                _boundarySketch = _targetView != null
                    ? _targetView.Sketches.Add()
                    : _sheet.Sketches.Add();
                _boundarySketch.Edit();

                Point2d skPt1 = _boundarySketch.SheetToSketchSpace(_pt1);
                Point2d skPt2 = _boundarySketch.SheetToSketchSpace(_pt2);

                if (_isCircle)
                {
                    double radius = Dist(skPt1, skPt2);
                    SketchCircle circ = _boundarySketch.SketchCircles.AddByCenterRadius(skPt1, radius);
                    circ.LineType = LineTypeEnum.kDashDottedLineType;
                    circ.LineWeight = 0.035;
                    try
                    {
                        Point2d dimTxt = _inventorApp.TransientGeometry.CreatePoint2d(
                            skPt1.X + radius * 0.7, skPt1.Y + radius * 0.7);
                        _boundarySketch.DimensionConstraints.AddRadius((SketchEntity)circ, dimTxt, false);
                    }
                    catch { }
                }
                else
                {
                    double minX = Math.Min(skPt1.X, skPt2.X), maxX = Math.Max(skPt1.X, skPt2.X);
                    double minY = Math.Min(skPt1.Y, skPt2.Y), maxY = Math.Max(skPt1.Y, skPt2.Y);
                    Point2d c1 = _inventorApp.TransientGeometry.CreatePoint2d(minX, minY);
                    Point2d c2 = _inventorApp.TransientGeometry.CreatePoint2d(maxX, minY);
                    Point2d c3 = _inventorApp.TransientGeometry.CreatePoint2d(maxX, maxY);
                    Point2d c4 = _inventorApp.TransientGeometry.CreatePoint2d(minX, maxY);
                    SketchLine l1 = _boundarySketch.SketchLines.AddByTwoPoints(c1, c2);
                    SketchLine l2 = _boundarySketch.SketchLines.AddByTwoPoints(c2, c3);
                    SketchLine l3 = _boundarySketch.SketchLines.AddByTwoPoints(c3, c4);
                    SketchLine l4 = _boundarySketch.SketchLines.AddByTwoPoints(c4, c1);
                    l1.LineType = LineTypeEnum.kDashDottedLineType; l1.LineWeight = 0.035;
                    l2.LineType = LineTypeEnum.kDashDottedLineType; l2.LineWeight = 0.035;
                    l3.LineType = LineTypeEnum.kDashDottedLineType; l3.LineWeight = 0.035;
                    l4.LineType = LineTypeEnum.kDashDottedLineType; l4.LineWeight = 0.035;
                    try
                    {
                        Point2d wTxt = _inventorApp.TransientGeometry.CreatePoint2d((minX + maxX) / 2, minY - 1.5);
                        _boundarySketch.DimensionConstraints.AddTwoPointDistance(
                            l1.StartSketchPoint, l1.EndSketchPoint,
                            DimensionOrientationEnum.kHorizontalDim, wTxt, false);
                        Point2d hTxt = _inventorApp.TransientGeometry.CreatePoint2d(maxX + 1.5, (minY + maxY) / 2);
                        _boundarySketch.DimensionConstraints.AddTwoPointDistance(
                            l2.StartSketchPoint, l2.EndSketchPoint,
                            DimensionOrientationEnum.kVerticalDim, hTxt, false);
                    }
                    catch { }
                }

                _boundarySketch.ExitEdit();
            }
            catch (Exception ex)
            {
                _inventorApp.StatusBarText = $"Error creating boundary: {ex.Message}";
            }
        }

        private void ProcessDetail()
        {
            if (_pt1 == null || _pt2 == null || _labelPt == null || _boundarySketch == null) return;

            try
            {
                // Mo lai sketch da tao boundary de them connection + text
                _boundarySketch.Edit();

                Point2d skPt1 = _boundarySketch.SheetToSketchSpace(_pt1);
                Point2d skPt2 = _boundarySketch.SheetToSketchSpace(_pt2);
                Point2d skLabel = _boundarySketch.SheetToSketchSpace(_labelPt);

                // Tinh boundary edge point
                Point2d boundaryEdgePt;
                if (_isCircle)
                {
                    double radius = Dist(skPt1, skPt2);
                    double dx = skLabel.X - skPt1.X, dy = skLabel.Y - skPt1.Y;
                    double d = Math.Sqrt(dx * dx + dy * dy);
                    boundaryEdgePt = d > 0.01
                        ? _inventorApp.TransientGeometry.CreatePoint2d(skPt1.X + dx / d * radius, skPt1.Y + dy / d * radius)
                        : _inventorApp.TransientGeometry.CreatePoint2d(skPt1.X + radius, skPt1.Y);
                }
                else
                {
                    double minX = Math.Min(skPt1.X, skPt2.X), maxX = Math.Max(skPt1.X, skPt2.X);
                    double minY = Math.Min(skPt1.Y, skPt2.Y), maxY = Math.Max(skPt1.Y, skPt2.Y);
                    boundaryEdgePt = ClosestCorner(skLabel,
                        _inventorApp.TransientGeometry.CreatePoint2d(minX, minY),
                        _inventorApp.TransientGeometry.CreatePoint2d(maxX, minY),
                        _inventorApp.TransientGeometry.CreatePoint2d(maxX, maxY),
                        _inventorApp.TransientGeometry.CreatePoint2d(minX, maxY));
                }

                // View scale: text size co dinh tren paper, geometry trong model space
                double viewScale = _targetView != null ? _targetView.Scale : 1.0;
                double textHeightPaper = 0.525; // 5.25mm
                double shelfLenModel = Math.Max(_detailName.Length * 0.5, 2.5) / viewScale;
                // Text offset = textHeight + gap: bottom of text chạm shelf
                double textOffsetModel = (textHeightPaper + 0.05) / viewScale;

                string textFmt = $"<StyleOverride Font='ISOCPEUR' FontSize='0.525'>{_detailName}</StyleOverride>";

                if (_hasConnectionLine)
                {
                    // Xac dinh huong shelf: extend ra xa khoi boundary
                    // Neu boundary ben TRAI labelPos -> shelf extend PHAI
                    // Neu boundary ben PHAI labelPos -> shelf extend TRAI
                    bool shelfGoesRight = boundaryEdgePt.X <= skLabel.X;

                    // Shelf length: character-based estimate trong paper space
                    // ISOCPEUR 3.5mm -> ~2.5mm/char + 0.5cm margin
                    double shelfLenPaper = _detailName.Length * 0.25 + 0.5;
                    double shelfLen = shelfLenPaper / viewScale;

                    double dir = shelfGoesRight ? 1.0 : -1.0;

                    // Connection line: boundary -> skLabel
                    SketchLine connLine = _boundarySketch.SketchLines.AddByTwoPoints(boundaryEdgePt, skLabel);
                    connLine.LineType = LineTypeEnum.kDashDottedLineType;
                    connLine.LineWeight = 0.025;

                    // Shelf: extend theo huong dir
                    Point2d shelfEnd = _inventorApp.TransientGeometry.CreatePoint2d(
                        skLabel.X + shelfLen * dir, skLabel.Y);
                    SketchLine shelf = _boundarySketch.SketchLines.AddByTwoPoints(skLabel, shelfEnd);
                    shelf.LineType = LineTypeEnum.kDashDottedLineType;
                    shelf.LineWeight = 0.025;

                    // Coincident: connection endpoint = shelf start
                    try
                    {
                        object p1 = connLine.EndSketchPoint;
                        object p2 = shelf.StartSketchPoint;
                        _boundarySketch.GeometricConstraints.AddCoincident(
                            (SketchEntity)p1, (SketchEntity)p2);
                    }
                    catch (Exception exc) { _log($"Coincident: {exc.Message}"); }

                    // Horizontal: shelf line
                    try
                    {
                        object sh = shelf;
                        _boundarySketch.GeometricConstraints.AddHorizontal((SketchEntity)sh);
                    }
                    catch (Exception exh) { _log($"Horizontal: {exh.Message}"); }

                    // Text: dat o giua shelf (tren)
                    double textCenterX = skLabel.X + (shelfLen * dir) / 2;
                    Point2d textPos = _inventorApp.TransientGeometry.CreatePoint2d(
                        textCenterX - (shelfLen * 0.4), // dich trai de text centered
                        skLabel.Y + textOffsetModel);
                    _boundarySketch.TextBoxes.AddFitted(textPos, textFmt);
                }
                else
                {
                    _boundarySketch.TextBoxes.AddFitted(skLabel, textFmt);
                }

                _boundarySketch.ExitEdit();
                _inventorApp.StatusBarText = $"Detail '{_detailName}' created successfully!";
            }
            catch (Exception ex)
            {
                _inventorApp.StatusBarText = $"Error: {ex.Message}";
            }
        }

        // ============================================================
        // HELPERS
        // ============================================================
        private double Dist(Point2d a, Point2d b)
        {
            return Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
        }

        private Point2d ClosestCorner(Point2d target, params Point2d[] corners)
        {
            Point2d best = corners[0];
            double bestDist = Dist(target, corners[0]);
            for (int i = 1; i < corners.Length; i++)
            {
                double d = Dist(target, corners[i]);
                if (d < bestDist) { bestDist = d; best = corners[i]; }
            }
            return best;
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
