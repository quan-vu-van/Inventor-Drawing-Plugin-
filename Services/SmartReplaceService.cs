using System;
using System.Collections.Generic;
using Inventor;

namespace InventorDrawingPlugin.Services
{
    public class SmartReplaceService
    {
        private Inventor.Application _inventorApp;
        private Action<string> _log;

        // --- Kết quả báo cáo ---
        public int CountGreen { get; private set; }
        public int CountRed { get; private set; }
        public int CountLost { get; private set; }
        public int CountNoteGreen { get; private set; }
        public int CountNoteLost { get; private set; }
        public int CountHoleNoteGreen { get; private set; }
        public int CountHoleNoteRed { get; private set; }
        public int CountHoleNoteLost { get; private set; }
        public int CountSketchRestored { get; private set; }
        public int CountSketchLost { get; private set; }

        // --- DNA Classes ---
        private class DimDNA
        {
            public string Category;
            public DimensionTypeEnum DimType;
            public Point2d TextPos;
            public Point2d Anchor1, Anchor2, Anchor3;
            public string FormattedText;
            public string SheetName;
        }

        private class NoteDNA
        {
            public Point2d NotePos;
            public Point2d Anchor;
            public string FormattedText;
            public string SheetName;
        }

        private class HoleNoteDNA
        {
            public Point2d NotePos;
            public Point2d Anchor;
            public string FormattedText;
            public string SheetName;
        }

        public SmartReplaceService(Inventor.Application app, Action<string> logger = null)
        {
            _inventorApp = app;
            _log = logger ?? (msg => { });
        }

        public void ExecuteSmartReplace(DrawingDocument dwgDoc, string newModelPath)
        {
            if (dwgDoc == null || string.IsNullOrEmpty(newModelPath)) return;

            // Reset counters
            CountGreen = 0; CountRed = 0; CountLost = 0;
            CountNoteGreen = 0; CountNoteLost = 0;
            CountHoleNoteGreen = 0; CountHoleNoteRed = 0; CountHoleNoteLost = 0;

            List<DimDNA> dimBank = new List<DimDNA>();
            List<NoteDNA> noteBank = new List<NoteDNA>();
            List<HoleNoteDNA> holeNoteBank = new List<HoleNoteDNA>();

            try
            {
                // ==========================================
                // BUOC 1: TRICH XUAT DNA VA XOA
                // ==========================================
                _log("--- BUOC 1: Trich xuat DNA ---");

                foreach (Sheet sheet in dwgDoc.Sheets)
                {
                    List<DrawingDimension> dimsToDelete = new List<DrawingDimension>();

                    foreach (DrawingDimension dim in sheet.DrawingDimensions)
                    {
                        if (!dim.Attached) continue;
                        try
                        {
                            if (dim is LinearGeneralDimension linDim)
                            {
                                if (linDim.IntentOne?.Geometry is DrawingCurve && linDim.IntentTwo?.Geometry is DrawingCurve)
                                {
                                    dimBank.Add(new DimDNA
                                    {
                                        Category = "Linear",
                                        DimType = linDim.DimensionType,
                                        TextPos = linDim.Text.Origin,
                                        Anchor1 = linDim.IntentOne.PointOnSheet,
                                        Anchor2 = linDim.IntentTwo.PointOnSheet,
                                        FormattedText = linDim.Text.FormattedText,
                                        SheetName = sheet.Name
                                    });
                                    dimsToDelete.Add(dim);
                                }
                            }
                            else if (dim is RadiusGeneralDimension radDim)
                            {
                                if (radDim.Intent?.Geometry is DrawingCurve)
                                {
                                    dimBank.Add(new DimDNA
                                    {
                                        Category = "Radius",
                                        TextPos = radDim.Text.Origin,
                                        Anchor1 = radDim.Intent.PointOnSheet,
                                        FormattedText = radDim.Text.FormattedText,
                                        SheetName = sheet.Name
                                    });
                                    dimsToDelete.Add(dim);
                                }
                            }
                            else if (dim is AngularGeneralDimension angDim)
                            {
                                if (angDim.IntentOne?.Geometry is DrawingCurve && angDim.IntentTwo?.Geometry is DrawingCurve)
                                {
                                    dimBank.Add(new DimDNA
                                    {
                                        Category = "Angular",
                                        TextPos = angDim.Text.Origin,
                                        Anchor1 = angDim.IntentOne.PointOnSheet,
                                        Anchor2 = angDim.IntentTwo.PointOnSheet,
                                        Anchor3 = angDim.IntentThree?.PointOnSheet,
                                        FormattedText = angDim.Text.FormattedText,
                                        SheetName = sheet.Name
                                    });
                                    dimsToDelete.Add(dim);
                                }
                            }
                            else if (dim is DiameterGeneralDimension diaDim)
                            {
                                if (diaDim.Intent?.Geometry is DrawingCurve)
                                {
                                    dimBank.Add(new DimDNA
                                    {
                                        Category = "Diameter",
                                        TextPos = diaDim.Text.Origin,
                                        Anchor1 = diaDim.Intent.PointOnSheet,
                                        FormattedText = diaDim.Text.FormattedText,
                                        SheetName = sheet.Name
                                    });
                                    dimsToDelete.Add(dim);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _log($"[WARN] Khong doc duoc dim tren Sheet '{sheet.Name}': {ex.Message}");
                        }
                    }

                    // --- Trich xuat LeaderNotes ---
                    List<LeaderNote> notesToDelete = new List<LeaderNote>();
                    foreach (LeaderNote note in sheet.DrawingNotes.LeaderNotes)
                    {
                        try
                        {
                            Point2d anchor = null;

                            // Thu lay diem bam tu nhieu loai AttachedEntity
                            object attached = note.Leader.RootNode.AttachedEntity;
                            if (attached is GeometryIntent gi && gi.PointOnSheet != null)
                            {
                                anchor = gi.PointOnSheet;
                            }
                            else if (attached is DrawingCurve dc)
                            {
                                // Curve khong co intent -> lay midpoint cua RangeBox
                                Box2d box = dc.Evaluator2D.RangeBox;
                                anchor = _inventorApp.TransientGeometry.CreatePoint2d(
                                    (box.MinPoint.X + box.MaxPoint.X) / 2,
                                    (box.MinPoint.Y + box.MaxPoint.Y) / 2);
                            }

                            // Neu khong lay duoc anchor tu entity, lay Position cua RootNode
                            if (anchor == null)
                            {
                                try
                                {
                                    anchor = note.Leader.RootNode.Position;
                                }
                                catch { }
                            }

                            // Fallback cuoi: dung chinh vi tri note
                            if (anchor == null)
                                anchor = note.Position;

                            noteBank.Add(new NoteDNA
                            {
                                NotePos = note.Position,
                                Anchor = anchor,
                                FormattedText = note.FormattedText,
                                SheetName = sheet.Name
                            });
                            notesToDelete.Add(note);
                        }
                        catch (Exception ex)
                        {
                            _log($"[WARN] Khong doc duoc note tren Sheet '{sheet.Name}': {ex.Message}");
                        }
                    }

                    // --- Trich xuat HoleThreadNotes ---
                    List<HoleThreadNote> holeNotesToDelete = new List<HoleThreadNote>();
                    foreach (HoleThreadNote hNote in sheet.DrawingNotes.HoleThreadNotes)
                    {
                        try
                        {
                            if (hNote.Intent?.Geometry is DrawingCurve)
                            {
                                holeNoteBank.Add(new HoleNoteDNA
                                {
                                    NotePos = hNote.Text.Origin,
                                    Anchor = hNote.Intent.PointOnSheet,
                                    FormattedText = hNote.FormattedHoleThreadNote,
                                    SheetName = sheet.Name
                                });
                                holeNotesToDelete.Add(hNote);
                            }
                        }
                        catch (Exception ex)
                        {
                            _log($"[WARN] Khong doc duoc HoleThreadNote tren Sheet '{sheet.Name}': {ex.Message}");
                        }
                    }

                    foreach (var d in dimsToDelete) d.Delete();
                    foreach (var n in notesToDelete) n.Delete();
                    foreach (var h in holeNotesToDelete) h.Delete();

                    _log($"Sheet '{sheet.Name}': Da trich {dimsToDelete.Count} dim, {notesToDelete.Count} note, {holeNotesToDelete.Count} hole note");
                }

                _log($"Tong DNA: {dimBank.Count} dim, {noteBank.Count} note, {holeNoteBank.Count} hole note");

                // ==========================================
                // BUOC 1.5: TRICH XUAT SKETCH CONSTRAINTS
                // ==========================================
                _log("--- BUOC 1.5: Trich xuat sketch constraints ---");
                var sketchService = new SketchConstraintService(_inventorApp, _log);
                sketchService.ExtractAndClean(dwgDoc);

                // ==========================================
                // BUOC 2: REPLACE MODEL
                // ==========================================
                _log("--- BUOC 2: Replace Model ---");

                bool replaced = false;
                foreach (FileDescriptor fd in dwgDoc.File.ReferencedFileDescriptors)
                {
                    string ext = System.IO.Path.GetExtension(fd.FullFileName).ToLower();
                    if (ext == ".iam" || ext == ".ipt")
                    {
                        _log($"Replace: {fd.FullFileName} -> {newModelPath}");
                        fd.ReplaceReference(newModelPath);
                        replaced = true;
                        break;
                    }
                }

                if (!replaced)
                {
                    _log("[ERROR] Khong tim thay file .iam/.ipt de replace!");
                    return;
                }

                dwgDoc.Update();
                _log("Model da duoc replace va update thanh cong.");

                // ==========================================
                // BUOC 3: TAI SINH
                // ==========================================
                _log("--- BUOC 3: Tai sinh kich thuoc ---");

                foreach (Sheet sheet in dwgDoc.Sheets)
                {
                    // --- Tai sinh Dimensions ---
                    foreach (DimDNA d in dimBank)
                    {
                        // Chi tai sinh dim thuoc dung sheet goc
                        if (d.SheetName != sheet.Name) continue;

                        bool success = false;
                        try
                        {
                            // Loc curve theo loai phu hop voi tung category
                            string filterForCategory = null;
                            if (d.Category == "Radius" || d.Category == "Diameter")
                                filterForCategory = "Arc";
                            else if (d.Category == "Angular")
                                filterForCategory = "Line";

                            DrawingCurve c1 = FindClosestCurveOnSheet(sheet, d.Anchor1, filterForCategory);
                            DrawingCurve c2 = FindClosestCurveOnSheet(sheet, d.Anchor2, filterForCategory);
                            DrawingCurve c3 = FindClosestCurveOnSheet(sheet, d.Anchor3, filterForCategory);

                            if (d.Category == "Linear" && c1 != null && c2 != null)
                            {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    d.TextPos,
                                    sheet.CreateGeometryIntent(c1, d.Anchor1),
                                    sheet.CreateGeometryIntent(c2, d.Anchor2),
                                    d.DimType);
                                newDim.Text.FormattedText = d.FormattedText;
                                success = true;
                            }
                            else if (d.Category == "Radius" && c1 != null)
                            {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddRadius(
                                    d.TextPos,
                                    sheet.CreateGeometryIntent(c1, d.Anchor1));
                                newDim.Text.FormattedText = d.FormattedText;
                                success = true;
                            }
                            else if (d.Category == "Angular" && c1 != null && c2 != null)
                            {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddAngular(
                                    d.TextPos,
                                    sheet.CreateGeometryIntent(c1, d.Anchor1),
                                    sheet.CreateGeometryIntent(c2, d.Anchor2),
                                    c3 != null ? sheet.CreateGeometryIntent(c3, d.Anchor3) : null);
                                newDim.Text.FormattedText = d.FormattedText;
                                success = true;
                            }
                            else if (d.Category == "Diameter" && c1 != null)
                            {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddDiameter(
                                    d.TextPos,
                                    sheet.CreateGeometryIntent(c1, d.Anchor1));
                                newDim.Text.FormattedText = d.FormattedText;
                                success = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _log($"[WARN] Tai sinh XANH that bai ({d.Category}): {ex.Message}");
                        }

                        if (success)
                        {
                            CountGreen++;
                        }
                        else
                        {
                            // Fallback: tao kich thuoc DO bang Dummy Sketch
                            bool sickOk = SafeCreateSickDimension(sheet, d);
                            if (sickOk) CountRed++;
                            else CountLost++;
                        }
                    }

                    // --- Tai sinh LeaderNotes ---
                    foreach (NoteDNA n in noteBank)
                    {
                        if (n.SheetName != sheet.Name) continue;

                        bool success = false;
                        try
                        {
                            DrawingCurve c = FindClosestCurveOnSheet(sheet, n.Anchor);
                            if (c != null)
                            {
                                ObjectCollection pts = _inventorApp.TransientObjects.CreateObjectCollection();
                                pts.Add(n.NotePos);
                                pts.Add(sheet.CreateGeometryIntent(c, n.Anchor));
                                LeaderNote newNote = sheet.DrawingNotes.LeaderNotes.Add(pts, "");
                                newNote.FormattedText = n.FormattedText;
                                success = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _log($"[WARN] Tai sinh note that bai: {ex.Message}");
                        }

                        if (success)
                        {
                            CountNoteGreen++;
                        }
                        else
                        {
                            bool sickOk = SafeCreateSickNote(sheet, n);
                            if (sickOk) CountNoteLost++;
                            else CountNoteLost++;
                        }
                    }

                    // --- Tai sinh HoleThreadNotes ---
                    foreach (HoleNoteDNA h in holeNoteBank)
                    {
                        if (h.SheetName != sheet.Name) continue;

                        bool success = false;
                        try
                        {
                            DrawingCurve c = FindClosestCurveOnSheet(sheet, h.Anchor);
                            if (c != null)
                            {
                                // HoleThreadNotes.Add(Position, HoleOrThreadEdge, LinearDiameterType, DimensionStyle)
                                HoleThreadNote newHNote = sheet.DrawingNotes.HoleThreadNotes.Add(
                                    h.NotePos, c);
                                success = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            _log($"[WARN] Tai sinh HoleThreadNote XANH that bai: {ex.Message}");
                        }

                        if (success)
                        {
                            CountHoleNoteGreen++;
                        }
                        else
                        {
                            // Fallback: HoleNote khong the tao "do" truc tiep vi can hole feature.
                            // Thay the bang LeaderNote giu nguyen text + vi tri.
                            bool sickOk = SafeCreateSickHoleNote(sheet, h);
                            if (sickOk) CountHoleNoteRed++;
                            else CountHoleNoteLost++;
                        }
                    }
                }

                // ==========================================
                // BUOC 4: KHOI PHUC SKETCH CONSTRAINTS
                // ==========================================
                _log("--- BUOC 4: Khoi phuc sketch constraints ---");
                sketchService.Restore(dwgDoc);
                CountSketchRestored = sketchService.CountConstraintRestored;
                CountSketchLost = sketchService.CountConstraintLost;

                // --- Bao cao tong ket ---
                _log("=========================================");
                _log($"KET QUA:");
                _log($"  Dim:      XANH={CountGreen}, DO={CountRed}, MAT={CountLost}");
                _log($"  Note:     XANH={CountNoteGreen}, MAT={CountNoteLost}");
                _log($"  HoleNote: XANH={CountHoleNoteGreen}, DO={CountHoleNoteRed}, MAT={CountHoleNoteLost}");
                _log($"  Sketch:   RESTORED={CountSketchRestored}, LOST={CountSketchLost}");
                int totalDim = CountGreen + CountRed + CountLost;
                int totalAll = totalDim + CountNoteGreen + CountNoteLost
                             + CountHoleNoteGreen + CountHoleNoteRed + CountHoleNoteLost
                             + CountSketchRestored + CountSketchLost;
                int totalGreen = CountGreen + CountNoteGreen + CountHoleNoteGreen + CountSketchRestored;
                if (totalAll > 0)
                {
                    _log($"  Tong: {totalGreen}/{totalAll} thanh cong ({(double)totalGreen / totalAll * 100:F1}%)");
                }
                _log("=========================================");
            }
            catch (Exception ex)
            {
                _log($"[FATAL] {ex.Message}");
                throw;
            }
        }

        // ============================================================
        // CO CHE TAO KICH THUOC DO AN TOAN TUYET DOI (Dummy Sketch)
        // ============================================================
        private bool SafeCreateSickDimension(Sheet sheet, DimDNA d)
        {
            DrawingSketch dummySketch = null;
            try
            {
                dummySketch = sheet.Sketches.Add();
                dummySketch.Edit();

                GeneralDimension fakeDim = null;

                if (d.Category == "Linear")
                {
                    // Dam bao 2 diem khong trung nhau (Inventor tu choi AddLinear neu trung)
                    Point2d pt1 = d.Anchor1 ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0);
                    Point2d pt2 = d.Anchor2 ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0);
                    double gap = Math.Sqrt(Math.Pow(pt1.X - pt2.X, 2) + Math.Pow(pt1.Y - pt2.Y, 2));
                    if (gap < 0.1) // < 1mm => day 2 diem ra xa nhau 5mm
                        pt2 = _inventorApp.TransientGeometry.CreatePoint2d(pt1.X + 0.5, pt1.Y);

                    SketchLine sLine = dummySketch.SketchLines.AddByTwoPoints(pt1, pt2);
                    dummySketch.ExitEdit();

                    fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddLinear(
                        d.TextPos,
                        sheet.CreateGeometryIntent(sLine.StartSketchPoint),
                        sheet.CreateGeometryIntent(sLine.EndSketchPoint),
                        d.DimType);
                }
                else if (d.Category == "Radius" || d.Category == "Diameter")
                {
                    SketchCircle circ = dummySketch.SketchCircles.AddByCenterRadius(
                        d.Anchor1 ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0), 1.0);
                    dummySketch.ExitEdit();

                    if (d.Category == "Radius")
                        fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddRadius(
                            d.TextPos, sheet.CreateGeometryIntent(circ));
                    else
                        fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddDiameter(
                            d.TextPos, sheet.CreateGeometryIntent(circ));
                }
                else if (d.Category == "Angular")
                {
                    Point2d pt1 = _inventorApp.TransientGeometry.CreatePoint2d(d.Anchor1.X, d.Anchor1.Y + 1);
                    SketchLine L1 = dummySketch.SketchLines.AddByTwoPoints(d.Anchor1, pt1);
                    Point2d pt2 = _inventorApp.TransientGeometry.CreatePoint2d(d.Anchor2.X + 1, d.Anchor2.Y);
                    SketchLine L2 = dummySketch.SketchLines.AddByTwoPoints(d.Anchor2, pt2);
                    dummySketch.ExitEdit();

                    fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddAngular(
                        d.TextPos,
                        sheet.CreateGeometryIntent(L1),
                        sheet.CreateGeometryIntent(L2));
                }

                if (fakeDim != null)
                    fakeDim.Text.FormattedText = d.FormattedText;

                // Xoa Sketch de bien kich thuoc thanh mau do
                dummySketch.Delete();
                _log($"  -> Dim DO ({d.Category}) tai ({d.TextPos.X:F2}, {d.TextPos.Y:F2})");
                return true;
            }
            catch (Exception ex)
            {
                _log($"[ERROR] SafeCreateSickDimension ({d.Category}): {ex.Message}");
                // Cuu ho khan cap de khong bi ket Sketch
                if (dummySketch != null)
                {
                    try { dummySketch.ExitEdit(); } catch { }
                    try { dummySketch.Delete(); } catch { }
                }
                return false;
            }
        }

        // ============================================================
        // CO CHE TAO LEADER NOTE DO (Dummy Sketch Fallback)
        // ============================================================
        private bool SafeCreateSickNote(Sheet sheet, NoteDNA n)
        {
            DrawingSketch dummySketch = null;
            try
            {
                dummySketch = sheet.Sketches.Add();
                dummySketch.Edit();

                // Tao 1 diem trong sketch lam diem bam
                SketchPoint pt = dummySketch.SketchPoints.Add(
                    n.Anchor ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0), false);
                dummySketch.ExitEdit();

                ObjectCollection pts = _inventorApp.TransientObjects.CreateObjectCollection();
                pts.Add(n.NotePos);
                pts.Add(sheet.CreateGeometryIntent(pt));

                LeaderNote newNote = sheet.DrawingNotes.LeaderNotes.Add(pts, "");
                newNote.FormattedText = n.FormattedText;

                // Xoa sketch -> note mat diem bam -> chuyen do
                dummySketch.Delete();
                _log($"  -> Note DO tai ({n.NotePos.X:F2}, {n.NotePos.Y:F2})");
                return true;
            }
            catch (Exception ex)
            {
                _log($"[ERROR] SafeCreateSickNote: {ex.Message}");
                if (dummySketch != null)
                {
                    try { dummySketch.ExitEdit(); } catch { }
                    try { dummySketch.Delete(); } catch { }
                }
                return false;
            }
        }

        // ============================================================
        // CO CHE FALLBACK CHO HOLE NOTE (Dummy Sketch -> LeaderNote)
        // HoleNote khong the tao "do" truc tiep vi API yeu cau hole feature.
        // Fallback: tao LeaderNote voi Dummy Sketch, giu nguyen text + vi tri.
        // ============================================================
        private bool SafeCreateSickHoleNote(Sheet sheet, HoleNoteDNA h)
        {
            DrawingSketch dummySketch = null;
            try
            {
                dummySketch = sheet.Sketches.Add();
                dummySketch.Edit();

                SketchPoint pt = dummySketch.SketchPoints.Add(
                    h.Anchor ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0), false);
                dummySketch.ExitEdit();

                ObjectCollection pts = _inventorApp.TransientObjects.CreateObjectCollection();
                pts.Add(h.NotePos);
                pts.Add(sheet.CreateGeometryIntent(pt));

                // Tao LeaderNote thay the (khong phai HoleNote) de giu text
                LeaderNote fallbackNote = sheet.DrawingNotes.LeaderNotes.Add(pts, "");
                fallbackNote.FormattedText = h.FormattedText;

                dummySketch.Delete();
                _log($"  -> HoleNote DO (fallback LeaderNote) tai ({h.NotePos.X:F2}, {h.NotePos.Y:F2})");
                return true;
            }
            catch (Exception ex)
            {
                _log($"[ERROR] SafeCreateSickHoleNote: {ex.Message}");
                if (dummySketch != null)
                {
                    try { dummySketch.ExitEdit(); } catch { }
                    try { dummySketch.Delete(); } catch { }
                }
                return false;
            }
        }

        // ============================================================
        // TIM CURVE GAN NHAT TREN SHEET
        // curveFilter: null = moi loai, "Arc" = circle/arc, "Line" = duong thang
        // ============================================================
        private DrawingCurve FindClosestCurveOnSheet(Sheet sheet, Point2d pt, string curveFilter = null)
        {
            if (pt == null) return null;

            const double TOLERANCE = 0.5;
            const int SAMPLE_COUNT = 25;
            const double EXACT_THRESHOLD = 0.05;

            double minDist = TOLERANCE;
            DrawingCurve bestCurve = null;

            foreach (DrawingView view in sheet.DrawingViews)
            {
                foreach (DrawingCurve c in view.DrawingCurves)
                {
                    try
                    {
                        // Loc theo loai curve truoc khi tinh toan nang
                        if (curveFilter != null)
                        {
                            CurveTypeEnum ct = c.CurveType;
                            if (curveFilter == "Arc"
                                && ct != CurveTypeEnum.kCircleCurve
                                && ct != CurveTypeEnum.kCircularArcCurve
                                && ct != CurveTypeEnum.kEllipseFullCurve
                                && ct != CurveTypeEnum.kEllipticalArcCurve)
                                continue;
                            if (curveFilter == "Line"
                                && ct != CurveTypeEnum.kLineCurve
                                && ct != CurveTypeEnum.kLineSegmentCurve)
                                continue;
                        }

                        // Loc nhanh bang RangeBox
                        Box2d box = c.Evaluator2D.RangeBox;
                        if (pt.X < box.MinPoint.X - TOLERANCE || pt.X > box.MaxPoint.X + TOLERANCE ||
                            pt.Y < box.MinPoint.Y - TOLERANCE || pt.Y > box.MaxPoint.Y + TOLERANCE)
                            continue;

                        c.Evaluator2D.GetParamExtents(out double pMin, out double pMax);

                        for (int i = 0; i <= SAMPLE_COUNT; i++)
                        {
                            double param = pMin + (pMax - pMin) * ((double)i / SAMPLE_COUNT);
                            double[] pArr = { param };
                            double[] ptArr = new double[2];
                            c.Evaluator2D.GetPointAtParam(ref pArr, ref ptArr);

                            double dist = Math.Sqrt(
                                (ptArr[0] - pt.X) * (ptArr[0] - pt.X) +
                                (ptArr[1] - pt.Y) * (ptArr[1] - pt.Y));

                            if (dist < minDist)
                            {
                                minDist = dist;
                                bestCurve = c;

                                if (dist < EXACT_THRESHOLD)
                                    return bestCurve;
                            }
                        }
                    }
                    catch { }
                }
            }

            return bestCurve;
        }
    }
}
