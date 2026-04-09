using System;
using System.Collections.Generic;
using Inventor;

namespace InventorDrawingPlugin.Services
{
    public class SmartReplaceService
    {
        private Inventor.Application _inventorApp;

        private class DimDNA
        {
            public string Category;
            public DimensionTypeEnum DimType;
            public Point2d TextPos; 
            public Point2d Anchor1, Anchor2, Anchor3;
            public string FormattedText; 
        }

        private class NoteDNA
        {
            public Point2d NotePos;
            public Point2d Anchor;
            public string FormattedText;
        }

        public SmartReplaceService(Inventor.Application app)
        {
            _inventorApp = app;
        }

        public void ExecuteSmartReplace(DrawingDocument dwgDoc, string newModelPath)
        {
            if (dwgDoc == null || string.IsNullOrEmpty(newModelPath)) return;

            List<DimDNA> dimBank = new List<DimDNA>();
            List<NoteDNA> noteBank = new List<NoteDNA>();

            try
            {
                // ==========================================
                // BƯỚC 1: TRÍCH XUẤT DNA VÀ XÓA (Như V1)
                // ==========================================
                foreach (Sheet sheet in dwgDoc.Sheets)
                {
                    List<DrawingDimension> dimsToDelete = new List<DrawingDimension>();

                    foreach (DrawingDimension dim in sheet.DrawingDimensions)
                    {
                        if (!dim.Attached) continue;
                        try {
                            if (dim is LinearGeneralDimension linDim) {
                                if (linDim.IntentOne?.Geometry is DrawingCurve && linDim.IntentTwo?.Geometry is DrawingCurve) {
                                    dimBank.Add(new DimDNA { Category = "Linear", DimType = linDim.DimensionType, TextPos = linDim.Text.Origin, Anchor1 = linDim.IntentOne.PointOnSheet, Anchor2 = linDim.IntentTwo.PointOnSheet, FormattedText = linDim.Text.FormattedText });
                                    dimsToDelete.Add(dim);
                                }
                            }
                            else if (dim is RadiusGeneralDimension radDim) {
                                if (radDim.Intent?.Geometry is DrawingCurve) {
                                    dimBank.Add(new DimDNA { Category = "Radius", TextPos = radDim.Text.Origin, Anchor1 = radDim.Intent.PointOnSheet, FormattedText = radDim.Text.FormattedText });
                                    dimsToDelete.Add(dim);
                                }
                            }
                            else if (dim is AngularGeneralDimension angDim) {
                                if (angDim.IntentOne?.Geometry is DrawingCurve && angDim.IntentTwo?.Geometry is DrawingCurve) {
                                    dimBank.Add(new DimDNA { Category = "Angular", TextPos = angDim.Text.Origin, Anchor1 = angDim.IntentOne.PointOnSheet, Anchor2 = angDim.IntentTwo.PointOnSheet, Anchor3 = angDim.IntentThree?.PointOnSheet, FormattedText = angDim.Text.FormattedText });
                                    dimsToDelete.Add(dim);
                                }
                            }
                            else if (dim is DiameterGeneralDimension diaDim) {
                                if (diaDim.Intent?.Geometry is DrawingCurve) {
                                    dimBank.Add(new DimDNA { Category = "Diameter", TextPos = diaDim.Text.Origin, Anchor1 = diaDim.Intent.PointOnSheet, FormattedText = diaDim.Text.FormattedText });
                                    dimsToDelete.Add(dim);
                                }
                            }
                        } catch { } 
                    }

                    List<LeaderNote> notesToDelete = new List<LeaderNote>();
                    foreach (LeaderNote note in sheet.DrawingNotes.LeaderNotes)
                    {
                        try {
                            GeometryIntent intent = note.Leader.RootNode.AttachedEntity as GeometryIntent;
                            if (intent?.Geometry is DrawingCurve && intent.PointOnSheet != null) {
                                noteBank.Add(new NoteDNA { NotePos = note.Position, Anchor = intent.PointOnSheet, FormattedText = note.FormattedText });
                                notesToDelete.Add(note);
                            }
                        } catch { }
                    }

                    foreach (var d in dimsToDelete) d.Delete();
                    foreach (var n in notesToDelete) n.Delete();
                }

                // ==========================================
                // BƯỚC 2: REPLACE MODEL 
                // ==========================================
                foreach (FileDescriptor fd in dwgDoc.File.ReferencedFileDescriptors)
                {
                    string ext = System.IO.Path.GetExtension(fd.FullFileName).ToLower();
                    if (ext == ".iam" || ext == ".ipt") { fd.ReplaceReference(newModelPath); break; }
                }
                dwgDoc.Update();

                // ==========================================
                // BƯỚC 3: TÁI SINH (QUÉT TOÀN BỘ GIẤY NHƯ V1)
                // ==========================================
                foreach (Sheet sheet in dwgDoc.Sheets)
                {
                    foreach (DimDNA d in dimBank)
                    {
                        bool success = false;
                        try {
                            DrawingCurve c1 = FindClosestCurveOnSheet(sheet, d.Anchor1);
                            DrawingCurve c2 = FindClosestCurveOnSheet(sheet, d.Anchor2);
                            DrawingCurve c3 = FindClosestCurveOnSheet(sheet, d.Anchor3);

                            if (d.Category == "Linear" && c1 != null && c2 != null) {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddLinear(d.TextPos, sheet.CreateGeometryIntent(c1, d.Anchor1), sheet.CreateGeometryIntent(c2, d.Anchor2), d.DimType);
                                newDim.Text.FormattedText = d.FormattedText; success = true;
                            }
                            else if (d.Category == "Radius" && c1 != null) {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddRadius(d.TextPos, sheet.CreateGeometryIntent(c1, d.Anchor1));
                                newDim.Text.FormattedText = d.FormattedText; success = true;
                            }
                            else if (d.Category == "Angular" && c1 != null && c2 != null) {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddAngular(d.TextPos, sheet.CreateGeometryIntent(c1, d.Anchor1), sheet.CreateGeometryIntent(c2, d.Anchor2), c3 != null ? sheet.CreateGeometryIntent(c3, d.Anchor3) : null);
                                newDim.Text.FormattedText = d.FormattedText; success = true;
                            }
                            else if (d.Category == "Diameter" && c1 != null) {
                                var newDim = sheet.DrawingDimensions.GeneralDimensions.AddDiameter(d.TextPos, sheet.CreateGeometryIntent(c1, d.Anchor1));
                                newDim.Text.FormattedText = d.FormattedText; success = true;
                            }
                        } catch { } 

                        // Nếu quét V1 thất bại, an toàn tạo kích thước đỏ
                        if (!success) SafeCreateSickDimension(sheet, d);
                    }

                    foreach (NoteDNA n in noteBank)
                    {
                        bool success = false;
                        try {
                            DrawingCurve c = FindClosestCurveOnSheet(sheet, n.Anchor);
                            if (c != null) {
                                ObjectCollection pts = _inventorApp.TransientObjects.CreateObjectCollection();
                                pts.Add(n.NotePos); pts.Add(sheet.CreateGeometryIntent(c, n.Anchor));
                                LeaderNote newNote = sheet.DrawingNotes.LeaderNotes.Add(pts, "");
                                newNote.FormattedText = n.FormattedText; success = true;
                            }
                        } catch { }
                    }
                }
            }
            catch (Exception ex) { throw new Exception("Error: " + ex.Message); }
        }

        // --- CƠ CHẾ TẠO KÍCH THƯỚC ĐỎ AN TOÀN TUYỆT ĐỐI ---
        private void SafeCreateSickDimension(Sheet sheet, DimDNA d)
        {
            DrawingSketch dummySketch = null;
            try {
                dummySketch = sheet.Sketches.Add();
                dummySketch.Edit(); // Mở Sketch
                
                GeneralDimension fakeDim = null;

                if (d.Category == "Linear") {
                    SketchPoint p1 = dummySketch.SketchPoints.Add(d.Anchor1 ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0), false);
                    SketchPoint p2 = dummySketch.SketchPoints.Add(d.Anchor2 ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0), false);
                    dummySketch.ExitEdit(); // Đóng Sketch ngay lập tức
                    
                    fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddLinear(d.TextPos, sheet.CreateGeometryIntent(p1), sheet.CreateGeometryIntent(p2), d.DimType);
                }
                else if (d.Category == "Radius" || d.Category == "Diameter") {
                    SketchCircle circ = dummySketch.SketchCircles.AddByCenterRadius(d.Anchor1 ?? _inventorApp.TransientGeometry.CreatePoint2d(0, 0), 1.0);
                    dummySketch.ExitEdit();
                    
                    if (d.Category == "Radius") fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddRadius(d.TextPos, sheet.CreateGeometryIntent(circ));
                    else fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddDiameter(d.TextPos, sheet.CreateGeometryIntent(circ));
                }
                else if (d.Category == "Angular") {
                    Point2d pt1 = _inventorApp.TransientGeometry.CreatePoint2d(d.Anchor1.X, d.Anchor1.Y + 1);
                    SketchLine L1 = dummySketch.SketchLines.AddByTwoPoints(d.Anchor1, pt1);
                    Point2d pt2 = _inventorApp.TransientGeometry.CreatePoint2d(d.Anchor2.X + 1, d.Anchor2.Y);
                    SketchLine L2 = dummySketch.SketchLines.AddByTwoPoints(d.Anchor2, pt2);
                    dummySketch.ExitEdit();
                    
                    fakeDim = (GeneralDimension)sheet.DrawingDimensions.GeneralDimensions.AddAngular(d.TextPos, sheet.CreateGeometryIntent(L1), sheet.CreateGeometryIntent(L2));
                }

                if (fakeDim != null) fakeDim.Text.FormattedText = d.FormattedText;

                // Xóa Sketch để biến kích thước thành màu đỏ
                dummySketch.Delete(); 
            } 
            catch {
                // NẾU CÓ LỖI, CỨU HỘ KHẨN CẤP ĐỂ KHÔNG BỊ KẸT SKETCH
                if (dummySketch != null) {
                    try { dummySketch.ExitEdit(); } catch { }
                    try { dummySketch.Delete(); } catch { }
                }
            }
        }

        // --- CÔNG CỤ TÌM KIẾM CỦA V1 (QUÉT TOÀN BỘ SHEET, KHÔNG QUAN TÂM TÊN VIEW) ---
        private DrawingCurve FindClosestCurveOnSheet(Sheet sheet, Point2d pt)
        {
            if (pt == null) return null;
            double minDist = 1.0; // Dung sai 1cm
            DrawingCurve bestCurve = null;

            foreach (DrawingView view in sheet.DrawingViews)
            {
                foreach (DrawingCurve c in view.DrawingCurves)
                {
                    try {
                        Box2d box = c.Evaluator2D.RangeBox;
                        // Lọc nhanh bằng RangeBox để tối ưu tốc độ
                        if (pt.X < box.MinPoint.X - minDist || pt.X > box.MaxPoint.X + minDist ||
                            pt.Y < box.MinPoint.Y - minDist || pt.Y > box.MaxPoint.Y + minDist)
                            continue;

                        c.Evaluator2D.GetParamExtents(out double min, out double max);
                        for (int i = 0; i <= 10; i++) {
                            double param = min + (max - min) * (i / 10.0);
                            double[] pArr = { param }; double[] ptArr = new double[2];
                            c.Evaluator2D.GetPointAtParam(ref pArr, ref ptArr);
                            
                            double dist = Math.Sqrt(Math.Pow(ptArr[0] - pt.X, 2) + Math.Pow(ptArr[1] - pt.Y, 2));
                            if (dist < minDist) { minDist = dist; bestCurve = c; }
                        }
                    } catch { }
                }
            }
            return bestCurve;
        }
    }
}