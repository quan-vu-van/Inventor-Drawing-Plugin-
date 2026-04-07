using System;
using System.Collections.Generic;
using Inventor;

namespace InventorDrawingPlugin.Services
{
    public class SmartReplaceService
    {
        private Inventor.Application _inventorApp;

        private class DimData
        {
            public Point2d TextPos;
            public DimensionTypeEnum DimType;
            public Point2d Anchor1;
            public Point2d Anchor2;
            public string FormattedText; 
        }

        private class NoteData
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

            List<DimData> dimBank = new List<DimData>();
            List<NoteData> noteBank = new List<NoteData>();

            try
            {
                // ==========================================================
                // BƯỚC 1: TRÍCH XUẤT TỌA ĐỘ 2D (CHỈ LINEAR VÀ NOTE)
                // ==========================================================
                foreach (Sheet sheet in dwgDoc.Sheets)
                {
                    List<DrawingDimension> dimsToDelete = new List<DrawingDimension>();

                    foreach (DrawingDimension dim in sheet.DrawingDimensions)
                    {
                        if (dim is LinearGeneralDimension linDim && linDim.Attached)
                        {
                            try {
                                if (linDim.IntentOne?.Geometry is DrawingCurve && linDim.IntentTwo?.Geometry is DrawingCurve)
                                {
                                    dimBank.Add(new DimData {
                                        TextPos = linDim.Text.Origin,
                                        DimType = linDim.DimensionType,
                                        Anchor1 = linDim.IntentOne.PointOnSheet,
                                        Anchor2 = linDim.IntentTwo.PointOnSheet,
                                        FormattedText = linDim.Text.FormattedText
                                    });
                                    dimsToDelete.Add(dim); 
                                }
                            } catch { }
                        }
                    }

                    List<LeaderNote> notesToDelete = new List<LeaderNote>();
                    foreach (LeaderNote note in sheet.DrawingNotes.LeaderNotes)
                    {
                        try {
                            GeometryIntent intent = note.Leader.RootNode.AttachedEntity as GeometryIntent;
                            if (intent?.Geometry is DrawingCurve && intent.PointOnSheet != null)
                            {
                                noteBank.Add(new NoteData {
                                    NotePos = note.Position,
                                    Anchor = intent.PointOnSheet,
                                    FormattedText = note.FormattedText
                                });
                                notesToDelete.Add(note);
                            }
                        } catch { }
                    }

                    foreach (var d in dimsToDelete) d.Delete();
                    foreach (var n in notesToDelete) n.Delete();
                }

                // ==========================================================
                // BƯỚC 2: REPLACE MODEL VÀ CẬP NHẬT 2D
                // ==========================================================
                bool isReplaced = false;
                foreach (FileDescriptor fd in dwgDoc.File.ReferencedFileDescriptors)
                {
                    string ext = System.IO.Path.GetExtension(fd.FullFileName).ToLower();
                    if (ext == ".iam" || ext == ".ipt")
                    {
                        fd.ReplaceReference(newModelPath);
                        isReplaced = true;
                        break;
                    }
                }

                if (!isReplaced) throw new Exception("Không tìm thấy file 3D để thay thế.");
                dwgDoc.Update();

                // ==========================================================
                // BƯỚC 3: TÁI TẠO DỰA TRÊN TỌA ĐỘ QUÉT (DUNG SAI 1CM)
                // ==========================================================
                foreach (Sheet sheet in dwgDoc.Sheets)
                {
                    foreach (DimData d in dimBank)
                    {
                        try
                        {
                            DrawingCurve c1 = FindClosestCurveOnSheet(sheet, d.Anchor1);
                            DrawingCurve c2 = FindClosestCurveOnSheet(sheet, d.Anchor2);

                            if (c1 != null && c2 != null)
                            {
                                GeometryIntent int1 = sheet.CreateGeometryIntent(c1, d.Anchor1);
                                GeometryIntent int2 = sheet.CreateGeometryIntent(c2, d.Anchor2);
                                
                                LinearGeneralDimension newDim = sheet.DrawingDimensions.GeneralDimensions.AddLinear(d.TextPos, int1, int2, d.DimType);
                                newDim.Text.FormattedText = d.FormattedText; 
                            }
                        } catch { }
                    }

                    foreach (NoteData n in noteBank)
                    {
                        try
                        {
                            DrawingCurve curve = FindClosestCurveOnSheet(sheet, n.Anchor);
                            if (curve != null)
                            {
                                ObjectCollection pts = _inventorApp.TransientObjects.CreateObjectCollection();
                                pts.Add(n.NotePos);
                                pts.Add(sheet.CreateGeometryIntent(curve, n.Anchor));
                                
                                LeaderNote newNote = sheet.DrawingNotes.LeaderNotes.Add(pts, "");
                                newNote.FormattedText = n.FormattedText; 
                            }
                        } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Lỗi hệ thống: " + ex.Message);
            }
        }

        // --- CÔNG CỤ DÒ TÌM ĐƯỜNG NÉT TRÊN GIẤY ---
        private DrawingCurve FindClosestCurveOnSheet(Sheet sheet, Point2d pt)
        {
            if (pt == null) return null;
            double minDist = 1.0; // Dung sai 1cm (10mm)
            DrawingCurve bestCurve = null;

            foreach (DrawingView view in sheet.DrawingViews)
            {
                foreach (DrawingCurve c in view.DrawingCurves)
                {
                    try
                    {
                        Box2d box = c.Evaluator2D.RangeBox;
                        if (pt.X < box.MinPoint.X - minDist || pt.X > box.MaxPoint.X + minDist ||
                            pt.Y < box.MinPoint.Y - minDist || pt.Y > box.MaxPoint.Y + minDist)
                            continue;

                        c.Evaluator2D.GetParamExtents(out double min, out double max);
                        for (int i = 0; i <= 10; i++)
                        {
                            double param = min + (max - min) * (i / 10.0);
                            double[] pArr = { param };
                            double[] ptArr = new double[2];
                            c.Evaluator2D.GetPointAtParam(ref pArr, ref ptArr);

                            double dist = Math.Sqrt(Math.Pow(ptArr[0] - pt.X, 2) + Math.Pow(ptArr[1] - pt.Y, 2));
                            if (dist < minDist)
                            {
                                minDist = dist;
                                bestCurve = c;
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