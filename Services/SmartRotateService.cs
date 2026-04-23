using System;
using System.Collections.Generic;
using Inventor;

namespace InventorDrawingPlugin.Services
{
    /// <summary>
    /// Phase 1 + Phase 2: Rotate DrawingView + preserve section views
    /// Co che:
    /// - Inventor API chan rotate khi view co BAT KY sketch nao
    /// - Giai phap: save raw geometry cua sketches (sheet space) -> delete sketches
    ///   -> rotate view -> recreate sketches voi toa do da xoay
    /// - Children (section/detail/projected) duoc rotate theo cung goc de giu phuong cat
    /// Mat mat:
    /// - Sketch constraints (coincident, parallel, ...): MAT
    /// - Projected edges trong sketch: MAT (Phase 3 xu ly sau)
    /// - Dimensions bam vao sketch entities: co the MAT
    /// Giu duoc:
    /// - Vi tri section line sketch (cut direction tuong doi voi main view)
    /// - Vi tri va orientation cua section/detail views
    /// - Dimensions/notes bam vao model edges
    /// </summary>
    public class SmartRotateService
    {
        private Inventor.Application _inventorApp;
        private Action<string> _log;

        public int CountAutoRotated { get; private set; }
        public int CountManuallyFixed { get; private set; }
        public int CountFailed { get; private set; }
        public int CountChildrenRotated { get; private set; }
        public int CountChildrenFailed { get; private set; }
        public int CountSketchesRestored { get; private set; }
        public int CountSketchEntitiesLost { get; private set; }

        // ============================================================
        // DATA CLASSES
        // ============================================================
        private class AnnotationSnapshot
        {
            public object Target;
            public string Kind;
            public Point2d OldTextPos;
            public DrawingView OwnerView;
        }

        private class ChildViewSnapshot
        {
            public DrawingView View;
            public double OriginalRotation;
            public Point2d OriginalCenter;
        }

        private class SavedEntity
        {
            public string Kind;         // "Point", "Line", "Circle", "Arc"
            public Point2d P1;           // Point: pos | Line: start | Circle/Arc: center
            public Point2d P2;           // Line: end | Arc: start
            public Point2d P3;           // Arc: end
            public double Radius;        // Circle/Arc
            public double StartAngle;    // Arc
            public double SweepAngle;    // Arc
            public bool CCW;             // Arc
            public bool Construction;
        }

        private class SavedSketch
        {
            public DrawingView OwnerView;
            public List<SavedEntity> Entities = new List<SavedEntity>();
            public int LostRefEntities; // so reference entities bi bo qua (projected tu 3D)
        }

        public SmartRotateService(Inventor.Application app, Action<string> logger = null)
        {
            _inventorApp = app;
            _log = logger ?? (msg => { });
        }

        // ============================================================
        // MAIN
        // ============================================================
        public void ExecuteRotate(DrawingView view, double angleDegrees)
        {
            if (view == null) { _log("[ERROR] View null"); return; }
            if (Math.Abs(angleDegrees) < 0.001) { _log("[INFO] Goc qua nho, bo qua"); return; }

            CountAutoRotated = 0; CountManuallyFixed = 0; CountFailed = 0;
            CountChildrenRotated = 0; CountChildrenFailed = 0;
            CountSketchesRestored = 0; CountSketchEntitiesLost = 0;

            double angleRad = angleDegrees * Math.PI / 180.0;
            Sheet sheet = view.Parent as Sheet;
            if (sheet == null) { _log("[ERROR] Khong lay duoc Sheet"); return; }

            LogViewDiagnostic(view);

            if (view.ParentView != null && view.ViewType == DrawingViewTypeEnum.kProjectedDrawingViewType)
            {
                _log("[ERROR] Projected view khong the rotate doc lap. Hay rotate parent view.");
                return;
            }

            Point2d viewCenter = view.Center;
            _log($"Rotate view '{view.Name}' {angleDegrees:F1}° quanh ({viewCenter.X:F2}, {viewCenter.Y:F2})");

            // BUOC 1: Tim child views
            var childViews = FindChildViews(sheet, view);
            _log($"Tim thay {childViews.Count} child views");

            // BUOC 2: Snapshot annotations (main + children)
            var snapshots = new List<AnnotationSnapshot>();
            CollectAllAnnotations(sheet, view, snapshots);
            foreach (var child in childViews)
                CollectAllAnnotations(sheet, child.View, snapshots);
            _log($"Snapshot {snapshots.Count} annotations");

            // BUOC 3: Save + delete sketches tren main view va cac child views
            var savedMainSketches = SaveAndDeleteSketches(view);
            _log($"Main view: saved {savedMainSketches.Count} sketches ({CountSavedEntities(savedMainSketches)} entities)");

            var savedChildSketches = new Dictionary<DrawingView, List<SavedSketch>>();
            foreach (var child in childViews)
            {
                var cs = SaveAndDeleteSketches(child.View);
                if (cs.Count > 0)
                {
                    savedChildSketches[child.View] = cs;
                    _log($"Child '{child.View.Name}': saved {cs.Count} sketches ({CountSavedEntities(cs)} entities)");
                }
            }

            // BUOC 4: Rotate main view
            bool mainRotated = false;
            try
            {
                view.Rotation += angleRad;
                mainRotated = true;
                _log("Main view rotated OK");
            }
            catch (Exception ex)
            {
                _log($"[FATAL] Rotate main view that bai: {ex.Message}");
            }

            // BUOC 5: Rotate child views theo cung goc (giu cut direction phu hop voi main moi)
            if (mainRotated)
            {
                foreach (var child in childViews)
                    RotateChildView(child, angleRad);
            }

            // BUOC 6: Restore sketches voi toa do da rotate
            if (mainRotated)
            {
                RestoreSketches(view, savedMainSketches, viewCenter, angleRad);
                foreach (var kv in savedChildSketches)
                {
                    // Child view cung rotate cung goc, nen sketch tren child
                    // cung rotate quanh center cua chinh no
                    var childCenter = kv.Key.Center;
                    RestoreSketches(kv.Key, kv.Value, childCenter, angleRad);
                }
            }

            // BUOC 7: Fix annotations khong tu rotate
            if (mainRotated)
            {
                foreach (var snap in snapshots)
                {
                    try { FixAnnotationIfNeeded(snap, angleRad); }
                    catch (Exception ex)
                    {
                        _log($"[WARN] Fix annotation: {ex.Message}");
                        CountFailed++;
                    }
                }
            }

            // BAO CAO
            _log("=========================================");
            _log($"Main:     AUTO-ROT={CountAutoRotated}, FIXED={CountManuallyFixed}, FAIL={CountFailed}");
            _log($"Children: ROTATED={CountChildrenRotated}, FAIL={CountChildrenFailed}");
            _log($"Sketches: RESTORED={CountSketchesRestored}, LOST_REF_ENTITIES={CountSketchEntitiesLost}");
            _log("=========================================");
        }

        // ============================================================
        // DIAGNOSTIC
        // ============================================================
        private void LogViewDiagnostic(DrawingView view)
        {
            try
            {
                _log($"  Type: {view.ViewType}");
                _log($"  ParentView: {(view.ParentView != null ? view.ParentView.Name : "none")}");
                _log($"  Current rotation: {view.Rotation * 180.0 / Math.PI:F1}°");
            }
            catch { }
        }

        // ============================================================
        // CHILD VIEWS
        // ============================================================
        private List<ChildViewSnapshot> FindChildViews(Sheet sheet, DrawingView parent)
        {
            var list = new List<ChildViewSnapshot>();
            foreach (DrawingView v in sheet.DrawingViews)
            {
                try
                {
                    if (v.ParentView != null && ReferenceEquals(v.ParentView, parent))
                    {
                        list.Add(new ChildViewSnapshot
                        {
                            View = v,
                            OriginalRotation = v.Rotation,
                            OriginalCenter = v.Center
                        });
                    }
                }
                catch { }
            }
            return list;
        }

        private void RotateChildView(ChildViewSnapshot child, double angleRad)
        {
            try
            {
                child.View.Rotation += angleRad;
                CountChildrenRotated++;
                _log($"  Rotated child '{child.View.Name}'");
            }
            catch (Exception ex)
            {
                CountChildrenFailed++;
                _log($"  [WARN] Rotate child '{child.View.Name}' that bai: {ex.Message}");
            }
        }

        // ============================================================
        // SAVE / RESTORE SKETCHES
        // ============================================================
        private List<SavedSketch> SaveAndDeleteSketches(DrawingView view)
        {
            var saved = new List<SavedSketch>();
            var sketchesToDelete = new List<DrawingSketch>();

            try
            {
                foreach (DrawingSketch sk in view.Sketches)
                {
                    var ss = new SavedSketch { OwnerView = view };

                    // Save SketchPoints (user-drawn only)
                    try
                    {
                        foreach (SketchPoint pt in sk.SketchPoints)
                        {
                            if (pt.Reference) { ss.LostRefEntities++; continue; }
                            if (pt.HoleCenter) continue; // hole centers thuoc ve holes, bo qua
                            try
                            {
                                ss.Entities.Add(new SavedEntity
                                {
                                    Kind = "Point",
                                    P1 = sk.SketchToSheetSpace(pt.Geometry),
                                    Construction = pt.Construction
                                });
                            }
                            catch { }
                        }
                    }
                    catch { }

                    // Save SketchLines
                    try
                    {
                        foreach (SketchLine ln in sk.SketchLines)
                        {
                            if (ln.Reference) { ss.LostRefEntities++; continue; }
                            try
                            {
                                ss.Entities.Add(new SavedEntity
                                {
                                    Kind = "Line",
                                    P1 = sk.SketchToSheetSpace(ln.StartSketchPoint.Geometry),
                                    P2 = sk.SketchToSheetSpace(ln.EndSketchPoint.Geometry),
                                    Construction = ln.Construction
                                });
                            }
                            catch { }
                        }
                    }
                    catch { }

                    // Save SketchCircles
                    try
                    {
                        foreach (SketchCircle c in sk.SketchCircles)
                        {
                            if (c.Reference) { ss.LostRefEntities++; continue; }
                            try
                            {
                                ss.Entities.Add(new SavedEntity
                                {
                                    Kind = "Circle",
                                    P1 = sk.SketchToSheetSpace(c.CenterSketchPoint.Geometry),
                                    Radius = c.Radius,
                                    Construction = c.Construction
                                });
                            }
                            catch { }
                        }
                    }
                    catch { }

                    // Save SketchArcs
                    try
                    {
                        foreach (SketchArc a in sk.SketchArcs)
                        {
                            if (a.Reference) { ss.LostRefEntities++; continue; }
                            try
                            {
                                ss.Entities.Add(new SavedEntity
                                {
                                    Kind = "Arc",
                                    P1 = sk.SketchToSheetSpace(a.CenterSketchPoint.Geometry),
                                    P2 = sk.SketchToSheetSpace(a.StartSketchPoint.Geometry),
                                    P3 = sk.SketchToSheetSpace(a.EndSketchPoint.Geometry),
                                    Radius = a.Radius,
                                    StartAngle = a.StartAngle,
                                    SweepAngle = a.SweepAngle,
                                    CCW = a.SweepAngle > 0,
                                    Construction = a.Construction
                                });
                            }
                            catch { }
                        }
                    }
                    catch { }

                    CountSketchEntitiesLost += ss.LostRefEntities;
                    saved.Add(ss);
                    sketchesToDelete.Add(sk);
                }
            }
            catch (Exception ex)
            {
                _log($"[WARN] SaveAndDeleteSketches: {ex.Message}");
            }

            // Delete tat ca sketches sau khi da save xong
            foreach (var sk in sketchesToDelete)
            {
                try { sk.Delete(); }
                catch (Exception ex) { _log($"[WARN] Delete sketch that bai: {ex.Message}"); }
            }

            return saved;
        }

        private void RestoreSketches(DrawingView view, List<SavedSketch> saved, Point2d center, double angleRad)
        {
            foreach (var ss in saved)
            {
                if (ss.Entities.Count == 0) continue;

                DrawingSketch newSk = null;
                try
                {
                    newSk = view.Sketches.Add();
                    newSk.Edit();

                    foreach (var e in ss.Entities)
                    {
                        try { RecreateEntity(newSk, e, center, angleRad); }
                        catch { }
                    }

                    newSk.ExitEdit();
                    CountSketchesRestored++;
                }
                catch (Exception ex)
                {
                    _log($"[WARN] Restore sketch that bai: {ex.Message}");
                    if (newSk != null)
                    {
                        try { newSk.ExitEdit(); } catch { }
                    }
                }
            }
        }

        private void RecreateEntity(DrawingSketch sk, SavedEntity e, Point2d center, double angleRad)
        {
            // Xoay toa do sheet -> convert ve sketch space cua sketch moi
            Point2d p1Rot = RotatePoint(e.P1, center, angleRad);
            Point2d p1 = sk.SheetToSketchSpace(p1Rot);

            switch (e.Kind)
            {
                case "Point":
                    {
                        var sp = sk.SketchPoints.Add(p1, false);
                        sp.Construction = e.Construction;
                    }
                    break;

                case "Line":
                    {
                        Point2d p2 = sk.SheetToSketchSpace(RotatePoint(e.P2, center, angleRad));
                        var sl = sk.SketchLines.AddByTwoPoints(p1, p2);
                        sl.Construction = e.Construction;
                    }
                    break;

                case "Circle":
                    {
                        var sc = sk.SketchCircles.AddByCenterRadius(p1, e.Radius);
                        sc.Construction = e.Construction;
                    }
                    break;

                case "Arc":
                    {
                        Point2d p2 = sk.SheetToSketchSpace(RotatePoint(e.P2, center, angleRad));
                        Point2d p3 = sk.SheetToSketchSpace(RotatePoint(e.P3, center, angleRad));
                        var sa = sk.SketchArcs.AddByCenterStartEndPoint(p1, p2, p3, e.CCW);
                        sa.Construction = e.Construction;
                    }
                    break;
            }
        }

        private int CountSavedEntities(List<SavedSketch> saved)
        {
            int n = 0;
            foreach (var s in saved) n += s.Entities.Count;
            return n;
        }

        // ============================================================
        // ANNOTATIONS
        // ============================================================
        private void CollectAllAnnotations(Sheet sheet, DrawingView view, List<AnnotationSnapshot> snaps)
        {
            foreach (DrawingDimension dim in sheet.DrawingDimensions)
            {
                try
                {
                    if (!dim.Attached) continue;
                    if (!DimensionBelongsToView(dim, view)) continue;
                    snaps.Add(new AnnotationSnapshot
                    {
                        Target = dim, Kind = "Dim",
                        OldTextPos = dim.Text.Origin, OwnerView = view
                    });
                }
                catch { }
            }

            foreach (LeaderNote note in sheet.DrawingNotes.LeaderNotes)
            {
                try
                {
                    if (!LeaderNoteBelongsToView(note, view)) continue;
                    snaps.Add(new AnnotationSnapshot
                    {
                        Target = note, Kind = "LeaderNote",
                        OldTextPos = note.Position, OwnerView = view
                    });
                }
                catch { }
            }

            foreach (HoleThreadNote hNote in sheet.DrawingNotes.HoleThreadNotes)
            {
                try
                {
                    if (hNote.Intent?.Geometry is DrawingCurve dc && ReferenceEquals(dc.Parent, view))
                    {
                        snaps.Add(new AnnotationSnapshot
                        {
                            Target = hNote, Kind = "HoleNote",
                            OldTextPos = hNote.Text.Origin, OwnerView = view
                        });
                    }
                }
                catch { }
            }
        }

        private void FixAnnotationIfNeeded(AnnotationSnapshot snap, double angleRad)
        {
            Point2d viewCenter = snap.OwnerView.Center;
            Point2d expectedPos = RotatePoint(snap.OldTextPos, viewCenter, angleRad);
            Point2d currentPos = GetCurrentTextPos(snap);
            if (currentPos == null) { CountFailed++; return; }

            double distToExpected = DistBetween(currentPos, expectedPos);
            double distToOld = DistBetween(currentPos, snap.OldTextPos);

            if (distToExpected < 0.1) { CountAutoRotated++; return; }
            if (distToOld < 0.1)
            {
                if (SetTextPos(snap, expectedPos)) CountManuallyFixed++;
                else CountFailed++;
                return;
            }
            CountAutoRotated++;
        }

        private Point2d GetCurrentTextPos(AnnotationSnapshot snap)
        {
            try
            {
                switch (snap.Kind)
                {
                    case "Dim":        return ((DrawingDimension)snap.Target).Text.Origin;
                    case "LeaderNote": return ((LeaderNote)snap.Target).Position;
                    case "HoleNote":   return ((HoleThreadNote)snap.Target).Text.Origin;
                }
            }
            catch { }
            return null;
        }

        private bool SetTextPos(AnnotationSnapshot snap, Point2d newPos)
        {
            try
            {
                switch (snap.Kind)
                {
                    case "Dim":        ((DrawingDimension)snap.Target).Text.Origin = newPos; return true;
                    case "LeaderNote": ((LeaderNote)snap.Target).Position = newPos; return true;
                    case "HoleNote":   ((HoleThreadNote)snap.Target).Text.Origin = newPos; return true;
                }
            }
            catch { }
            return false;
        }

        // ============================================================
        // HELPERS
        // ============================================================
        private Point2d RotatePoint(Point2d pt, Point2d center, double angleRad)
        {
            double dx = pt.X - center.X;
            double dy = pt.Y - center.Y;
            double cos = Math.Cos(angleRad);
            double sin = Math.Sin(angleRad);
            return _inventorApp.TransientGeometry.CreatePoint2d(
                center.X + dx * cos - dy * sin,
                center.Y + dx * sin + dy * cos);
        }

        private double DistBetween(Point2d a, Point2d b)
        {
            return Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
        }

        private bool DimensionBelongsToView(DrawingDimension dim, DrawingView view)
        {
            try
            {
                if (dim is LinearGeneralDimension lin)
                {
                    if (lin.IntentOne?.Geometry is DrawingCurve c1 && ReferenceEquals(c1.Parent, view)) return true;
                    if (lin.IntentTwo?.Geometry is DrawingCurve c2 && ReferenceEquals(c2.Parent, view)) return true;
                }
                else if (dim is RadiusGeneralDimension rad)
                {
                    if (rad.Intent?.Geometry is DrawingCurve c && ReferenceEquals(c.Parent, view)) return true;
                }
                else if (dim is DiameterGeneralDimension dia)
                {
                    if (dia.Intent?.Geometry is DrawingCurve c && ReferenceEquals(c.Parent, view)) return true;
                }
                else if (dim is AngularGeneralDimension ang)
                {
                    if (ang.IntentOne?.Geometry is DrawingCurve c1 && ReferenceEquals(c1.Parent, view)) return true;
                    if (ang.IntentTwo?.Geometry is DrawingCurve c2 && ReferenceEquals(c2.Parent, view)) return true;
                }
            }
            catch { }
            return false;
        }

        private bool LeaderNoteBelongsToView(LeaderNote note, DrawingView view)
        {
            try
            {
                object attached = note.Leader.RootNode.AttachedEntity;
                if (attached is GeometryIntent gi && gi.Geometry is DrawingCurve dc)
                    return ReferenceEquals(dc.Parent, view);
                if (attached is DrawingCurve dc2)
                    return ReferenceEquals(dc2.Parent, view);
            }
            catch { }
            return false;
        }
    }
}
