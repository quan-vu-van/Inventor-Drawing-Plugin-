using System;
using System.Collections.Generic;
using Inventor;

namespace InventorDrawingPlugin.Services
{
    /// <summary>
    /// Luu va khoi phuc constraint giua sketch entities va projected edges (tu 3D model).
    /// Sau khi replace model, projected edges mat proxy -> constraint bi dut.
    /// Service nay luu vi tri cua cac reference entity truoc replace,
    /// roi tim projected edge moi va tao lai constraint sau replace.
    /// </summary>
    public class SketchConstraintService
    {
        private Inventor.Application _inventorApp;
        private Action<string> _log;

        // --- Ket qua bao cao ---
        public int CountConstraintSaved { get; private set; }
        public int CountConstraintRestored { get; private set; }
        public int CountConstraintLost { get; private set; }

        // ============================================================
        // DNA: Luu thong tin 1 reference entity trong sketch
        // va cac constraint ket noi no voi user entities
        // ============================================================
        private class RefEntityDNA
        {
            // Dinh danh sketch
            public string SheetName;
            public string ViewName;       // ten view chua sketch (null neu la section line sketch)
            public string SketchIndex;    // vi tri sketch trong collection (de tim lai)
            public bool IsSectionLineSketch;

            // Thong tin reference entity (projected tu 3D)
            public string RefType;        // "Point", "Line", "Circle", "Arc"
            public Point2d RefPos;        // vi tri (sheet space) - dung de tim DrawingCurve moi
            public Point2d RefStartPt;    // cho Line: start point (sheet space)
            public Point2d RefEndPt;      // cho Line: end point (sheet space)

            // Cac user entities duoc constraint voi reference entity nay
            public List<ConnectionDNA> Connections = new List<ConnectionDNA>();
        }

        private class ConnectionDNA
        {
            public string ConstraintType; // "Coincident", "TwoPointDistance", "Collinear", etc.

            // Vi tri user entity (sketch space) - de tim lai sau replace
            public Point2d UserEntityPos;
            public string UserEntityType; // "Point", "Line", "Circle"

            // Cho DimensionConstraint
            public double? DimValue;
            public Point2d TextPoint;
            public string Orientation;    // "Horizontal", "Vertical", "Aligned"
            public bool? Driven;
        }

        private List<RefEntityDNA> _dnaBank;

        public SketchConstraintService(Inventor.Application app, Action<string> logger = null)
        {
            _inventorApp = app;
            _log = logger ?? (msg => { });
        }

        // ============================================================
        // BUOC A: TRICH XUAT DNA VA XOA CONSTRAINT BI HU
        // Goi TRUOC khi replace model
        // ============================================================
        public void ExtractAndClean(DrawingDocument dwgDoc)
        {
            CountConstraintSaved = 0;
            CountConstraintRestored = 0;
            CountConstraintLost = 0;

            var dnaBank = new List<RefEntityDNA>();

            foreach (Sheet sheet in dwgDoc.Sheets)
            {
                foreach (DrawingView view in sheet.DrawingViews)
                {
                    // 1. Section line sketch (neu view la SectionDrawingView)
                    if (view.ViewType == DrawingViewTypeEnum.kSectionDrawingViewType)
                    {
                        try
                        {
                            SectionDrawingView secView = (SectionDrawingView)view;
                            DrawingSketch secSketch = secView.SectionLineSketch;
                            if (secSketch != null)
                            {
                                int count = ExtractFromSketch(secSketch, sheet.Name, view.Name, true, dnaBank);
                                if (count > 0)
                                    _log($"  Section '{view.Name}': {count} ref entities saved");
                            }
                        }
                        catch (Exception ex)
                        {
                            _log($"[WARN] Khong doc duoc section sketch '{view.Name}': {ex.Message}");
                        }
                    }

                    // 2. Independent sketches tren view
                    try
                    {
                        foreach (DrawingSketch sketch in view.Sketches)
                        {
                            int count = ExtractFromSketch(sketch, sheet.Name, view.Name, false, dnaBank);
                            if (count > 0)
                                _log($"  Sketch on '{view.Name}': {count} ref entities saved");
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[WARN] Khong doc duoc sketches tren view '{view.Name}': {ex.Message}");
                    }
                }
            }

            _dnaBank = dnaBank;
            CountConstraintSaved = dnaBank.Count;
            _log($"Tong: {dnaBank.Count} ref entities saved tu tat ca sketches");
        }

        // ============================================================
        // BUOC B: KHOI PHUC CONSTRAINT SAU REPLACE
        // Goi SAU khi replace model va dwgDoc.Update()
        // ============================================================
        public void Restore(DrawingDocument dwgDoc)
        {
            if (_dnaBank == null || _dnaBank.Count == 0) return;

            foreach (Sheet sheet in dwgDoc.Sheets)
            {
                foreach (DrawingView view in sheet.DrawingViews)
                {
                    // Tim DNA thuoc view nay
                    foreach (var dna in _dnaBank)
                    {
                        if (dna.SheetName != sheet.Name || dna.ViewName != view.Name)
                            continue;

                        DrawingSketch targetSketch = null;

                        // Tim lai sketch
                        if (dna.IsSectionLineSketch && view.ViewType == DrawingViewTypeEnum.kSectionDrawingViewType)
                        {
                            try
                            {
                                targetSketch = ((SectionDrawingView)view).SectionLineSketch;
                            }
                            catch { }
                        }
                        else
                        {
                            // Tim sketch trong view.Sketches theo index
                            try
                            {
                                int idx;
                                if (int.TryParse(dna.SketchIndex, out idx) && idx <= view.Sketches.Count)
                                    targetSketch = view.Sketches[idx];
                            }
                            catch { }
                        }

                        if (targetSketch == null)
                        {
                            _log($"[WARN] Khong tim lai duoc sketch tren view '{dna.ViewName}'");
                            CountConstraintLost += dna.Connections.Count;
                            continue;
                        }

                        RestoreForSketch(targetSketch, dna, sheet, view);
                    }
                }
            }

            _log($"Sketch constraints: RESTORED={CountConstraintRestored}, LOST={CountConstraintLost}");
        }

        // ============================================================
        // TRICH XUAT TU 1 SKETCH CU THE
        // ============================================================
        private int ExtractFromSketch(DrawingSketch sketch, string sheetName, string viewName,
                                       bool isSectionLine, List<RefEntityDNA> dnaBank)
        {
            int savedCount = 0;

            // Tim sketch index trong parent view (de tim lai sau replace)
            string sketchIndex = "1";
            if (!isSectionLine)
            {
                try
                {
                    var parentView = sketch.Parent as DrawingView;
                    if (parentView != null)
                    {
                        int idx = 1;
                        foreach (DrawingSketch s in parentView.Sketches)
                        {
                            if (ReferenceEquals(s, sketch)) { sketchIndex = idx.ToString(); break; }
                            idx++;
                        }
                    }
                }
                catch { }
            }

            // Thu thap tat ca reference entities va constraint cua chung
            var refEntities = new List<object>();
            var refDnaMap = new Dictionary<object, RefEntityDNA>();

            // Quet SketchPoints
            foreach (SketchPoint pt in sketch.SketchPoints)
            {
                try
                {
                    if (!pt.Reference) continue;
                    Point2d sheetPos = sketch.SketchToSheetSpace(pt.Geometry);

                    var dna = new RefEntityDNA
                    {
                        SheetName = sheetName,
                        ViewName = viewName,
                        SketchIndex = sketchIndex,
                        IsSectionLineSketch = isSectionLine,
                        RefType = "Point",
                        RefPos = sheetPos
                    };

                    refEntities.Add(pt);
                    refDnaMap[pt] = dna;
                }
                catch { }
            }

            // Quet SketchLines
            foreach (SketchLine line in sketch.SketchLines)
            {
                try
                {
                    if (!line.Reference) continue;
                    Point2d startSheet = sketch.SketchToSheetSpace(line.StartSketchPoint.Geometry);
                    Point2d endSheet = sketch.SketchToSheetSpace(line.EndSketchPoint.Geometry);
                    Point2d midSheet = _inventorApp.TransientGeometry.CreatePoint2d(
                        (startSheet.X + endSheet.X) / 2, (startSheet.Y + endSheet.Y) / 2);

                    var dna = new RefEntityDNA
                    {
                        SheetName = sheetName,
                        ViewName = viewName,
                        SketchIndex = sketchIndex,
                        IsSectionLineSketch = isSectionLine,
                        RefType = "Line",
                        RefPos = midSheet,
                        RefStartPt = startSheet,
                        RefEndPt = endSheet
                    };

                    refEntities.Add(line);
                    refDnaMap[line] = dna;
                }
                catch { }
            }

            // Quet SketchCircles
            foreach (SketchCircle circ in sketch.SketchCircles)
            {
                try
                {
                    if (!circ.Reference) continue;
                    Point2d centerSheet = sketch.SketchToSheetSpace(circ.CenterSketchPoint.Geometry);

                    var dna = new RefEntityDNA
                    {
                        SheetName = sheetName,
                        ViewName = viewName,
                        SketchIndex = sketchIndex,
                        IsSectionLineSketch = isSectionLine,
                        RefType = "Circle",
                        RefPos = centerSheet
                    };

                    refEntities.Add(circ);
                    refDnaMap[circ] = dna;
                }
                catch { }
            }

            if (refEntities.Count == 0) return 0;

            // Thu thap connections tu DimensionConstraints
            foreach (DimensionConstraint dc in sketch.DimensionConstraints)
            {
                try
                {
                    CollectDimConstraintConnections(dc, sketch, refDnaMap);
                }
                catch { }
            }

            // Thu thap connections tu GeometricConstraints
            foreach (GeometricConstraint gc in sketch.GeometricConstraints)
            {
                try
                {
                    CollectGeoConstraintConnections(gc, sketch, refDnaMap);
                }
                catch { }
            }

            // Chi luu DNA co connections (co constraint thuc su)
            foreach (var dna in refDnaMap.Values)
            {
                if (dna.Connections.Count > 0)
                {
                    dnaBank.Add(dna);
                    savedCount++;
                }
            }

            // Xoa constraints va reference entities (chung se hu sau replace)
            // Lam trong try/catch vi co the da bi hu
            foreach (var refEntity in refEntities)
            {
                if (!refDnaMap.ContainsKey(refEntity)) continue;
                if (refDnaMap[refEntity].Connections.Count == 0) continue;

                try
                {
                    // Xoa constraints truoc
                    var constraintsToDelete = new List<object>();
                    // Cast to SketchEntity de truy cap .Constraints
                    var se = (SketchEntity)refEntity;
                    foreach (object c in se.Constraints)
                        constraintsToDelete.Add(c);

                    foreach (object c in constraintsToDelete)
                    {
                        try
                        {
                            if (c is GeometricConstraint gc && gc.Deletable) gc.Delete();
                            else if (c is DimensionConstraint dc) dc.Delete();
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return savedCount;
        }

        // ============================================================
        // THU THAP CONNECTION TU DIMENSION CONSTRAINT
        // ============================================================
        private void CollectDimConstraintConnections(DimensionConstraint dc, DrawingSketch sketch,
                                                      Dictionary<object, RefEntityDNA> refDnaMap)
        {
            if (dc is TwoPointDistanceDimConstraint tpd)
            {
                // Kiem tra xem point nao la reference, point nao la user
                SketchPoint refPt = null, userPt = null;
                RefEntityDNA dna = null;

                if (tpd.PointOne.Reference && refDnaMap.ContainsKey(tpd.PointOne))
                    { refPt = tpd.PointOne; userPt = tpd.PointTwo; dna = refDnaMap[tpd.PointOne]; }
                else if (tpd.PointTwo.Reference && refDnaMap.ContainsKey(tpd.PointTwo))
                    { refPt = tpd.PointTwo; userPt = tpd.PointOne; dna = refDnaMap[tpd.PointTwo]; }

                if (dna != null && userPt != null)
                {
                    string orient = "Aligned";
                    try { orient = tpd.Orientation.ToString().Replace("kDimensionOrientation", ""); } catch { }

                    dna.Connections.Add(new ConnectionDNA
                    {
                        ConstraintType = "TwoPointDistance",
                        UserEntityPos = sketch.SketchToSheetSpace(userPt.Geometry),
                        UserEntityType = "Point",
                        DimValue = (double)tpd.Parameter.Value,
                        TextPoint = tpd.TextPoint,
                        Orientation = orient,
                        Driven = tpd.Driven
                    });
                }
            }
            else if (dc is TwoLineAngleDimConstraint tla)
            {
                SketchLine refLine = null, userLine = null;
                RefEntityDNA dna = null;

                if (tla.LineOne.Reference && refDnaMap.ContainsKey(tla.LineOne))
                    { refLine = tla.LineOne; userLine = tla.LineTwo; dna = refDnaMap[tla.LineOne]; }
                else if (tla.LineTwo.Reference && refDnaMap.ContainsKey(tla.LineTwo))
                    { refLine = tla.LineTwo; userLine = tla.LineOne; dna = refDnaMap[tla.LineTwo]; }

                if (dna != null && userLine != null)
                {
                    Point2d mid = _inventorApp.TransientGeometry.CreatePoint2d(
                        (userLine.StartSketchPoint.Geometry.X + userLine.EndSketchPoint.Geometry.X) / 2,
                        (userLine.StartSketchPoint.Geometry.Y + userLine.EndSketchPoint.Geometry.Y) / 2);

                    dna.Connections.Add(new ConnectionDNA
                    {
                        ConstraintType = "TwoLineAngle",
                        UserEntityPos = sketch.SketchToSheetSpace(mid),
                        UserEntityType = "Line",
                        DimValue = (double)tla.Parameter.Value,
                        TextPoint = tla.TextPoint,
                        Driven = tla.Driven
                    });
                }
            }
        }

        // ============================================================
        // THU THAP CONNECTION TU GEOMETRIC CONSTRAINT
        // ============================================================
        private void CollectGeoConstraintConnections(GeometricConstraint gc, DrawingSketch sketch,
                                                      Dictionary<object, RefEntityDNA> refDnaMap)
        {
            if (gc is CoincidentConstraint cc)
            {
                SketchEntity refEntity = null, userEntity = null;
                RefEntityDNA dna = null;

                if (cc.EntityOne.Reference && refDnaMap.ContainsKey(cc.EntityOne))
                    { refEntity = cc.EntityOne; userEntity = cc.EntityTwo; dna = refDnaMap[cc.EntityOne]; }
                else if (cc.EntityTwo.Reference && refDnaMap.ContainsKey(cc.EntityTwo))
                    { refEntity = cc.EntityTwo; userEntity = cc.EntityOne; dna = refDnaMap[cc.EntityTwo]; }

                if (dna != null && userEntity != null)
                {
                    Point2d userPos = GetEntityPosition(userEntity, sketch);
                    if (userPos != null)
                    {
                        dna.Connections.Add(new ConnectionDNA
                        {
                            ConstraintType = "Coincident",
                            UserEntityPos = sketch.SketchToSheetSpace(userPos),
                            UserEntityType = GetEntityTypeName(userEntity)
                        });
                    }
                }
            }
            else if (gc is CollinearConstraint col)
            {
                SketchEntity refEntity = null, userEntity = null;
                RefEntityDNA dna = null;

                if (col.EntityOne.Reference && refDnaMap.ContainsKey(col.EntityOne))
                    { refEntity = col.EntityOne; userEntity = col.EntityTwo; dna = refDnaMap[col.EntityOne]; }
                else if (col.EntityTwo.Reference && refDnaMap.ContainsKey(col.EntityTwo))
                    { refEntity = col.EntityTwo; userEntity = col.EntityOne; dna = refDnaMap[col.EntityTwo]; }

                if (dna != null && userEntity != null)
                {
                    Point2d userPos = GetEntityPosition(userEntity, sketch);
                    if (userPos != null)
                    {
                        dna.Connections.Add(new ConnectionDNA
                        {
                            ConstraintType = "Collinear",
                            UserEntityPos = sketch.SketchToSheetSpace(userPos),
                            UserEntityType = GetEntityTypeName(userEntity)
                        });
                    }
                }
            }
            else if (gc is PerpendicularConstraint perp)
            {
                SketchEntity refEntity = null, userEntity = null;
                RefEntityDNA dna = null;

                if (perp.EntityOne.Reference && refDnaMap.ContainsKey(perp.EntityOne))
                    { refEntity = perp.EntityOne; userEntity = perp.EntityTwo; dna = refDnaMap[perp.EntityOne]; }
                else if (perp.EntityTwo.Reference && refDnaMap.ContainsKey(perp.EntityTwo))
                    { refEntity = perp.EntityTwo; userEntity = perp.EntityOne; dna = refDnaMap[perp.EntityTwo]; }

                if (dna != null && userEntity != null)
                {
                    Point2d userPos = GetEntityPosition(userEntity, sketch);
                    if (userPos != null)
                    {
                        dna.Connections.Add(new ConnectionDNA
                        {
                            ConstraintType = "Perpendicular",
                            UserEntityPos = sketch.SketchToSheetSpace(userPos),
                            UserEntityType = GetEntityTypeName(userEntity)
                        });
                    }
                }
            }
            else if (gc is ParallelConstraint par)
            {
                SketchEntity refEntity = null, userEntity = null;
                RefEntityDNA dna = null;

                if (par.EntityOne.Reference && refDnaMap.ContainsKey(par.EntityOne))
                    { refEntity = par.EntityOne; userEntity = par.EntityTwo; dna = refDnaMap[par.EntityOne]; }
                else if (par.EntityTwo.Reference && refDnaMap.ContainsKey(par.EntityTwo))
                    { refEntity = par.EntityTwo; userEntity = par.EntityOne; dna = refDnaMap[par.EntityTwo]; }

                if (dna != null && userEntity != null)
                {
                    Point2d userPos = GetEntityPosition(userEntity, sketch);
                    if (userPos != null)
                    {
                        dna.Connections.Add(new ConnectionDNA
                        {
                            ConstraintType = "Parallel",
                            UserEntityPos = sketch.SketchToSheetSpace(userPos),
                            UserEntityType = GetEntityTypeName(userEntity)
                        });
                    }
                }
            }
        }

        // ============================================================
        // KHOI PHUC CHO 1 SKETCH CU THE
        // ============================================================
        private void RestoreForSketch(DrawingSketch sketch, RefEntityDNA dna, Sheet sheet, DrawingView view)
        {
            foreach (var conn in dna.Connections)
            {
                try
                {
                    // Tim DrawingCurve moi gan nhat tai vi tri ref entity cu
                    string curveFilter = null;
                    if (dna.RefType == "Circle" || dna.RefType == "Arc") curveFilter = "Arc";
                    else if (dna.RefType == "Line") curveFilter = "Line";

                    DrawingCurve newCurve = FindClosestCurveOnView(view, dna.RefPos, curveFilter);
                    if (newCurve == null)
                    {
                        _log($"  [WARN] Khong tim thay curve moi cho ref entity tai ({dna.RefPos.X:F2}, {dna.RefPos.Y:F2})");
                        CountConstraintLost++;
                        continue;
                    }

                    // Project curve moi vao sketch
                    sketch.Edit();
                    SketchEntity newRefEntity = null;
                    try
                    {
                        newRefEntity = sketch.AddByProjectingEntity(newCurve);
                    }
                    catch (Exception ex)
                    {
                        _log($"  [WARN] Khong project duoc curve: {ex.Message}");
                        try { sketch.ExitEdit(); } catch { }
                        CountConstraintLost++;
                        continue;
                    }

                    // Tim user entity trong sketch (theo vi tri)
                    SketchEntity userEntity = FindSketchEntityByPosition(sketch, conn.UserEntityPos, conn.UserEntityType);
                    if (userEntity == null)
                    {
                        _log($"  [WARN] Khong tim thay user entity tai ({conn.UserEntityPos.X:F2}, {conn.UserEntityPos.Y:F2})");
                        try { sketch.ExitEdit(); } catch { }
                        CountConstraintLost++;
                        continue;
                    }

                    // Tao lai constraint
                    bool restored = false;
                    try
                    {
                        restored = RecreateConstraint(sketch, conn, userEntity, newRefEntity);
                    }
                    catch (Exception ex)
                    {
                        _log($"  [WARN] Khong tao lai duoc constraint '{conn.ConstraintType}': {ex.Message}");
                    }

                    sketch.ExitEdit();

                    if (restored) CountConstraintRestored++;
                    else CountConstraintLost++;
                }
                catch (Exception ex)
                {
                    _log($"  [ERROR] RestoreForSketch: {ex.Message}");
                    try { sketch.ExitEdit(); } catch { }
                    CountConstraintLost++;
                }
            }
        }

        // ============================================================
        // TAO LAI CONSTRAINT
        // ============================================================
        private bool RecreateConstraint(DrawingSketch sketch, ConnectionDNA conn,
                                         SketchEntity userEntity, SketchEntity newRefEntity)
        {
            // COM interop voi EmbedInteropTypes can cast tuong minh
            SketchEntity user = (SketchEntity)userEntity;
            SketchEntity newRef = (SketchEntity)newRefEntity;

            switch (conn.ConstraintType)
            {
                case "Coincident":
                    sketch.GeometricConstraints.AddCoincident(user, newRef);
                    return true;

                case "Collinear":
                    sketch.GeometricConstraints.AddCollinear(user, newRef);
                    return true;

                case "Perpendicular":
                    sketch.GeometricConstraints.AddPerpendicular(user, newRef);
                    return true;

                case "Parallel":
                    sketch.GeometricConstraints.AddParallel(user, newRef);
                    return true;

                case "TwoPointDistance":
                    if (userEntity is SketchPoint userPt && newRefEntity is SketchPoint refPt)
                    {
                        DimensionOrientationEnum orient = DimensionOrientationEnum.kAlignedDim;
                        if (conn.Orientation == "Horizontal") orient = DimensionOrientationEnum.kHorizontalDim;
                        else if (conn.Orientation == "Vertical") orient = DimensionOrientationEnum.kVerticalDim;

                        var dc = sketch.DimensionConstraints.AddTwoPointDistance(
                            userPt, refPt, orient, conn.TextPoint ?? userPt.Geometry, conn.Driven ?? false);

                        if (conn.DimValue.HasValue)
                            dc.Parameter.Value = conn.DimValue.Value;

                        return true;
                    }
                    // Neu newRefEntity la Line/Circle, thu lay point gan nhat
                    if (userEntity is SketchPoint userPt2)
                    {
                        // Tim SketchPoint cua ref entity (start hoac end)
                        SketchPoint refPoint = GetNearestSketchPoint(newRefEntity, conn.UserEntityPos, sketch);
                        if (refPoint != null)
                        {
                            DimensionOrientationEnum orient = DimensionOrientationEnum.kAlignedDim;
                            if (conn.Orientation == "Horizontal") orient = DimensionOrientationEnum.kHorizontalDim;
                            else if (conn.Orientation == "Vertical") orient = DimensionOrientationEnum.kVerticalDim;

                            var dc = sketch.DimensionConstraints.AddTwoPointDistance(
                                userPt2, refPoint, orient, conn.TextPoint ?? userPt2.Geometry, conn.Driven ?? false);

                            if (conn.DimValue.HasValue)
                                dc.Parameter.Value = conn.DimValue.Value;

                            return true;
                        }
                    }
                    return false;

                case "TwoLineAngle":
                    if (userEntity is SketchLine userLine && newRefEntity is SketchLine refLine)
                    {
                        var dc = sketch.DimensionConstraints.AddTwoLineAngle(
                            userLine, refLine, conn.TextPoint ?? userLine.StartSketchPoint.Geometry, conn.Driven ?? false);

                        if (conn.DimValue.HasValue)
                            dc.Parameter.Value = conn.DimValue.Value;

                        return true;
                    }
                    return false;

                default:
                    _log($"  [WARN] Loai constraint '{conn.ConstraintType}' chua duoc ho tro");
                    return false;
            }
        }

        // ============================================================
        // HELPERS
        // ============================================================

        private Point2d GetEntityPosition(SketchEntity entity, DrawingSketch sketch)
        {
            if (entity is SketchPoint pt) return pt.Geometry;
            if (entity is SketchLine line)
                return _inventorApp.TransientGeometry.CreatePoint2d(
                    (line.StartSketchPoint.Geometry.X + line.EndSketchPoint.Geometry.X) / 2,
                    (line.StartSketchPoint.Geometry.Y + line.EndSketchPoint.Geometry.Y) / 2);
            if (entity is SketchCircle circ) return circ.CenterSketchPoint.Geometry;
            if (entity is SketchArc arc) return arc.CenterSketchPoint.Geometry;
            return null;
        }

        private string GetEntityTypeName(SketchEntity entity)
        {
            if (entity is SketchPoint) return "Point";
            if (entity is SketchLine) return "Line";
            if (entity is SketchCircle) return "Circle";
            if (entity is SketchArc) return "Arc";
            return "Unknown";
        }

        /// <summary>
        /// Tim SketchEntity trong sketch gan nhat voi vi tri cho truoc (sheet space)
        /// Chi tim user entities (Reference == false)
        /// </summary>
        private SketchEntity FindSketchEntityByPosition(DrawingSketch sketch, Point2d sheetPos, string entityType)
        {
            if (sheetPos == null) return null;

            Point2d sketchPos = sketch.SheetToSketchSpace(sheetPos);
            double bestDist = 0.5; // tolerance 5mm
            SketchEntity bestEntity = null;

            if (entityType == "Point")
            {
                foreach (SketchPoint pt in sketch.SketchPoints)
                {
                    if (pt.Reference) continue;
                    double dist = DistBetween(pt.Geometry, sketchPos);
                    if (dist < bestDist) { bestDist = dist; bestEntity = (SketchEntity)pt; }
                }
            }
            else if (entityType == "Line")
            {
                foreach (SketchLine line in sketch.SketchLines)
                {
                    if (line.Reference) continue;
                    Point2d mid = _inventorApp.TransientGeometry.CreatePoint2d(
                        (line.StartSketchPoint.Geometry.X + line.EndSketchPoint.Geometry.X) / 2,
                        (line.StartSketchPoint.Geometry.Y + line.EndSketchPoint.Geometry.Y) / 2);
                    double dist = DistBetween(mid, sketchPos);
                    if (dist < bestDist) { bestDist = dist; bestEntity = (SketchEntity)line; }
                }
            }
            else if (entityType == "Circle" || entityType == "Arc")
            {
                foreach (SketchCircle circ in sketch.SketchCircles)
                {
                    if (circ.Reference) continue;
                    double dist = DistBetween(circ.CenterSketchPoint.Geometry, sketchPos);
                    if (dist < bestDist) { bestDist = dist; bestEntity = (SketchEntity)circ; }
                }
            }

            // Fallback: tim bat ky entity nao gan nhat
            if (bestEntity == null)
            {
                foreach (SketchPoint pt in sketch.SketchPoints)
                {
                    if (pt.Reference) continue;
                    double dist = DistBetween(pt.Geometry, sketchPos);
                    if (dist < bestDist) { bestDist = dist; bestEntity = (SketchEntity)pt; }
                }
            }

            return bestEntity;
        }

        private SketchPoint GetNearestSketchPoint(SketchEntity entity, Point2d targetSheetPos, DrawingSketch sketch)
        {
            Point2d targetSketchPos = sketch.SheetToSketchSpace(targetSheetPos);

            if (entity is SketchPoint pt) return pt;
            if (entity is SketchLine line)
            {
                double d1 = DistBetween(line.StartSketchPoint.Geometry, targetSketchPos);
                double d2 = DistBetween(line.EndSketchPoint.Geometry, targetSketchPos);
                return d1 < d2 ? line.StartSketchPoint : line.EndSketchPoint;
            }
            if (entity is SketchCircle circ) return circ.CenterSketchPoint;
            if (entity is SketchArc arc) return arc.CenterSketchPoint;
            return null;
        }

        private double DistBetween(Point2d a, Point2d b)
        {
            return Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
        }

        /// <summary>
        /// Tim DrawingCurve gan nhat TRONG 1 VIEW (khong quet toan bo sheet)
        /// </summary>
        private DrawingCurve FindClosestCurveOnView(DrawingView view, Point2d sheetPt, string curveFilter = null)
        {
            if (sheetPt == null) return null;

            const double TOLERANCE = 0.5;
            const int SAMPLE_COUNT = 25;
            const double EXACT_THRESHOLD = 0.05;

            double minDist = TOLERANCE;
            DrawingCurve bestCurve = null;

            foreach (DrawingCurve c in view.DrawingCurves)
            {
                try
                {
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

                    Box2d box = c.Evaluator2D.RangeBox;
                    if (sheetPt.X < box.MinPoint.X - TOLERANCE || sheetPt.X > box.MaxPoint.X + TOLERANCE ||
                        sheetPt.Y < box.MinPoint.Y - TOLERANCE || sheetPt.Y > box.MaxPoint.Y + TOLERANCE)
                        continue;

                    c.Evaluator2D.GetParamExtents(out double pMin, out double pMax);

                    for (int i = 0; i <= SAMPLE_COUNT; i++)
                    {
                        double param = pMin + (pMax - pMin) * ((double)i / SAMPLE_COUNT);
                        double[] pArr = { param };
                        double[] ptArr = new double[2];
                        c.Evaluator2D.GetPointAtParam(ref pArr, ref ptArr);

                        double dist = Math.Sqrt(
                            (ptArr[0] - sheetPt.X) * (ptArr[0] - sheetPt.X) +
                            (ptArr[1] - sheetPt.Y) * (ptArr[1] - sheetPt.Y));

                        if (dist < minDist)
                        {
                            minDist = dist;
                            bestCurve = c;
                            if (dist < EXACT_THRESHOLD) return bestCurve;
                        }
                    }
                }
                catch { }
            }

            return bestCurve;
        }
    }
}
