using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

namespace AutoCADPlugin
{
    public class Commands
    {
        private static double surfaceHeight = 400.0; // Default height for the surface

        [CommandMethod("DRS")]
        public void DrawSurface()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // Prompt for the surface height
                    PromptDoubleOptions pdo = new PromptDoubleOptions("\nEnter surface height (current height: " + surfaceHeight + "): ");
                    pdo.AllowNone = true;
                    PromptDoubleResult pdr = ed.GetDouble(pdo);

                    if (pdr.Status == PromptStatus.OK)
                    {
                        surfaceHeight = pdr.Value;
                    }

                    // Select a line in the XY plane
                    PromptEntityOptions peo = new PromptEntityOptions("\nSelect a line in the XY plane: ");
                    peo.SetRejectMessage("\nSelected object must be a line.");
                    peo.AddAllowedClass(typeof(Line), true);
                    PromptEntityResult per = ed.GetEntity(peo);

                    if (per.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nFailed to select a line.");
                        return;
                    }

                    Line line = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Line;
                    if (line == null)
                    {
                        ed.WriteMessage("\nFailed to get the selected line.");
                        return;
                    }

                    // Create vertical lines from the endpoints of the selected line
                    Point3d pt1 = line.StartPoint;
                    Point3d pt2 = line.EndPoint;
                    Point3d pt1Top = new Point3d(pt1.X, pt1.Y, pt1.Z + surfaceHeight);
                    Point3d pt2Top = new Point3d(pt2.X, pt2.Y, pt2.Z + surfaceHeight);

                    // Create boundary lines for the surface
                    Line line1 = new Line(pt1, pt1Top);
                    Line line2 = new Line(pt2, pt2Top);
                    Line line3 = new Line(pt1Top, pt2Top);
                    Line line4 = new Line(pt1, pt2);

                    btr.AppendEntity(line1);
                    tr.AddNewlyCreatedDBObject(line1, true);

                    btr.AppendEntity(line2);
                    tr.AddNewlyCreatedDBObject(line2, true);

                    btr.AppendEntity(line3);
                    tr.AddNewlyCreatedDBObject(line3, true);

                    btr.AppendEntity(line4);
                    tr.AddNewlyCreatedDBObject(line4, true);

                    // Create a surface from the boundary lines
                    DBObjectCollection lines = new DBObjectCollection();
                    lines.Add(line1);
                    lines.Add(line2);
                    lines.Add(line3);
                    lines.Add(line4);

                    DBObjectCollection regions = Region.CreateFromCurves(lines);
                    if (regions.Count > 0)
                    {
                        using (Region region = regions[0] as Region)
                        {
                            using (PlaneSurface surface = new PlaneSurface())
                            {
                                surface.CreateFromRegion(region);
                                surface.ColorIndex = 6; // Pink color

                                btr.AppendEntity(surface);
                                tr.AddNewlyCreatedDBObject(surface, true);
                            }
                        }
                    }

                    ed.WriteMessage("\nCreated the surface.");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nException: " + ex.Message);
            }
        }

        [CommandMethod("CPI")]
        public void CurvePlaneIntersection()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
                return;

            var db = doc.Database;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect a curve");
            peo.SetRejectMessage("Must be a curve.");
            peo.AddAllowedClass(typeof(Curve), false);
            var per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
                return;

            var curId = per.ObjectId;

            var peo2 = new PromptEntityOptions("\nSelect plane surface");
            peo2.SetRejectMessage("Must be a planar surface.");
            peo2.AddAllowedClass(typeof(PlaneSurface), false);
            var per2 = ed.GetEntity(peo2);

            if (per2.Status != PromptStatus.OK)
                return;

            var planeId = per2.ObjectId;

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var plane = tr.GetObject(planeId, OpenMode.ForRead) as PlaneSurface;
                if (plane != null)
                {
                    var p = plane.GetPlane();
                    var cur = tr.GetObject(curId, OpenMode.ForRead) as Curve;
                    if (cur != null)
                    {
                        var pts = p.IntersectWith(cur);
                        var ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                        foreach (Point3d pt in pts)
                        {
                            var dbp = new DBPoint(pt);
                            dbp.ColorIndex = 2; // Make them yellow
                            ms.AppendEntity(dbp);
                            tr.AddNewlyCreatedDBObject(dbp, true);

                            ed.WriteMessage("\nPoint: " + pt.ToString());
                            ed.WriteMessage("\nPoint is " + (cur.IsOn(pt) ? "on" : "off") + " curve.");
                        }
                    }
                }
                tr.Commit();
            }
        }
    }

    public static class Extensions
    {
        ///<summary>
        /// Returns an array of Point3d objects from a Point3dCollection.
        ///</summary>
        ///<returns>An array of Point3d objects.</returns>
        public static Point3d[] ToArray(this Point3dCollection pts)
        {
            var res = new Point3d[pts.Count];
            pts.CopyTo(res, 0);
            return res;
        }

        ///<summary>
        /// Get the intersection points between this planar entity and a curve.
        ///</summary>
        ///<param name="cur">The curve to check intersections against.</param>
        ///<returns>An array of Point3d intersections.</returns>
        public static Point3d[] IntersectWith(this Plane p, Curve cur)
        {
            var pts = new Point3dCollection();
            var gcur = cur.GetGeCurve();
            var proj = gcur.GetProjectedEntity(p, p.Normal) as Curve3d;
            if (proj != null)
            {
                using (var gcur2 = Curve.CreateFromGeCurve(proj))
                {
                    cur.IntersectWith(gcur2, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);
                }
            }
            return pts.ToArray();
        }

        ///<summary>
        /// Test whether a point is on this curve.
        ///</summary>
        ///<param name="pt">The point to check against this curve.</param>
        ///<returns>Boolean indicating whether the point is on the curve.</returns>
        public static bool IsOn(this Curve cv, Point3d pt)
        {
            try
            {
                var p = cv.GetClosestPointTo(pt, false);
                return (p - pt).Length <= Tolerance.Global.EqualPoint;
            }
            catch
            {
                // Ignoring the exception and returning false
                return false;
            }
        }
    }
}
