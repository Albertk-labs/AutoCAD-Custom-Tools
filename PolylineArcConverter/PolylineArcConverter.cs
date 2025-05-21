using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

[assembly: CommandClass(typeof(PolylineArcConverter))]

public class PolylineArcConverter
{
    private double bulgeValue = 0.01;
    private static double tolerance = 0.001;
    private static Random random = new Random();

    [CommandMethod("PAC")]
    public void ShowMenu()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        while (true)
        {
            PromptKeywordOptions pko = new PromptKeywordOptions("\nSelect option [S1/S2/T/Quit]: ");
            pko.Keywords.Add("S1");
            pko.Keywords.Add("S2");
            pko.Keywords.Add("T", "Tolerance", "Set_Tolerance");
            pko.Keywords.Add("Quit");
            pko.AllowNone = true;

            PromptResult pr = ed.GetKeywords(pko);

            if (pr.Status != PromptStatus.OK || pr.StringResult == "Quit")
            {
                ed.WriteMessage("\nExiting...");
                break;
            }

            switch (pr.StringResult)
            {
                case "S1":
                    CreateArcs();
                    break;
                case "S2":
                    ConvertToArcs();
                    break;
                case "T":
                    ConfigureTolerance();
                    break;
            }
        }
    }

    private void CreateArcs()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        PromptEntityOptions peo = new PromptEntityOptions("\nSelect polyline: ");
        peo.SetRejectMessage("\nOnly polylines are allowed.");
        peo.AddAllowedClass(typeof(Polyline), true);

        PromptEntityResult per = ed.GetEntity(peo);
        if (per.Status != PromptStatus.OK)
            return;

        PromptIntegerOptions pio = new PromptIntegerOptions("\nEnter number of divisions (minimum 2): ");
        pio.LowerLimit = 2;
        pio.UseDefaultValue = true;
        pio.DefaultValue = 2;

        PromptIntegerResult pir = ed.GetInteger(pio);
        if (pir.Status != PromptStatus.OK)
            return;

        int divisions = pir.Value;

        PromptKeywordOptions pko = new PromptKeywordOptions("\nChoose arc direction [Left/Right]: ");
        pko.Keywords.Add("Left");
        pko.Keywords.Add("Right");

        PromptResult pr = ed.GetKeywords(pko);
        if (pr.Status != PromptStatus.OK)
            return;

        bool isLeftSide = pr.StringResult == "Left";

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            Polyline polyline = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Polyline;

            if (polyline != null)
            {
                Point3dCollection points = new Point3dCollection();
                for (int i = 0; i <= divisions; i++)
                {
                    double param = i / (double)divisions;
                    Point3d point = polyline.GetPointAtDist(polyline.Length * param);
                    points.Add(point);
                }

                polyline.Erase();

                for (int i = 0; i < points.Count - 1; i++)
                {
                    Point3d p1 = points[i];
                    Point3d p2 = points[i + 1];

                    using (Polyline newPolyline = new Polyline())
                    {
                        double actualBulge = isLeftSide ? bulgeValue : -bulgeValue;
                        newPolyline.AddVertexAt(0, new Point2d(p1.X, p1.Y), actualBulge, 0, 0);
                        newPolyline.AddVertexAt(1, new Point2d(p2.X, p2.Y), 0, 0, 0);

                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        btr.AppendEntity(newPolyline);
                        tr.AddNewlyCreatedDBObject(newPolyline, true);
                    }
                }
            }

            tr.Commit();
        }

        ed.WriteMessage("\nConversion completed.");
    }

    private void ConvertToArcs()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        ed.WriteMessage($"\nCurrent tolerance: {tolerance:F4}");

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                PromptEntityResult result = ed.GetEntity("\nSelect polyline: ");
                if (result.Status != PromptStatus.OK) return;

                Entity ent = tr.GetObject(result.ObjectId, OpenMode.ForRead) as Entity;
                if (!(ent is Polyline))
                {
                    ed.WriteMessage("\nSelected object is not a polyline.");
                    return;
                }

                Polyline pline = ent as Polyline;
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                List<Point3d> points = new List<Point3d>();
                for (int i = 0; i < pline.NumberOfVertices; i++)
                {
                    points.Add(pline.GetPoint3dAt(i));
                }

                for (int i = 0; i < points.Count - 2; i += 2)
                {
                    Point3d startPoint = points[i];
                    Point3d midPoint = points[i + 1];
                    Point3d endPoint = points[i + 2];

                    Arc arc = TryFitArc(startPoint, midPoint, endPoint, tolerance);
                    if (arc != null)
                    {
                        arc.ColorIndex = random.Next(1, 255);
                        btr.AppendEntity(arc);
                        tr.AddNewlyCreatedDBObject(arc, true);
                    }
                    else
                    {
                        Line line1 = new Line(startPoint, midPoint);
                        Line line2 = new Line(midPoint, endPoint);
                        btr.AppendEntity(line1);
                        btr.AppendEntity(line2);
                        tr.AddNewlyCreatedDBObject(line1, true);
                        tr.AddNewlyCreatedDBObject(line2, true);
                    }
                }

                pline.UpgradeOpen();
                pline.Erase();

                tr.Commit();
                ed.WriteMessage("\nPolyline was converted into arcs and lines.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError: " + ex.Message);
                tr.Abort();
            }
        }
    }

    private Arc TryFitArc(Point3d startPoint, Point3d midPoint, Point3d endPoint, double tolerance)
    {
        try
        {
            Vector3d v1 = midPoint - startPoint;
            Vector3d v2 = endPoint - startPoint;

            Vector3d normal = v1.CrossProduct(v2);
            if (normal.Length < tolerance) return null;

            Point3d center = CalculateCircleCenter(startPoint, midPoint, endPoint);
            if (center.DistanceTo(Point3d.Origin) < tolerance) return null;

            double radius = center.DistanceTo(startPoint);
            if (Math.Abs(center.DistanceTo(midPoint) - radius) > tolerance ||
                Math.Abs(center.DistanceTo(endPoint) - radius) > tolerance)
            {
                return null;
            }

            Vector3d startVector = startPoint - center;
            Vector3d endVector = endPoint - center;
            double startAngle = Math.Atan2(startVector.Y, startVector.X);
            double endAngle = Math.Atan2(endVector.Y, endVector.X);

            return new Arc(center, normal, radius, startAngle, endAngle);
        }
        catch
        {
            return null;
        }
    }

    private Point3d CalculateCircleCenter(Point3d p1, Point3d p2, Point3d p3)
    {
        double offset = Math.Pow(p2.X, 2) + Math.Pow(p2.Y, 2);
        double bc = (Math.Pow(p1.X, 2) + Math.Pow(p1.Y, 2) - offset) / 2.0;
        double cd = (offset - Math.Pow(p3.X, 2) - Math.Pow(p3.Y, 2)) / 2.0;
        double det = (p1.X - p2.X) * (p2.Y - p3.Y) - (p2.X - p3.X) * (p1.Y - p2.Y);

        if (Math.Abs(det) < 1e-6) return Point3d.Origin;

        double centerX = (bc * (p2.Y - p3.Y) - cd * (p1.Y - p2.Y)) / det;
        double centerY = ((p1.X - p2.X) * cd - (p2.X - p3.X) * bc) / det;

        return new Point3d(centerX, centerY, 0);
    }

    private void ConfigureTolerance()
    {
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

        PromptDoubleOptions pdo = new PromptDoubleOptions($"\nEnter new tolerance value (current: {tolerance:F4}):");
        pdo.DefaultValue = tolerance;
        pdo.AllowZero = false;
        pdo.AllowNegative = false;
        pdo.UseDefaultValue = true;

        PromptDoubleResult result = ed.GetDouble(pdo);

        if (result.Status == PromptStatus.OK)
        {
            tolerance = result.Value;
            ed.WriteMessage($"\nTolerance updated to: {tolerance:F4}");
        }
    }
}
