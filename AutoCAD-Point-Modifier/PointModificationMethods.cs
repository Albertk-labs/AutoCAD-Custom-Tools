using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

public static class PointModificationMethods
{
    public static (List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)>, List<double>, int) SelectAndTranslatePoints(Editor ed, Document doc, int currentSetNumber, double translationDistance)
    {
        PromptSelectionOptions pso = new PromptSelectionOptions();
        pso.MessageForAdding = $"\nSelect points for set {currentSetNumber}: ";
        PromptSelectionResult selRes = ed.GetSelection(pso);

        if (selRes.Status != PromptStatus.OK)
        {
            return (new List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)>(), new List<double>(), currentSetNumber);
        }

        List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)> selectedPoints = new List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)>();
        List<double> distances = new List<double>();



        using (Transaction tr = doc.TransactionManager.StartTransaction())
        {
            SelectionSet ss = selRes.Value;
            int pointCounter = 1;

            foreach (SelectedObject obj in ss)
            {
                if (obj.ObjectId.ObjectClass.DxfName == "POINT")
                {
                    DBPoint point = tr.GetObject(obj.ObjectId, OpenMode.ForRead) as DBPoint;
                    if (point != null)
                    {
                        Point3d originalPosition = point.Position;
                        Vector3d translationVector = new Vector3d(-translationDistance, 0, 0);
                        Point3d modifiedPosition = originalPosition + translationVector;

                        string label = $"P{currentSetNumber}.{pointCounter}";
                        selectedPoints.Add((label, originalPosition, modifiedPosition));
                        distances.Add(Math.Abs(translationDistance));

                        pointCounter++;
                    }
                }
            }

            tr.Commit();
        }

        return (selectedPoints, distances, currentSetNumber + 1);
    }

    public static bool ConfirmZeroZPoint(Editor ed, Point3d point)
    {
        if (Math.Abs(point.Z) < 1e-6)  // Check if Z is approximately zero
        {
            PromptKeywordOptions pko = new PromptKeywordOptions($"\nSelected point has Z=0 ({point}). Are you sure? [Yes/No]: ", "Yes No");
            pko.AllowNone = false;
            PromptResult res = ed.GetKeywords(pko);
            return res.StringResult == "Yes";
        }
        return true;
    }
}
