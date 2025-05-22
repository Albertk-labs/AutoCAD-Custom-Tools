using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

public static class DrawingMethods
{
    private static ObjectId modifiedPointsGroupId = ObjectId.Null;

    // Removes all previously drawn modified points
    public static void EraseModifiedPoints(Document doc)
    {
        if (modifiedPointsGroupId != ObjectId.Null)
        {
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                Group group = tr.GetObject(modifiedPointsGroupId, OpenMode.ForWrite) as Group;
                if (group != null)
                {
                    foreach (ObjectId id in group.GetAllEntityIds())
                    {
                        DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
                        obj.Erase();
                    }
                    group.Erase();
                }
                tr.Commit();
            }
            modifiedPointsGroupId = ObjectId.Null;
        }
    }

    // Draws a modified point in the model space with red color and adds it to a group
    public static void DrawModifiedPoint(Document doc, Point3d point)
    {
        using (Transaction tr = doc.TransactionManager.StartTransaction())
        {
            BlockTable bt = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            DBPoint dbPoint = new DBPoint(point);
            dbPoint.SetDatabaseDefaults();
            dbPoint.ColorIndex = 1; // Red color for modified points

            ObjectId pointId = btr.AppendEntity(dbPoint);
            tr.AddNewlyCreatedDBObject(dbPoint, true);

            DBDictionary groupDict = tr.GetObject(doc.Database.GroupDictionaryId, OpenMode.ForRead) as DBDictionary;

            if (modifiedPointsGroupId == ObjectId.Null)
            {
                Group group = new Group("ModifiedPointsGroup", true);
                groupDict.UpgradeOpen();
                modifiedPointsGroupId = groupDict.SetAt("ModifiedPointsGroup", group);
                tr.AddNewlyCreatedDBObject(group, true);
            }

            Group modifiedPointsGroup = tr.GetObject(modifiedPointsGroupId, OpenMode.ForWrite) as Group;
            modifiedPointsGroup.InsertAt(modifiedPointsGroup.NumEntities, pointId);

            tr.Commit();
        }
    }

    // Calculates new position based on direction vector and translation distance
    public static Point3d CalculateModifiedPosition(Point3d originalPosition, double translationDistance, Vector3d direction)
    {
        Vector3d translationVector = direction.GetNormal() * translationDistance;
        return originalPosition + translationVector;
    }
}
