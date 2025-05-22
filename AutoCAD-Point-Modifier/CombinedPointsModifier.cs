using System;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using OfficeOpenXml;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using System.Diagnostics;
using System.Collections.Generic;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Interop.Excel;

namespace AutoCADExport
{
    public class CombinedPointsModifier
    {
        private List<(string ChapterName, List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)> Points, double TranslationDistance)> chapters = new List<(string, List<(string, Point3d, Point3d)>, double)>();
        private List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)> allPoints = new List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)>();
        private List<double> distancesList = new List<double>();
        private static int setCounter = 1;
        private double? lastUsedTranslationDistance = null;

        [CommandMethod("ModifyPoints")]
        public void ModifyPointsCommand()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            while (true)
            {
                PromptKeywordOptions pko = new PromptKeywordOptions("\nChoose option [Select(S)/Modify(M)/Project(P)/Export(E)/Draw(D)/Next(N)/Reset(R)/Status(T)/NameSet(A)/Exit(X)]: ");
                pko.Keywords.Add("S");
                pko.Keywords.Add("M");
                pko.Keywords.Add("P");
                pko.Keywords.Add("E");
                pko.Keywords.Add("D");
                pko.Keywords.Add("N");
                pko.Keywords.Add("R");
                pko.Keywords.Add("T");
                pko.Keywords.Add("A");
                pko.Keywords.Add("X");
                pko.AllowNone = false;

                PromptResult pRes = ed.GetKeywords(pko);

                if (pRes.Status != PromptStatus.OK || pRes.StringResult == "X")
                {
                    ed.WriteMessage("\nExiting program.");
                    return;
                }

                switch (pRes.StringResult)
                {
                    case "S":
                        SelectPoints(ed, doc);
                        break;
                    case "M":
                        ModifyNamedSet(ed, doc);
                        break;
                    case "P":
                        ProjectPoints(ed, doc);
                        break;
                    case "E":
                        ExportToExcel(ed);
                        break;
                    case "D":
                        DrawPoints(ed, doc);
                        break;
                    case "N":
                        setCounter += 2;
                        ed.WriteMessage($"\nMoved to the next set. Current set number: {setCounter}");
                        break;
                    case "R":
                        ResetAll(ed);
                        break;
                    case "T":
                        ShowCurrentState(ed);
                        break;
                    case "A":
                        NameCurrentSet(ed);
                        break;
                }
            }
        }

        public void SelectPoints(Editor ed, Document doc)
        {
            try
            {
                var options = new PromptSelectionOptions
                {
                    MessageForAdding = "\nSelect points:"
                };
                var selectionResult = ed.GetSelection(options);

                if (selectionResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nPoint selection cancelled.");
                    return;
                }

                var selectionSet = selectionResult.Value;

                var lineOptions = new PromptEntityOptions("\nSelect reference line:");
                lineOptions.SetRejectMessage("\nPlease select a line.");
                lineOptions.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Line), true);
                var lineResult = ed.GetEntity(lineOptions);

                if (lineResult.Status != PromptStatus.OK) return;

                var distanceOptions = new PromptDoubleOptions("\nEnter translation distance (positive = right, negative = left):");
                var distanceResult = ed.GetDouble(distanceOptions);

                if (distanceResult.Status != PromptStatus.OK) return;

                var translationDistance = distanceResult.Value;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.Line referenceLine = tr.GetObject(lineResult.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Line;
                    if (referenceLine == null)
                    {
                        ed.WriteMessage("\nSelected object is not a line.");
                        return;
                    }

                    Vector3d lineDirection = (referenceLine.EndPoint - referenceLine.StartPoint).GetNormal();
                    Vector3d normalVector = new Vector3d(-lineDirection.Y, lineDirection.X, 0).GetNormal();

                    List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)> currentSetPoints = new List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)>();

                    foreach (SelectedObject selectedObject in selectionSet)
                    {
                        var point = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as DBPoint;
                        if (point != null)
                        {
                            Point3d originalPosition = point.Position;
                            Point3d modifiedPosition = originalPosition + normalVector * translationDistance;
                            string label = $"Point {setCounter}";
                            currentSetPoints.Add((label, originalPosition, modifiedPosition));
                            allPoints.Add((label, originalPosition, modifiedPosition));
                            setCounter++;
                        }
                    }

                    chapters.Add(($"Chapter {chapters.Count + 1}", currentSetPoints, translationDistance));

                    tr.Commit();

                    ed.WriteMessage($"\nAdded {currentSetPoints.Count} points to a new chapter.");
                    ed.WriteMessage($"\nTranslation distance for this set: {translationDistance}");
                }

                lastUsedTranslationDistance = translationDistance;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }
        public void DrawPoints(Editor ed, Document doc)
        {
            if (chapters.Count == 0)
            {
                ed.WriteMessage("\nNo points to draw. Please select points first.");
                return;
            }

            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (var chapter in chapters)
                {
                    foreach (var (_, _, modifiedPosition) in chapter.Points)
                    {
                        DBPoint point = new DBPoint(modifiedPosition);
                        point.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3); // Green
                        btr.AppendEntity(point);
                        tr.AddNewlyCreatedDBObject(point, true);
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nPoints drawn for {chapters.Count} chapters.");
            doc.Editor.Regen();
        }
        public void ModifyPoints(Editor ed, Document doc, int chapterIndex)
        {
            var (chapterName, chapterPoints, translationDistance) = chapters[chapterIndex];

            PromptDoubleOptions pdoSlope = new PromptDoubleOptions("\nEnter slope (%) value: ");
            PromptDoubleResult pdrSlope = ed.GetDouble(pdoSlope);
            if (pdrSlope.Status != PromptStatus.OK) return;
            double slope = pdrSlope.Value;

            PromptDoubleOptions pdoHeight = new PromptDoubleOptions("\nEnter height value: ");
            PromptDoubleResult pdrHeight = ed.GetDouble(pdoHeight);
            if (pdrHeight.Status != PromptStatus.OK) return;
            double height = pdrHeight.Value;

            List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)> modifiedPoints = new List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)>();

            ed.WriteMessage($"\nModifying points in chapter: {chapterName}");
            ed.WriteMessage($"\nSlope: {slope}%, Height: {height}, Translation Distance: {translationDistance}");

            foreach (var (label, originalPosition, _) in chapterPoints)
            {
                double newZ = CalculateModifiedZ(slope, translationDistance, height, originalPosition.Z);
                Point3d newModifiedPosition = new Point3d(originalPosition.X + translationDistance, originalPosition.Y, newZ);
                modifiedPoints.Add((label, originalPosition, newModifiedPosition));

                ed.WriteMessage($"\nPoint {label}:");
                ed.WriteMessage($"  Original: X={originalPosition.X:F3}, Y={originalPosition.Y:F3}, Z={originalPosition.Z:F3}");
                ed.WriteMessage($"  Modified: X={newModifiedPosition.X:F3}, Y={newModifiedPosition.Y:F3}, Z={newModifiedPosition.Z:F3}");
            }

            chapters[chapterIndex] = (chapterName, modifiedPoints, translationDistance);

            for (int i = 0; i < allPoints.Count; i++)
            {
                var (label, originalPosition, _) = allPoints[i];
                var modifiedPoint = modifiedPoints.FirstOrDefault(p => p.Label == label);
                if (modifiedPoint != default)
                {
                    allPoints[i] = modifiedPoint;
                }
            }

            ed.WriteMessage($"\nModified {modifiedPoints.Count} points in chapter {chapterName}.");
            ed.WriteMessage("\nModification Summary:");
            ed.WriteMessage($"Chapter: {chapterName}");
            ed.WriteMessage($"Total modified points: {modifiedPoints.Count}");
            ed.WriteMessage($"Slope: {slope}%");
            ed.WriteMessage($"Height: {height}");
            ed.WriteMessage($"Translation Distance: {translationDistance}");
        }
        private double CalculateModifiedZ(double slopePercent, double translationDistance, double height, double originalZ)
        {
            return (slopePercent / 100.0) * Math.Abs(translationDistance) + height + originalZ;
        }
        public void ProjectPoints(Editor ed, Document doc)
        {
            if (allPoints.Count == 0)
            {
                ed.WriteMessage("\nNo points to project. Please select points first.");
                return;
            }

            PromptKeywordOptions pko = new PromptKeywordOptions("\nSelect projection plane [XY/XZ/YZ]: ", "XY XZ YZ");
            pko.AllowNone = false;
            PromptResult res = ed.GetKeywords(pko);

            if (res.Status != PromptStatus.OK) return;

            string plane = res.StringResult;

            EraseDrawnPoints(ed, doc);

            for (int i = 0; i < allPoints.Count; i++)
            {
                var (label, originalPosition, modifiedPosition) = allPoints[i];
                Point3d projectedPosition = ProjectPoint(modifiedPosition, plane);
                allPoints[i] = (label, originalPosition, projectedPosition);
                DrawingMethods.DrawModifiedPoint(doc, projectedPosition);
            }

            ed.WriteMessage($"\nProjected and drawn {allPoints.Count} points onto the {plane} plane.");
        }

        private Point3d ProjectPoint(Point3d point, string plane)
        {
            switch (plane)
            {
                case "XY":
                    return new Point3d(point.X, point.Y, 0);
                case "XZ":
                    return new Point3d(point.X, 0, point.Z);
                case "YZ":
                    return new Point3d(0, point.Y, point.Z);
                default:
                    return point;
            }
        }
        public void ExportToExcel(Editor ed)
        {
            if (chapters.Count == 0)
            {
                ed.WriteMessage("\nNo chapters to export. Please select points first.");
                return;
            }

            ed.WriteMessage("\nExport options:");
            ed.WriteMessage("\n0. Export all chapters to a single sheet");
            for (int i = 0; i < chapters.Count; i++)
            {
                ed.WriteMessage($"\n{i + 1}. {chapters[i].ChapterName}");
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nSelect export option number: ")
            {
                LowerLimit = 0,
                UpperLimit = chapters.Count
            };
            PromptIntegerResult result = ed.GetInteger(pio);

            if (result.Status != PromptStatus.OK) return;

            int selectedOption = result.Value;

            try
            {
                string baseFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ExportedPoints");
                string filePath = baseFilePath + ".xlsx";
                int counter = 1;

                while (File.Exists(filePath))
                {
                    filePath = baseFilePath + $"_{counter}.xlsx";
                    counter++;
                }

                var fileInfo = new FileInfo(filePath);

                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Exported Points");

                    int currentRow = 1;

                    if (selectedOption == 0)
                    {
                        foreach (var chapter in chapters)
                        {
                            currentRow = ExportChapterToWorksheet(worksheet, chapter, currentRow);
                            currentRow++;
                        }
                        ed.WriteMessage($"\nAll chapters exported to {filePath}");
                    }
                    else
                    {
                        var selectedChapter = chapters[selectedOption - 1];
                        ExportChapterToWorksheet(worksheet, selectedChapter, currentRow);
                        ed.WriteMessage($"\nChapter {selectedChapter.ChapterName} exported to {filePath}");
                    }

                    package.Save();
                }

                ed.WriteMessage($"\nExport successful: {filePath}");
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during export: {ex.Message}");
            }
        }
        private int ExportChapterToWorksheet(ExcelWorksheet worksheet, (string ChapterName, List<(string Label, Point3d OriginalPosition, Point3d ModifiedPosition)> Points, double TranslationDistance) chapter, int startRow)
        {
            worksheet.Cells[startRow, 1].Value = chapter.ChapterName;
            worksheet.Cells[startRow, 1, startRow, 8].Merge = true;
            worksheet.Cells[startRow, 1].Style.Font.Bold = true;
            startRow++;

            worksheet.Cells[startRow, 1].Value = "Point Name";
            worksheet.Cells[startRow, 2].Value = "Original X";
            worksheet.Cells[startRow, 3].Value = "Original Y";
            worksheet.Cells[startRow, 4].Value = "Original Z";
            worksheet.Cells[startRow, 5].Value = "Modified X";
            worksheet.Cells[startRow, 6].Value = "Modified Y";
            worksheet.Cells[startRow, 7].Value = "Modified Z";
            worksheet.Cells[startRow, 8].Value = "Translation Distance";
            startRow++;

            foreach (var (label, originalPosition, modifiedPosition) in chapter.Points)
            {
                worksheet.Cells[startRow, 1].Value = label;
                worksheet.Cells[startRow, 2].Value = originalPosition.X;
                worksheet.Cells[startRow, 3].Value = originalPosition.Y;
                worksheet.Cells[startRow, 4].Value = originalPosition.Z;
                worksheet.Cells[startRow, 5].Value = modifiedPosition.X;
                worksheet.Cells[startRow, 6].Value = modifiedPosition.Y;
                worksheet.Cells[startRow, 7].Value = modifiedPosition.Z;
                worksheet.Cells[startRow, 8].Value = chapter.TranslationDistance;
                startRow++;
            }

            return startRow;
        }
        private void ResetAll(Editor ed)
        {
            PromptKeywordOptions pko = new PromptKeywordOptions("\nWhat do you want to reset? [All(A)/Selected(S)/Cancel(C)]: ", "All Selected Cancel");
            pko.AllowNone = false;
            PromptResult res = ed.GetKeywords(pko);

            switch (res.StringResult)
            {
                case "All":
                    setCounter = 1;
                    allPoints.Clear();
                    distancesList.Clear();
                    chapters.Clear();
                    ed.WriteMessage("\nAll data has been cleared. Ready to start fresh.");
                    break;
                case "Selected":
                    ClearSpecificChapter(ed);
                    break;
                case "Cancel":
                    ed.WriteMessage("\nReset operation canceled.");
                    break;
            }
        }
        private void ClearSpecificChapter(Editor ed)
        {
            if (chapters.Count == 0)
            {
                ed.WriteMessage("\nNo chapters available to clear.");
                return;
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nEnter the chapter number to clear: ");
            pio.LowerLimit = 1;
            pio.UpperLimit = chapters.Count;
            PromptIntegerResult result = ed.GetInteger(pio);

            if (result.Status == PromptStatus.OK)
            {
                int chapterToRemove = result.Value - 1;
                chapters.RemoveAt(chapterToRemove);
                ed.WriteMessage($"\nChapter {result.Value} has been cleared.");
            }
        }
        private void ShowCurrentState(Editor ed)
        {
            ed.WriteMessage($"\nCurrent state:");
            ed.WriteMessage($"\nNumber of chapters: {chapters.Count}");
            ed.WriteMessage($"\nTotal number of points: {allPoints.Count}");
            ed.WriteMessage($"\nNext set number: {setCounter}");
        }
        private void NameCurrentSet(Editor ed)
        {
            if (chapters.Count == 0)
            {
                ed.WriteMessage("\nNo chapters to rename. Please select points first.");
                return;
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nEnter the chapter number to rename: ");
            pio.LowerLimit = 1;
            pio.UpperLimit = chapters.Count;
            PromptIntegerResult result = ed.GetInteger(pio);

            if (result.Status != PromptStatus.OK) return;

            int chapterIndex = result.Value - 1;

            PromptStringOptions pso = new PromptStringOptions("\nEnter new name for the chapter: ");
            pso.AllowSpaces = true;
            PromptResult nameResult = ed.GetString(pso);

            if (nameResult.Status == PromptStatus.OK)
            {
                string newName = nameResult.StringResult;
                var (_, points, distance) = chapters[chapterIndex];
                chapters[chapterIndex] = (newName, points, distance);
                ed.WriteMessage($"\nChapter {result.Value} renamed to: {newName}");
            }
        }
        private void ModifyNamedSet(Editor ed, Document doc)
        {
            if (chapters.Count == 0)
            {
                ed.WriteMessage("\nNo chapters available for modification.");
                return;
            }

            ed.WriteMessage("\nAvailable chapters:");
            for (int i = 0; i < chapters.Count; i++)
            {
                ed.WriteMessage($"\n{i + 1}. {chapters[i].ChapterName}");
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nEnter the chapter number to modify: ")
            {
                LowerLimit = 1,
                UpperLimit = chapters.Count
            };
            PromptIntegerResult result = ed.GetInteger(pio);

            if (result.Status == PromptStatus.OK)
            {
                int chapterIndex = result.Value - 1;
                var (chapterName, chapterPoints, chapterTranslationDistance) = chapters[chapterIndex];
                ed.WriteMessage($"\nSelected chapter: {result.Value}. {chapterName}. You can now modify it.");
                ed.WriteMessage($"\nNumber of points in selected chapter: {chapterPoints.Count}");
                ed.WriteMessage($"\nTranslation distance for this chapter: {chapterTranslationDistance}");

                ModifyPoints(ed, doc, chapterIndex);
            }
        }
        private void EraseDrawnPoints(Editor ed, Document doc)
        {
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                var pointsToErase = new List<ObjectId>();

                foreach (ObjectId objId in btr)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    if (ent is DBPoint)
                    {
                        pointsToErase.Add(objId);
                    }
                }

                foreach (var objId in pointsToErase)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    ent.Erase();
                }

                tr.Commit();
            }

            ed.WriteMessage("\nAll previously drawn points have been erased.");
        }

    }
}
