using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace S3
{
    public class Commands
    {
        private static bool DrawVisibleAreas = false;

        [CommandMethod("TM")]
        public static void MainMenu()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("\nChoose an option:");
            pKeyOpts.Keywords.Add("Run");
            pKeyOpts.Keywords.Add("Configure");
            pKeyOpts.AllowNone = false;

            PromptResult pKeyRes = ed.GetKeywords(pKeyOpts);

            if (pKeyRes.Status == PromptStatus.OK)
            {
                switch (pKeyRes.StringResult)
                {
                    case "Run":
                        RunMainFunction();
                        break;
                    case "Configure":
                        Configure();
                        break;
                }
            }
        }

        private static void Configure()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("\nSelect area drawing mode:");
            pKeyOpts.Keywords.Add("Invisible");
            pKeyOpts.Keywords.Add("Visible");
            pKeyOpts.AllowNone = false;

            PromptResult pKeyRes = ed.GetKeywords(pKeyOpts);

            if (pKeyRes.Status == PromptStatus.OK)
            {
                DrawVisibleAreas = pKeyRes.StringResult == "Visible";
                ed.WriteMessage($"\nSelected to draw {(DrawVisibleAreas ? "visible" : "invisible")} areas.");
            }
        }
            private static void RunMainFunction()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                using (DocumentLock docLock = doc.LockDocument())
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    string sourceFilePath = @"C:\Projekty\TME1.xlsx";
                    string resultFilePath = @"C:\Projekty\RevitDane.xlsx";

                    if (!File.Exists(sourceFilePath))
                    {
                        ed.WriteMessage($"\nSource file does not exist: {sourceFilePath}");
                        return;
                    }

                    ed.WriteMessage($"\nLoading source file: {sourceFilePath}");
                    using var excelData = LoadExcelData(sourceFilePath);

                    ed.WriteMessage($"\nCreating result file: {resultFilePath}");
                    using var resultExcel = CreateResultExcel(resultFilePath);

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Dictionary<string, List<(string SectionNumber, double Coefficient)>> wallIdToSections = new();

                        var uniqueWallIds = FindUniqueWallIds(db, tr);
                        var textEntities = FindTextEntities(db, tr);

                        ed.WriteMessage($"\nFound {uniqueWallIds.Count} unique wall IDs.");
                        ed.WriteMessage($"\nFound {textEntities.Count} text entities.");

                        foreach (var wallId in uniqueWallIds)
                        {
                            Entity wall = (Entity)tr.GetObject(wallId, OpenMode.ForRead);
                            string wallIdText = GetTextString(wall);
                            var intersectingTexts = FindIntersectingTexts(wall, textEntities);

                            ed.WriteMessage($"\nWall {wallIdText}: found {intersectingTexts.Count} intersecting texts.");
                            foreach (var text in intersectingTexts)
                            {
                                string textContent = GetTextString(text);
                                ed.WriteMessage($"\n  - Text: {textContent}");
                                var sectionInfo = FindSectionInfo(excelData, textContent);
                                if (sectionInfo.HasValue)
                                {
                                    ed.WriteMessage($"\n    Section found: {sectionInfo.Value.SectionNumber}, Coefficient: {sectionInfo.Value.Coefficient}");
                                    if (!wallIdToSections.ContainsKey(wallIdText))
                                        wallIdToSections[wallIdText] = new();
                                    wallIdToSections[wallIdText].Add(sectionInfo.Value);
                                }
                                else
                                {
                                    ed.WriteMessage($"\n    No section info found for text: {textContent}");
                                }
                            }

                            DrawAreas(db, tr, wall, intersectingTexts);
                        }

                        ed.WriteMessage($"\nSection data found for {wallIdToSections.Count} walls.");
                        SaveToResultExcel(resultExcel, excelData, wallIdToSections);

                        tr.Commit();
                    }

                    resultExcel.Save();
                    ed.WriteMessage($"\nOperation completed successfully. Data saved to: {resultFilePath}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError occurred: {ex.Message}");
                if (ex.InnerException != null)
                {
                    ed.WriteMessage($"\nInner Exception: {ex.InnerException.Message}");
                }
                ed.WriteMessage($"\nStack Trace: {ex.StackTrace}");
            }
        }
        private static ExcelPackage LoadExcelData(string filePath)
        {
            var file = new FileInfo(filePath);
            if (!file.Exists)
            {
                throw new FileNotFoundException($"Excel source file not found: {filePath}");
            }
            return new ExcelPackage(file);
        }
        private static ExcelPackage CreateResultExcel(string basePath)
        {
            string directory = Path.GetDirectoryName(basePath) ?? "";
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string path = basePath;
            int counter = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(
                    directory,
                    $"{Path.GetFileNameWithoutExtension(basePath)}_{counter}{Path.GetExtension(basePath)}"
                );
                counter++;
            }

            var package = new ExcelPackage(new FileInfo(path));

            if (package.Workbook.Worksheets.Count == 0)
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                CreateHeaders(worksheet);
            }

            return package;
        }
        private static List<ObjectId> FindUniqueWallIds(Database db, Transaction tr)
        {
            var uniqueWallIds = new HashSet<string>();
            var result = new List<ObjectId>();
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in modelSpace)
            {
                var entity = (Entity)tr.GetObject(id, OpenMode.ForRead);
                if (entity is MText mtext && mtext.Layer == "A-WALL-____-IDEN")
                {
                    string wallIdText = mtext.Contents.Trim();
                    if (!uniqueWallIds.Contains(wallIdText))
                    {
                        uniqueWallIds.Add(wallIdText);
                        result.Add(id);
                    }
                }
            }

            return result;
        }
        private static List<Entity> FindTextEntities(Database db, Transaction tr)
        {
            var textEntities = new List<Entity>();
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in modelSpace)
            {
                var entity = (Entity)tr.GetObject(id, OpenMode.ForRead);
                if (entity is DBText && entity.Layer == "PRT_MOD_TXT")
                {
                    textEntities.Add(entity);
                }
            }

            return textEntities;
        }
        private static List<Entity> FindIntersectingTexts(Entity wall, List<Entity> textEntities)
        {
            var intersectingTexts = new List<Entity>();
            var wallCenter = GetEntityCenter(wall);
            var wallExtents = wall.GeometricExtents;
            double searchRadius = 0.2; // Initial search radius
            double maxSearchRadius = 10.0; // Maximum search radius

            while (searchRadius <= maxSearchRadius)
            {
                bool foundAny = false;

                foreach (var text in textEntities)
                {
                    var textCenter = GetEntityCenter(text);
                    var textExtents = text.GeometricExtents;

                    if (wallCenter.DistanceTo(textCenter) <= searchRadius ||
                        (wallExtents.MinPoint.X <= textExtents.MaxPoint.X && wallExtents.MaxPoint.X >= textExtents.MinPoint.X) &&
                        (wallExtents.MinPoint.Y <= textExtents.MaxPoint.Y && wallExtents.MaxPoint.Y >= textExtents.MinPoint.Y) &&
                        (wallExtents.MinPoint.Z <= textExtents.MaxPoint.Z && wallExtents.MaxPoint.Z >= textExtents.MinPoint.Z))
                    {
                        if (!intersectingTexts.Contains(text))
                        {
                            intersectingTexts.Add(text);
                            foundAny = true;
                        }
                    }
                }

                if (foundAny)
                    break;

                searchRadius += 0.2; // Increase search radius by 0.2 m
            }

            return intersectingTexts;
        }
        private static void DrawAreas(Database db, Transaction tr, Entity wall, List<Entity> intersectingTexts)
        {
            var wallCenter = GetEntityCenter(wall);

            if (intersectingTexts.Count > 0)
            {
                var maxDistance = intersectingTexts.Max(text => wallCenter.DistanceTo(GetEntityCenter(text)));

                if (DrawVisibleAreas)
                {
                    DrawVisibleArea(db, tr, wallCenter, maxDistance);
                }
                else
                {
                    DrawInvisibleArea(db, tr, wallCenter, maxDistance);
                }
            }

            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                $"\nWall {GetTextString(wall)}: found {intersectingTexts.Count} intersecting texts.");
        }
        private static void DrawVisibleArea(Database db, Transaction tr, Point3d center, double radius)
        {
            var modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            using (var circle = new Circle(center, Vector3d.ZAxis, radius))
            {
                circle.ColorIndex = 1; // Red color
                modelSpace.AppendEntity(circle);
                tr.AddNewlyCreatedDBObject(circle, true);
            }
        }
        private static void DrawInvisibleArea(Database db, Transaction tr, Point3d center, double radius)
        {
            var modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            using (var circle = new Circle(center, Vector3d.ZAxis, radius))
            {
                circle.ColorIndex = 256; // BYLAYER color
                circle.LineWeight = LineWeight.ByLayer;
                modelSpace.AppendEntity(circle);
                tr.AddNewlyCreatedDBObject(circle, true);
            }
        }
        private static void CreateHeaders(ExcelWorksheet worksheet)
        {
            string[] headers = new string[] { "Section Number", "Basket 1", "Basket 2", "Basket 3", "Basket 4", "Coefficient", "Wall ID" };
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }
        }
        private static string GetTextString(Entity textEntity)
        {
            return textEntity switch
            {
                DBText dbText => dbText.TextString,
                MText mText => mText.Contents,
                _ => throw new ArgumentException("Unsupported text entity type")
            };
        }
        private static Point3d GetEntityCenter(Entity entity)
        {
            var bounds = entity.GeometricExtents;
            return new Point3d(
                (bounds.MinPoint.X + bounds.MaxPoint.X) / 2,
                (bounds.MinPoint.Y + bounds.MaxPoint.Y) / 2,
                (bounds.MinPoint.Z + bounds.MaxPoint.Z) / 2
            );
        }
        private static (string SectionNumber, double Coefficient)? FindSectionInfo(ExcelPackage excelData, string textContent)
        {
            var worksheet = excelData.Workbook.Worksheets.FirstOrDefault();
            if (worksheet != null)
            {
                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    for (int col = 2; col <= 5; col++)
                    {
                        if (worksheet.Cells[row, col].Text.Trim().Equals(textContent, StringComparison.OrdinalIgnoreCase))
                        {
                            string fullSectionNumber = worksheet.Cells[row, 1].Text.Trim();

                            if (double.TryParse(worksheet.Cells[row, 6].Text.Trim(), out double coefficient))
                            {
                                return (fullSectionNumber, coefficient); // <--- ZACHOWAJ pełną nazwę np. "P2 | P2.02"
                            }
                        }
                    }
                }
            }
            return null;
        }


        private static void SaveToResultExcel(ExcelPackage resultExcel, ExcelPackage sourceExcel, Dictionary<string, List<(string SectionNumber, double Coefficient)>> wallIdToSections)
        {
            var sourceWorksheet = sourceExcel.Workbook.Worksheets.FirstOrDefault();
            var resultWorksheet = resultExcel.Workbook.Worksheets.FirstOrDefault() ?? resultExcel.Workbook.Worksheets.Add("Sheet1");

            // Copy headers
            string[] headers = { "Section Number", "Basket 1", "Basket 2", "Basket 3", "Basket 4", "Coefficient", "Wall ID" };
            for (int i = 0; i < headers.Length; i++)
            {
                resultWorksheet.Cells[1, i + 1].Value = headers[i];
                resultWorksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            // Flatten and group by SectionNumber + Coefficient
            var groupedData = wallIdToSections
                .SelectMany(kvp => kvp.Value.Select(v => new
                {
                    WallId = kvp.Key,
                    SectionNumber = v.SectionNumber,
                    Coefficient = v.Coefficient
                }))
                .GroupBy(x => new { x.SectionNumber, x.Coefficient })
                .OrderBy(g => g.Key.Coefficient);

            int resultRow = 2;
            foreach (var group in groupedData)
            {
                int sourceRow = FindSourceRow(sourceWorksheet, group.Key.SectionNumber, group.Key.Coefficient);

                // Get baskets from source worksheet
                string basket1 = sourceRow > 0 ? sourceWorksheet.Cells[sourceRow, 2].Text : "";
                string basket2 = sourceRow > 0 ? sourceWorksheet.Cells[sourceRow, 3].Text : "";
                string basket3 = sourceRow > 0 ? sourceWorksheet.Cells[sourceRow, 4].Text : "";
                string basket4 = sourceRow > 0 ? sourceWorksheet.Cells[sourceRow, 5].Text : "";

                // Write to result
                string displaySection = group.Key.SectionNumber.Contains('|') ? group.Key.SectionNumber.Split('|')[0].Trim() : group.Key.SectionNumber;
                resultWorksheet.Cells[resultRow, 1].Value = displaySection;
                resultWorksheet.Cells[resultRow, 2].Value = basket1;
                resultWorksheet.Cells[resultRow, 3].Value = basket2;
                resultWorksheet.Cells[resultRow, 4].Value = basket3;
                resultWorksheet.Cells[resultRow, 5].Value = basket4;
                resultWorksheet.Cells[resultRow, 6].Value = group.Key.Coefficient;

                // Write wall IDs
                var wallIds = group.Select(x => x.WallId).Distinct().OrderBy(x => x).ToList();
                for (int i = 0; i < wallIds.Count; i++)
                {
                    resultWorksheet.Cells[resultRow, 7 + i].Value = wallIds[i];
                }

                resultRow++;
            }
        }

        private static int FindSourceRow(ExcelWorksheet worksheet, string sectionNumber, double coefficient)
        {
            int rowCount = worksheet.Dimension.Rows;
            for (int row = 2; row <= rowCount; row++)
            {
                if (worksheet.Cells[row, 1].Text.Trim() == sectionNumber &&
                    double.Parse(worksheet.Cells[row, 6].Text.Trim()) == coefficient)
                {
                    return row;
                }
            }
            return -1;
        }

    } 
}