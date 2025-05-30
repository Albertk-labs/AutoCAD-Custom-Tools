using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Globalization;
using System.IO;
using System.Threading;

public class AutoCADTextManager
{
    private static bool loggingEnabled = true;

    [CommandMethod("TextManager")]
    public static void TextManagerMenu()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        PromptKeywordOptions options = new PromptKeywordOptions("\nSelect operation:");
        options.Keywords.Add("S1");
        options.Keywords.Add("S2");
        options.Keywords.Add("Log");
        options.Keywords.Add("Exit");
        options.AllowNone = true;

        while (true)
        {
            PromptResult result = ed.GetKeywords(options);
            if (result.Status != PromptStatus.OK)
                return;

            switch (result.StringResult)
            {
                case "S1":
                    OrganizeTexts();
                    break;
                case "S2":
                    ProcessCoefficients();
                    break;
                case "Log":
                    ToggleLogging();
                    break;
                case "Exit":
                    return;
            }
        }
    }
        private static void OrganizeTexts()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        ed.WriteMessage("\nStarted command: OrganizeTexts");

        // Ask user if the created text should be flipped
        PromptKeywordOptions flipOptions = new PromptKeywordOptions("\nFlip new text?");
        flipOptions.Keywords.Add("Yes");
        flipOptions.Keywords.Add("No");
        flipOptions.AllowNone = true;
        PromptResult flipResult = ed.GetKeywords(flipOptions);
        bool flipText = flipResult.StringResult == "Yes";

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                var filter = new SelectionFilter(new[]
                {
                new TypedValue((int)DxfCode.Operator, "<OR"),
                new TypedValue((int)DxfCode.LayerName, "!!!SOL-Opisy klatek"),
                new TypedValue((int)DxfCode.LayerName, "!!!SOL-nr koszy"),
                new TypedValue((int)DxfCode.Operator, "OR>")
            });

                PromptSelectionResult selRes = ed.GetSelection(filter);
                if (selRes.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nNo objects selected. Operation cancelled.");
                    return;
                }

                List<TextInfo> cageLabels = new List<TextInfo>();
                List<TextInfo> basketNumbers = new List<TextInfo>();

                foreach (SelectedObject selObj in selRes.Value)
                {
                    Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent is MText || ent is DBText)
                    {
                        string content = ent is MText ? ((MText)ent).Contents : ((DBText)ent).TextString;
                        Point3d position = ent is MText ? ((MText)ent).Location : ((DBText)ent).Position;

                        if (ent.Layer == "!!!SOL-Opisy klatek")
                        {
                            string trimmed = content.Trim();
                            // Ignoruj tekst, jeśli to tylko liczby (np. 200)
                            if (Regex.IsMatch(trimmed, @"^\d+$")) continue;

                            // Weź tylko litery (np. G z "G", "G 200", itd.)
                            Match match = Regex.Match(trimmed, @"^[A-Za-z]+");
                            if (match.Success)
                                cageLabels.Add(new TextInfo(ent, match.Value, position));
                        }

                        else if (ent.Layer == "!!!SOL-nr koszy")
                        {
                            basketNumbers.Add(new TextInfo(ent, content, position));
                        }
                    }
                }

                int createdCount = 0;
                foreach (var basket in basketNumbers)
                {
                    var nearestLabel = FindNearestText(basket, cageLabels);
                    if (nearestLabel != null)
                    {
                        string newText = $"{nearestLabel.Content}{basket.Content.Trim()}";
                        CreateNewText(tr, db, newText, basket.Position, basket.Entity, flipText);
                        createdCount++;
                    }
                }

                ed.WriteMessage($"\nCreated {createdCount} new texts.");
                tr.Commit();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                tr.Abort();
            }
        }
    }
    private class TextInfo
    {
        public Entity Entity { get; }
        public string Content { get; }
        public Point3d Position { get; }

        public TextInfo(Entity entity, string content, Point3d position)
        {
            Entity = entity;
            Content = content;
            Position = position;
        }
    }
    private static void CreateNewText(Transaction tr, Database db, string text, Point3d position, Entity reference, bool flip)
    {
        // Upewnij się, że warstwa istnieje
        const string targetLayerName = "PRT_MOD_TXT";
        EnsureLayerExists(tr, db, targetLayerName, 3); // zielony kolor

        DBText newText = new DBText
        {
            Position = position,
            TextString = text,
            Layer = targetLayerName,
            Color = Color.FromColorIndex(ColorMethod.ByAci, 3) // zielony
        };

        if (reference is DBText dbRef)
        {
            newText.Height = dbRef.Height;
            newText.Rotation = dbRef.Rotation + (flip ? Math.PI : 0);
            newText.WidthFactor = dbRef.WidthFactor;
            newText.TextStyleId = dbRef.TextStyleId;
            newText.HorizontalMode = dbRef.HorizontalMode;
            newText.VerticalMode = dbRef.VerticalMode;

            if (dbRef.HorizontalMode != TextHorizontalMode.TextLeft || dbRef.VerticalMode != TextVerticalMode.TextBase)
                newText.AlignmentPoint = dbRef.AlignmentPoint;
        }
        else if (reference is MText mRef)
        {
            newText.Height = mRef.TextHeight;
            newText.Rotation = mRef.Rotation + (flip ? Math.PI : 0);
            newText.WidthFactor = mRef.Width;
            newText.TextStyleId = mRef.TextStyleId;
            newText.HorizontalMode = TextHorizontalMode.TextLeft;
            newText.VerticalMode = TextVerticalMode.TextBase;
            newText.Position = mRef.Location;
        }

        var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        btr.AppendEntity(newText);
        tr.AddNewlyCreatedDBObject(newText, true);
    }

    private static void ProcessCoefficients()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        string excelPath = @"C:\Projekty\dane.xlsx";
        LogMessage($"Reading Excel data from: {excelPath}");

        var excelData = ReadExcelData(excelPath);
        LogMessage($"Read {excelData.Count} rows from Excel");

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                var sectionLabels = GetEntitiesFromLayer(tr, db, "!!!SOL-Opisy sekcji");
                var basketTexts = GetEntitiesFromLayer(tr, db, "PRT_MOD_TXT");

                LogMessage($"Found {sectionLabels.Count} section labels and {basketTexts.Count} baskets");

                var assignment = AssignBasketsToSections(sectionLabels, basketTexts);

                foreach (var pair in assignment)
                {
                    string sectionName = pair.Key is MText m ? m.Contents : (pair.Key as DBText)?.TextString;
                    if (string.IsNullOrWhiteSpace(sectionName)) continue;

                    string cleaned1 = CleanSectionNameStage1(sectionName);
                    string cleaned2 = CleanSectionNameStage2(cleaned1);

                    var match1 = MatchExcelData(cleaned1, pair.Value, excelData);
                    var match2 = MatchExcelData(cleaned2, pair.Value, excelData);

                    var bestMatch = ChooseBestMatch(match1, match2, pair.Key, pair.Value);

                    if (bestMatch.Any())
                    {
                        double total = bestMatch.Sum(x => x.coefficient);
                        string coefficient = total.ToString("F2", CultureInfo.InvariantCulture);

                        InsertCoefficient(tr, db, coefficient, pair.Key, GetEntityPosition(pair.Key));
                        LogAssignmentDetails(pair.Key, pair.Value, sectionName, coefficient);
                        continue;
                    }

                    // --- Faza 3: dopasowanie po koszach ---
                    var nearbyBasketNames = pair.Value
                        .Select(b => b is DBText txt ? txt.TextString : (b as MText)?.Contents)
                        .Where(n => !string.IsNullOrWhiteSpace(n))
                        .ToList();

                    var altMatches = excelData
                        .Where(row => row.baskets.Any(b => nearbyBasketNames.Contains(b, StringComparer.OrdinalIgnoreCase)))
                        .ToList();

                    if (altMatches.Any())
                    {
                        double total = altMatches.Sum(x => x.coefficient);
                        string coefficient = total.ToString("F2", CultureInfo.InvariantCulture);

                        InsertCoefficient(tr, db, coefficient, pair.Key, GetEntityPosition(pair.Key));
                        LogAssignmentDetails(pair.Key, pair.Value, sectionName, coefficient + " (Phase 3)");
                    }
                    else
                    {
                        LogMessage($"[S2] Section {sectionName} → no matching baskets found (Phase 3)");
                    }
                }

                tr.Commit();
                LogMessage("Transaction committed.");
            }
            catch (System.Exception ex)
            {
                LogMessage($"Error: {ex.Message}\n{ex.StackTrace}");
                tr.Abort();
            }
        }
    }


    private static List<(string sectionName, List<string> baskets, double coefficient)> ReadExcelData(string path)
    {
        var results = new List<(string, List<string>, double)>();

        using (var workbook = new XLWorkbook(path))
        {
            var sheet = workbook.Worksheet(1);
            foreach (var row in sheet.RowsUsed().Skip(1)) // skip header
            {
                string section = row.Cell(1).GetString().Trim();
                if (string.IsNullOrWhiteSpace(section)) continue;

                var baskets = new List<string>();
                for (int i = 2; i <= 5; i++)
                {
                    string basket = row.Cell(i).GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(basket))
                        baskets.Add(basket);
                }

                string coeffRaw = row.Cell(6).GetString().Replace(',', '.');
                if (double.TryParse(coeffRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out double coeff))
                {
                    results.Add((section, baskets, coeff));
                }
            }
        }

        return results;
    }
    private static Dictionary<Entity, List<Entity>> AssignBasketsToSections(List<Entity> sections, List<Entity> baskets)
    {
        var mapping = new Dictionary<Entity, List<Entity>>();

        foreach (var basket in baskets)
        {
            var closest = FindClosestEntity(basket, sections);
            if (closest != null)
            {
                if (!mapping.ContainsKey(closest))
                    mapping[closest] = new List<Entity>();

                mapping[closest].Add(basket);

                double distance = GetEntityPosition(closest).DistanceTo(GetEntityPosition(basket));
                LogMessage($"[S2] Basket at {FormatPoint(GetEntityPosition(basket))} assigned to Section at {FormatPoint(GetEntityPosition(closest))} (Distance: {distance:F2} m)");
            }
            else
            {
                LogMessage($"[S2] Basket at {FormatPoint(GetEntityPosition(basket))} has no matching section.");
            }
        }

        return mapping;
    }

    private static Entity FindClosestEntity(Entity reference, List<Entity> candidates)
    {
        Point3d refPos = GetEntityPosition(reference);
        var closest = candidates
            .Select(e => new { Entity = e, Distance = GetEntityPosition(e).DistanceTo(refPos) })
            .OrderBy(x => x.Distance)
            .FirstOrDefault();

        if (closest != null)
        {
            LogMessage($"[S2] Closest section to basket at {FormatPoint(refPos)} is at {FormatPoint(GetEntityPosition(closest.Entity))} (Distance: {closest.Distance:F2} m)");
            return closest.Entity;
        }

        return null;
    }

    private static Point3d GetEntityPosition(Entity ent)
    {
        return ent is MText m ? m.Location : (ent as DBText)?.Position ?? Point3d.Origin;
    }
    private static string CleanSectionNameStage1(string name)
    {
        return Regex.Replace(name, @"^([A-Z]).*?\.(.*)$", "$1$2");
    }

    private static string CleanSectionNameStage2(string name)
    {
        return Regex.Replace(name, @"^([A-Z])0+(\d+)$", "$1$2");
    }
    private static List<(string sectionName, List<string> baskets, double coefficient)> MatchExcelData(
        string cleanedSection,
        List<Entity> assignedBaskets,
        List<(string sectionName, List<string> baskets, double coefficient)> excelData)
    {
        var assignedNames = assignedBaskets
            .Select(b => b is DBText db ? db.TextString : (b as MText)?.Contents)
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToList();

        return excelData
            .Where(d =>
                CleanSectionNameStage2(CleanSectionNameStage1(d.sectionName)) == cleanedSection &&
                d.baskets.Any(b => assignedNames.Contains(b, StringComparer.OrdinalIgnoreCase)))
            .ToList();
    }
    private static List<(string sectionName, List<string> baskets, double coefficient)> ChooseBestMatch(
        List<(string sectionName, List<string> baskets, double coefficient)> match1,
        List<(string sectionName, List<string> baskets, double coefficient)> match2,
        Entity section,
        List<Entity> assignedBaskets)
    {
        if (!match1.Any() && !match2.Any())
            return new List<(string sectionName, List<string> baskets, double coefficient)>();


        double dist1 = CalculateAverageDistance(section, assignedBaskets, match1.SelectMany(x => x.baskets).ToList());
        double dist2 = CalculateAverageDistance(section, assignedBaskets, match2.SelectMany(x => x.baskets).ToList());

        return dist1 <= dist2 ? match1 : match2;
    }
    private static double CalculateAverageDistance(Entity section, List<Entity> allBaskets, List<string> targetBasketNames)
    {
        Point3d sectionPos = GetEntityPosition(section);

        var matched = allBaskets
            .Where(b =>
            {
                string name = b is DBText db ? db.TextString : (b as MText)?.Contents;
                return targetBasketNames.Contains(name, StringComparer.OrdinalIgnoreCase);
            })
            .ToList();

        if (!matched.Any()) return double.MaxValue;

        return matched
            .Average(b => GetEntityPosition(b).DistanceTo(sectionPos));
    }
    private static void InsertCoefficient(Transaction tr, Database db, string value, Entity referenceEntity, Point3d position)
    {
        // Upewnij się, że warstwa istnieje
        EnsureLayerExists(tr, db, "PRT_WSP_ZBR", 6);

        DBText coeffText = new DBText
        {
            TextString = value,
            Layer = "PRT_WSP_ZBR",
            Position = position,
            HorizontalMode = TextHorizontalMode.TextCenter,
            VerticalMode = TextVerticalMode.TextVerticalMid,
            AlignmentPoint = position
        };

        // Nadanie rozmiaru i stylu
        if (referenceEntity is MText mText)
        {
            coeffText.Height = mText.TextHeight > 0 ? mText.TextHeight : 2.5;
            coeffText.Rotation = mText.Rotation;
            coeffText.TextStyleId = mText.TextStyleId;
        }
        else if (referenceEntity is DBText dbText)
        {
            coeffText.Height = dbText.Height > 0 ? dbText.Height : 2.5;
            coeffText.Rotation = dbText.Rotation;
            coeffText.WidthFactor = dbText.WidthFactor;
            coeffText.TextStyleId = dbText.TextStyleId;
        }
        else
        {
            coeffText.Height = 2.5; // domyślna wartość, jeśli brak wzorca
        }

        // Dodanie do model space
        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        btr.AppendEntity(coeffText);
        tr.AddNewlyCreatedDBObject(coeffText, true);
    }

    private static void LogMessage(string message)
    {
        if (!loggingEnabled) return;

        try
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TextManagerLog.txt");
            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {message}");
            }
        }
        catch { }
    }
    private static TextInfo FindNearestText(TextInfo reference, List<TextInfo> candidates)
    {
        return candidates
            .OrderBy(t => t.Position.DistanceTo(reference.Position))
            .FirstOrDefault();
    }
    private static List<Entity> GetEntitiesFromLayer(Transaction tr, Database db, string layerName)
    {
        var result = new List<Entity>();
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        foreach (ObjectId id in btr)
        {
            if (tr.GetObject(id, OpenMode.ForRead) is Entity ent && ent.Layer == layerName)
            {
                result.Add(ent);
            }
        }

        return result;
    }
   
    private static void ToggleLogging()
    {
        loggingEnabled = !loggingEnabled;

        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        ed.WriteMessage($"\nLogging is now {(loggingEnabled ? "ENABLED" : "DISABLED")}");
    }
    private static void EnsureLayerExists(Transaction tr, Database db, string layerName, short colorIndex)
    {
        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

        if (!lt.Has(layerName))
        {
            lt.UpgradeOpen();
            LayerTableRecord newLayer = new LayerTableRecord
            {
                Name = layerName,
                Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex)
            };

            lt.Add(newLayer);
            tr.AddNewlyCreatedDBObject(newLayer, true);
        }
    }

    private static string FormatPoint(Point3d pt)
    {
        return $"X:{pt.X:F2}, Y:{pt.Y:F2}";
    }
    private static void LogAssignmentDetails(Entity sectionEntity, List<Entity> baskets, string sectionName, string coefficient)
    {
        if (baskets == null || baskets.Count == 0)
        {
            LogMessage($"[S2] Section {sectionName} → no matching baskets found");
            return;
        }

        Point3d sectionPos = GetEntityPosition(sectionEntity);

        var parts = baskets.Select(b =>
        {
            string name = b is DBText db ? db.TextString : (b as MText)?.Contents;
            double dist = GetEntityPosition(b).DistanceTo(sectionPos);
            return $"{name} ({dist:F2}m)";
        });

        string joined = string.Join(", ", parts);
        LogMessage($"[S2] Section {sectionName} → matched baskets: {joined} → Coefficient: {coefficient}");
    }


}
