using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Globalization;
using System.Text;
public class AutoCADTextManager
{
    // ---------------------------
    // GŁÓWNA METODA E1() WERSJA Z LOGIKĄ S2: CZYTANIE Z EXCELA I EKSPORT DO TME1
    // ---------------------------

    [CommandMethod("E1")]
    public void E1()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        string excelPath = @"C:\Projekty\dane.xlsx";
        string exportPath = @"C:\Projekty\TME1.xlsx";
        double searchRadius = 10.0;

        try
        {
            if (!Directory.Exists(@"C:\Projekty"))
                Directory.CreateDirectory(@"C:\Projekty");

            var excelData = ReadExcelData(excelPath);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var sections = GetEntitiesFromLayer(tr, db, "!!!SOL-Opisy sekcji");
                var baskets = GetEntitiesFromLayer(tr, db, "PRT_MOD_TXT");

                var results = new List<(string sectionName, List<string> baskets, string coefficient)>();

                foreach (var section in sections)
                {
                    string sectionText = section is DBText t ? t.TextString : (section as MText)?.Contents;
                    if (string.IsNullOrWhiteSpace(sectionText)) continue;

                    Point3d sectionPos = GetEntityPosition(section);
                    string keyEnding = GetSectionKeyEnding(sectionText);

                    LogMessage($"\n--- Przetwarzanie sekcji: {sectionText} ---");
                    LogMessage($"Pozycja sekcji: X={sectionPos.X}, Y={sectionPos.Y}");
                    LogMessage($"Końcówka klucza sekcji: {keyEnding}");

                    var nearbyBaskets = baskets
                        .Select(b => new
                        {
                            Text = b is DBText t1 ? t1.TextString : (b as MText)?.Contents,
                            Distance = GetDistanceBetween(sectionPos, GetEntityPosition(b))
                        })
                        .Where(b => !string.IsNullOrWhiteSpace(b.Text) && b.Distance <= searchRadius)
                        .OrderBy(b => b.Distance)
                        .ToList();

                    LogMessage($"Znaleziono {nearbyBaskets.Count} koszy w promieniu {searchRadius} m:");
                    foreach (var b in nearbyBaskets)
                        LogMessage($" - {b.Text} (odległość: {b.Distance:0.##} m)");

                    var finalBaskets = new HashSet<string>();
                    var usedRows = new HashSet<string>();
                    double totalCoeff = 0.0;

                    foreach (var basket in nearbyBaskets)
                    {
                        LogMessage($"[DOPASOWANIE] Próbuję dopasować kosz: {basket.Text}");

                        var matchingRows = excelData
                            .Where(row => row.baskets.Any(b => string.Equals(b.Trim(), basket.Text.Trim(), StringComparison.OrdinalIgnoreCase)))
                            .Where(row => GetSectionKeyEnding(row.sectionName) == keyEnding)
                            .ToList();

                        LogMessage($"[DOPASOWANIE] Dopasowano {matchingRows.Count} wierszy dla kosza {basket.Text} i końcówki {keyEnding}");

                        foreach (var row in matchingRows)
                        {
                            string uniqueKey = row.sectionName + "|" + string.Join(",", row.baskets).ToLowerInvariant().Trim();
                            if (usedRows.Contains(uniqueKey)) continue;
                            usedRows.Add(uniqueKey);

                            foreach (var b in row.baskets)
                            {
                                if (!finalBaskets.Contains(b) && finalBaskets.Count < 4)
                                    finalBaskets.Add(b);
                            }

                            totalCoeff += row.coefficient;

                            LogMessage($"Dopasowano wiersz z Excela: {row.sectionName} → kosze: {string.Join(", ", row.baskets)} | współczynnik: {row.coefficient}");

                            if (finalBaskets.Count >= 4) break;
                        }

                        if (finalBaskets.Count >= 4) break;
                    }

                    if (finalBaskets.Count > 0)
                    {
                        LogMessage($"Finalne kosze: {string.Join(", ", finalBaskets)}");
                        LogMessage($"Suma współczynników: {totalCoeff:0.##}");

                        results.Add((sectionText, finalBaskets.ToList(), totalCoeff.ToString("0.##", CultureInfo.InvariantCulture)));
                    }
                    else
                    {
                        LogMessage("Nie znaleziono dopasowań – współczynnik pusty.");
                        results.Add((sectionText, new List<string>(), ""));
                    }

                    LogMessage($"Zakończono analizę sekcji: {sectionText}\n");
                }

                ExportToExcel(results, exportPath);
                ed.WriteMessage($"\nZapisano plik: {exportPath}");
                tr.Commit();
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nBłąd: {ex.Message}");
            LogMessage($"Error: {ex.Message}\n{ex.StackTrace}");
        }
    }


    private List<(string sectionName, List<string> baskets, double coefficient)> ReadExcelData(string path)
    {
        var results = new List<(string, List<string>, double)>();

        using (var workbook = new XLWorkbook(path))
        {
            var sheet = workbook.Worksheet(1);
            foreach (var row in sheet.RowsUsed().Skip(1))
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
                    LogMessage($"[EXCEL] Sekcja: {section} | Kosze: {string.Join(", ", baskets)} | Współczynnik: {coeff}");
                    results.Add((section, baskets, coeff));
                }
                else
                {
                    LogMessage($"[EXCEL] Nieparsowalny współczynnik dla sekcji {section}: {coeffRaw}");
                }
            }
        }

        return results;
    }

    private List<Entity> GetEntitiesFromLayer(Transaction tr, Database db, string layerName)
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
    private Point3d GetEntityPosition(Entity ent)
    {
        return ent is MText m ? m.Location : (ent as DBText)?.Position ?? Point3d.Origin;
    }
    private string GetSectionKeyEnding(string sectionName)
    {
        int dotIndex = sectionName.IndexOf('.');
        string rawEnding = "";

        if (dotIndex >= 0)
            rawEnding = sectionName.Substring(dotIndex + 1).Trim(); // wszystko po kropce
        else
            rawEnding = new string(sectionName.Where(char.IsDigit).ToArray()); // gdy nie ma kropki

        // usuń wiodące zera, ale nie zostaw pustego
        if (int.TryParse(rawEnding, out int numericEnding))
            return numericEnding.ToString();

        return rawEnding; // fallback, jeśli nie liczba
    }




    private double GetDistanceBetween(Point3d p1, Point3d p2)
    {
        return p1.DistanceTo(p2);
    }
    private void ExportToExcel(List<(string sectionName, List<string> baskets, string coefficient)> data, string filePath)
    {
        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.Worksheets.Add("TME1");

            ws.Cell(1, 1).Value = "Section";
            ws.Cell(1, 2).Value = "Basket 1";
            ws.Cell(1, 3).Value = "Basket 2";
            ws.Cell(1, 4).Value = "Basket 3";
            ws.Cell(1, 5).Value = "Basket 4";
            ws.Cell(1, 6).Value = "Coefficient";

            int row = 2;
            foreach (var (section, baskets, coeff) in data)
            {
                ws.Cell(row, 1).Value = section;
                for (int i = 0; i < baskets.Count && i < 4; i++)
                    ws.Cell(row, i + 2).Value = baskets[i];

                ws.Cell(row, 6).Value = coeff;

                if (string.IsNullOrWhiteSpace(coeff))
                    ws.Cell(row, 6).Style.Fill.BackgroundColor = XLColor.LightPink;

                row++;
            }

            workbook.SaveAs(filePath);
        }

        LogMessage($"Saved Excel to: {filePath}");
    }

    private void LogMessage(string message)
    {
        try
        {
            string path = @"C:\Projekty\TextManagerLog.txt";
            File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {message}{Environment.NewLine}");
        }
        catch { }
    }

}
