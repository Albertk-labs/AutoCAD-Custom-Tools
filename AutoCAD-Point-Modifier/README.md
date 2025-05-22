# CombinedPointsModifier – AutoCAD Point Translation and Export Tool

This project is a plugin for AutoCAD, written in C#, that enhances point manipulation workflows. It enables users to interactively select and transform DBPoints in AutoCAD, project them onto various planes, and export the result to Excel.

---

## Features

- **Interactive point selection** with reference line
- **Point translation** based on perpendicular vector
- **Slope and height modification** for Z-coordinates
- **Projection** to XY, XZ, or YZ planes
- **Drawing translated points** in model space
- **Erase/Reset** options for managing data
- **Export to Excel** with EPPlus or Interop-based implementation

---

## Project Origin

This plugin was created with the help of **ChatGPT (OpenAI)** based on an idea and specification provided by the project author. The goal was to automate and simplify geometric workflows involving 3D points in design documents.

> “You don’t have to write all the code to be the author of the solution — systems thinking is the new programming.”

---

##  Repository Structure ```plaintext AutoCAD-Point-Modifier/ ├── CombinedPointsModifier.cs # Main logic and commands ├── DrawingMethods.cs # Drawing helpers ├── ExcelExportMethods.cs # Excel export ├── PointModificationMethods.cs # Point calculations ├── README.md # This file ```

---

## 💻 How to Use

1. Build the solution to produce a `.dll`.
2. Load the DLL into AutoCAD using the `NETLOAD` command.
3. Run the command: `ModifyPoints`
4. Follow the on-screen command-line prompts to:
   - Select points
   - Apply transformations
   - Export results

> Requires: AutoCAD with .NET support and `EPPlus` NuGet package (or MS Excel for Interop version)

---

## Technologies Used

- C# (.NET Framework)
- AutoCAD .NET API
- OfficeOpenXml (EPPlus)
- Microsoft.Office.Interop.Excel (optional)
- AutoCAD DBPoint & Geometry API



