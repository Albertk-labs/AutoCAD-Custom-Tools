# AutoCAD Text and Reinforcement Manager

A set of tools for AutoCAD developed in C# to automate the extraction, organization, and verification of text-based reinforcement data using Excel input/output. These tools streamline engineering workflows in civil and structural design processes.

## Contents

### 1. `AutoCADTextManager_S1_S2.cs`

Main command: `TextManager`

Provides two subcommands:

- `S1`: Automatically links text from different layers (e.g., `!!!SOL-Opisy klatek` and `!!!SOL-nr koszy`) and generates combined text entities on the `PRT_MOD_TXT` layer.
- `S2`: Reads reinforcement coefficients from an Excel file (`dane.xlsx`) and assigns them to nearby section labels based on basket texts. Displays the result in AutoCAD as temporary visual markers — useful for verifying the logic later used by `E1`.

> Note: `S2` does not generate Excel output — it's intended as a **visual testing tool**.

---

### 2. `AutoCADExport_E1.cs`

Main command: `E1`

- Uses the same Excel data source (`dane.xlsx`) and extracts basket-to-section relationships based on spatial proximity and matching suffixes in section IDs (e.g., `P2.17` → `.17`).
- Summarizes matched reinforcement coefficients and exports them to `TME1.xlsx` in a structured format.
- Intended for **production use** as the primary data extraction/export tool for AutoCAD-to-Excel workflows.

---

### 3. `AutoCADWallAnalysis_S3.cs`

Main command: `TM`

- Loads processed Excel file (`TME1.xlsx`) and correlates section data to **wall identifiers** in AutoCAD (from the `A-WALL-____-IDEN` layer).
- Automatically detects intersections between reinforcement text labels and wall regions.
- Outputs structured Excel file `RevitDane.xlsx`, associating wall IDs with their corresponding section data.
- Can optionally draw helper circles in AutoCAD to visualize detection areas.

---

## Requirements

- AutoCAD with .NET API (tested with AutoCAD 2022+)
- .NET libraries: `ClosedXML`, `EPPlus`
- Folder path `C:\Projekty` used for logs and data files:
  - `dane.xlsx`: source data
  - `TME1.xlsx`: generated data (E1)
  - `RevitDane.xlsx`: final wall-to-section mapping (S3)
  - `TextManagerLog.txt`: optional logging output

---

## Purpose

These scripts were developed as part of my engineering workflow to reduce manual errors and streamline repetitive tasks in large-scale construction documentation projects.

- Used in **real-world structural projects** involving concrete reinforcement.
- Designed to be **modular** and adaptable to different naming conventions.
- Demonstrates ability to integrate AutoCAD API with Excel and handle spatial data analysis.

---

## License

This code is shared for demonstration and portfolio purposes. Please contact me if you are interested in adapting it for your own production use.
