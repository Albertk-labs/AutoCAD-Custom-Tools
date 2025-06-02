# Revit Automation Scripts (Dynamo Python)

These scripts extend the functionality of the AutoCAD-based reinforcement tools by allowing automatic assignment and tagging of reinforcement data inside Autodesk Revit using Dynamo and Python.

## Scripts

### 1. assign_reinforcement.py
- Assigns reinforcement factor (`PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1`) to wall elements.
- Matches wall `Mark` parameter with data from Excel.

### 2. place_tags.py
- Automatically places wall tags in a specified view.
- Tags are placed at the midpoint of wall geometry.

## Workflow Connection

These scripts should be used **after** running AutoCAD tools (`E1`, `S1`, etc.), which export Excel data like `TME1.xlsx`.

## Requirements

- Revit 2022+
- Dynamo 2.x
- Shared parameter loaded in the project
