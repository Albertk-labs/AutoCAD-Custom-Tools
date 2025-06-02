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

## Script Overview – assign_reinforcement.py

This script:
- Reads a list of wall elements selected in Dynamo.
- Matches each wall's "Mark" parameter (e.g. `N31.02`) with cleaned-up keys from Excel.
- Assigns the reinforcement value to the parameter `PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1`.

### Input format (from Excel):
- Paired list: `[ "N31.02", 45.25, "K12.03", 38.75, ... ]`
- Matches are done based on cleaned form like `N02`, `K03`, etc.

---

## Node Graph in Dynamo

This graph illustrates the logic of parameter assignment from Excel:

![Dynamo Node Graph](assign_reinforcement_nodes.png)

## Requirements

- Revit 2022+
- Dynamo 2.x
- Shared parameter loaded in the project
