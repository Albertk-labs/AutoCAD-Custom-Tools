RevitAutomation – Dynamo Python Scripts

This folder contains Dynamo Python scripts designed to automate the process of wall tagging and reinforcement data assignment in Autodesk Revit.

These scripts are used in a hybrid Revit ↔ AutoCAD workflow, integrated with custom C# tools that process exported DWG drawings and generate reinforcement data.

Available Scripts

1. place_tags.py

Places tags at the midpoint of selected wall elements in a specified Revit view.

Uses tag family: PRT_PL_TXT_W_TAG_1

Ensures no duplicate tags are placed

Works only on a given view, e.g. BIM360_02 SSL PLATFORM LEVEL WALL ID

Run this script first, before exporting DWG views to AutoCAD.

2. check_missing_reinforcement.py

Filters selected walls and returns only those that:

Belong to the phase New Construction

Do not have a value in parameter PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1

Run to identify which elements still need reinforcement data.

3. assign_reinforcement.py

Assigns reinforcement values to Revit walls based on external Excel data.

Matches walls by Mark (e.g. N31.02) → cleaned to N02

Assigns value (e.g. 42.125 kg/m³) to parameter PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1

Input format: paired list from Excel export: ["N31.02", 42.125, "N31.03", 35.7, ...]

Run this after S3 generates the Excel file.

Workflow Overview

This Revit-to-AutoCAD loop integrates Python automation with C# tools:

Step 1: Place Tags in Revit

Run place_tags.py to insert tag families

Export the Revit view to DWG to make tag positions readable in AutoCAD

Step 2: Analyze DWG in AutoCAD (C# Tools)

Use S1, S2 to verify and structure section & basket texts

Use S3 to associate wall IDs with reinforcement baskets, then generate the final Excel TME1.xlsx

S3 is the key step that prepares the Excel used by Revit

Step 3: Re-import Data into Revit

Use assign_reinforcement.py with Excel input from S3

Match IDs and assign reinforcement values to wall parameters

Optional: Data Completeness Check

Use check_missing_reinforcement.py to find unfilled elements

Folder Contents

File

Description

place_tags.py

Places tags in a specific view at wall midpoints

check_missing_reinforcement.py

Finds unfilled walls in "New Construction"

assign_reinforcement.py

Assigns reinforcement values from Excel

assign_reinforcement_nodes.png

Screenshot of Dynamo node structure

Requirements

Revit 2022+

Dynamo 2.x

Shared parameter: PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1 (String type)

Tag family: PRT_PL_TXT_W_TAG_1 loaded in the project

Related Repo Structure

These scripts are part of a larger toolset available in this repository:

ReinforcementTools/
├── AutoCAD-C#/          # C# tools: S1, S2, S3, E1
│   ├── S1.cs             # Text match tool (sections)
│   ├── S2.cs             # Validation tool
│   ├── S3.cs             # Main processor – creates TME1.xlsx
│   └── E1.cs             # Final Excel summary
└── RevitAutomation/       # You are here
     ├── assign_reinforcement.py
     ├── check_missing_reinforcement.py
     ├── place_tags.py
     └── README.md

Author

Created by Albert Kłoczewiak as part of a full-cycle automation workflow for structural coordination.
