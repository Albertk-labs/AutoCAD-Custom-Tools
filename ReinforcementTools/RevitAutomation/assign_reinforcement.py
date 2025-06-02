import clr
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.DB import *

import re  # To handle regex for transforming input mark values

# Inputs
dynamo_walls = IN[0]  # List of walls (elements) from Dynamo
mark_values = IN[1]  # List of mark names and values

# Prepare output
output = []

# Convert Dynamo elements to Revit elements
revit_walls = [element.InternalElement for element in dynamo_walls]

# Helper function to clean the mark value
# This function removes the number between the first letter and the dot, but leaves the part after the dot
def clean_mark_value(mark):
    if mark:  # Check if mark is not None
        # Search for a pattern like "N31.02" -> "N02"
        match = re.match(r'([A-Za-z]+)\d+\.(.*)', mark)
        if match:
            return match.group(1) + match.group(2)  # Return letter + part after the dot
        else:
            return mark  # If no match, return the original mark
    return None

# Start a transaction in Revit
doc = DocumentManager.Instance.CurrentDBDocument
TransactionManager.Instance.EnsureInTransaction(doc)

# Iterate over the walls
for wall in revit_walls:
    if wall is not None:
        wall_mark = wall.LookupParameter("Mark").AsString()  # Get the Mark parameter value

        if wall_mark is None:
            output.append("Brak Mark")  # Log message for no mark found
            continue
        
        # Clean the wall mark
        cleaned_wall_mark = clean_mark_value(wall_mark)
        print("Oczyszczony Mark:", cleaned_wall_mark)

        # Check if the wskaźnik zbrojenia parameter is already filled
        reinforcement_param = wall.LookupParameter("PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1")
        if reinforcement_param and reinforcement_param.AsString():
            output.append("Wskaźnik już wypełniony dla: " + wall_mark)
            continue

        # Iterate through mark_values and find match
        for i in range(0, len(mark_values), 2):
            cleaned_mark = clean_mark_value(mark_values[i])
            print("Oczyszczony Mark z IN[1]:", cleaned_mark)
            
            if cleaned_mark == cleaned_wall_mark:
                value_to_assign = round(float(mark_values[i + 1]), 3)  # Round to 3 decimal places
                formatted_value = "{:.3f} kg/m³".format(value_to_assign)  # Format the value with the unit
                print("Przypisywanie wartości:", formatted_value)

                # Set the PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1 parameter
                if reinforcement_param:
                    if reinforcement_param.StorageType == StorageType.String:  # Assuming the parameter is a string
                        reinforcement_param.Set(formatted_value)
                        output.append(True)
                    else:
                        output.append("Zły typ parametru")
                else:
                    output.append("Nie znaleziono parametru")
                break
        else:
            output.append("Nie znaleziono dopasowania dla: " + wall_mark)

TransactionManager.Instance.TransactionTaskDone()

# Output result
OUT = output
