import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)

# Get the current document
doc = DocumentManager.Instance.CurrentDBDocument

# Input: List of wall elements
walls = [UnwrapElement(wall) for wall in IN[0]]  # Assuming IN[0] is a list of selected wall elements

# View name where the tags will be placed
view_name = "BIM360_02 SSL PLATFORM LEVEL WALL ID"

# Find the view by name
view = None
collector = FilteredElementCollector(doc).OfClass(View)
for v in collector:
    if v.Name == view_name:
        view = v
        break

if view is None:
    OUT = "View '{}' not found.".format(view_name)
else:
    # Start a transaction
    TransactionManager.Instance.EnsureInTransaction(doc)

    # Find the specific tag family symbol
    tag_family_name = "PRT_PL_TXT_W_TAG_1"  # The tag family you want to use
    wall_tag_symbol = None
    family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_WallTags)
    
    for symbol in family_symbols:
        if symbol.Family.Name == tag_family_name:
            wall_tag_symbol = symbol
            break
    
    if wall_tag_symbol is None:
        OUT = "Tag family '{}' not found.".format(tag_family_name)
    else:
        # Activate the tag family symbol if it is not active
        if not wall_tag_symbol.IsActive:
            wall_tag_symbol.Activate()
            doc.Regenerate()

        # Iterate over each wall and place a tag at the midpoint
        for wall in walls:
            # Get the wall's location (curve)
            location_curve = wall.Location
            if isinstance(location_curve, LocationCurve):
                curve = location_curve.Curve
                # Find the midpoint of the wall
                midpoint = curve.Evaluate(0.5, True)

                # Check if a tag already exists for this wall in the view
                tags = FilteredElementCollector(doc, view.Id).OfClass(IndependentTag).WhereElementIsNotElementType().ToElements()
                tag_exists = False
                for tag in tags:
                    if tag.TaggedLocalElementId == wall.Id:
                        tag_exists = True
                        break

                # If no tag exists, place a new tag
                if not tag_exists:
                    wall_ref = Reference(wall)
                    IndependentTag.Create(doc, view.Id, wall_ref, False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, midpoint)

        # Complete the transaction
        TransactionManager.Instance.TransactionTaskDone()

        # Output success message
        OUT = "Tags placed at the midpoint of the selected walls in view '{}'.".format(view_name)
