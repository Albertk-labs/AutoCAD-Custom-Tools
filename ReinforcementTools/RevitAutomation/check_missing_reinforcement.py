# Importowanie potrzebnych bibliotek
import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# Pobranie dokumentu
doc = DocumentManager.Instance.CurrentDBDocument

# Wejście - lista wybranych ścian
selected_walls = UnwrapElement(IN[0])

# Faza, którą chcemy sprawdzić (New Construction)
new_construction_phase = None
for phase in FilteredElementCollector(doc).OfClass(Phase):
    if phase.Name == "New Construction":  # Zamień na polski odpowiednik jeśli potrzebne
        new_construction_phase = phase
        break

# Lista ścian spełniających oba warunki
filtered_walls = []

# Przetwarzanie każdej wybranej ściany
for wall in selected_walls:
    # Sprawdzenie fazy utworzenia
    if wall.CreatedPhaseId == new_construction_phase.Id:
        # Pobranie wartości parametru
        param = wall.LookupParameter("PRT_PL_TXT_RC_W_Wskaźnik_Zbrojenia_1")
        if param:
            param_value = param.AsString()
            # Sprawdzenie, czy parametr jest pusty lub ma wartość spacji
            if not param_value or param_value.strip() == "":
                filtered_walls.append(wall)

# Wyjście - lista ścian spełniających oba warunki
OUT = filtered_walls
