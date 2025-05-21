# Surface and Curve-Plane Intersection Tools

This module contains custom AutoCAD commands written in C# using the .NET API. It provides tools for generating simple 3D surfaces and detecting intersection points between curves and planar surfaces.

---

## 🛠 Commands

### `DRS` – Draw Surface
Creates a vertical surface (region and PlaneSurface) based on a user-selected line in the XY plane. The surface height is defined by the user via prompt.

**Functionality:**
- Prompts user to select a line and enter a height
- Draws vertical boundary lines
- Generates a `Region` and converts it to a `PlaneSurface`
- Appends geometry to the model space

---

### `CPI` – Curve-Plane Intersection
Calculates intersection points between any curve and a planar surface.

**Functionality:**
- Prompts user to select a curve and a `PlaneSurface`
- Projects the curve onto the plane
- Detects intersection points between original curve and projection
- Adds `DBPoint` markers at intersection points
- Highlights points in yellow and outputs debug info

---

## Technologies Used

- AutoCAD .NET API (C#)
- `PlaneSurface`, `Region`, `Curve`, `Point3d`
- Extension methods:
  - `IsOn(this Curve, Point3d)`
  - `IntersectWith(this Plane, Curve)`

---

##  Usage

1. Compile the code into a `.dll` using Visual Studio
2. Load it into AutoCAD via the `NETLOAD` command
3. Use the following commands:
   - `DRS` → draw a vertical surface from a line
   - `CPI` → find intersections between a curve and planar surface

---

## Author

**Albert Kłoczewiak**  
