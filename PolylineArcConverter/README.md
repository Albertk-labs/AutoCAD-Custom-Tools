# Polyline Arc Converter

This AutoCAD plugin allows you to convert polylines into segmented arc representations using bulge values or true arc fitting.

## Features

- `S1`: Splits a selected polyline into arc-shaped segments using a user-defined number of divisions and arc direction.
- `S2`: Converts vertex-defined polyline shapes into arcs by fitting real arc geometry through every 3 points.
- `T`: Lets you adjust the arc-fitting tolerance dynamically from the command line.

## Usage

1. Load the plugin into AutoCAD (using NETLOAD or external application).
2. Run the command `PAC`.
3. Choose an option from the interactive menu.

## Notes

- Arc fitting is based on geometric calculations between three consecutive points.
- If arc fitting fails (e.g. nearly colinear points), the plugin will fall back to lines.
- The user can change the tolerance to make arc-fitting more or less strict.

---

*Part of the [AutoCAD-Custom-Tools](https://github.com/Albertk-labs/AutoCAD-Custom-Tools) collection.*
