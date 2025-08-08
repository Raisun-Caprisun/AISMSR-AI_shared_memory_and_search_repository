# Project: Crossdeck Workload Based Layout Optimization

## 1. Overview

This project is a tool designed to find an optimal physical layout based on workload data. It uses VBA within Excel to perform the core calculations and VBA within Visio to import and visualize the layout and results.

The primary function is to calculate layout "costs" using different distance metrics (Euclidean vs. Manhattan) and help users make data-driven decisions on facility design.

---

## 2. Core Components & File Structure

This project is divided into two main parts: Excel for calculation and Visio for visualization.

### Excel Components (`/Excel/`)

This folder contains the main calculation engine, data files, and documentation.

**Key Files:**
*   `visio-excel-objectdata-and-macros.xlsm`: The main Excel workbook. This is where the user inputs data and runs the primary macros.
*   **Calculation Modules (.bas):**
    *   `LayoutCostCalculator.bas`: The central module that likely calculates the final cost of a layout.
    *   `MatrixDefaultEuclidian.bas` & `MatrixDefaultManhattan.bas`: Modules to calculate straight-line vs. grid-based distance matrices.
    *   `MatrixWorkloadEuclidian.bas` & `MatrixWorkloadManhattan.bas`: Modules that combine workload data with the distance matrices.
*   **Class Modules (.cls):**
    *   `RectangleDef.cls` & `ZoneDef.cls`: Define custom data structures for "Rectangles" and "Zones", making the code more robust.
*   **Documentation (.txt):**
    *   `DOCUMENTATION_CZECH.txt` & `DOCUMENTATION_ENGLISH.txt`: User manuals in Czech and English.

### Visio Components (`/Visio/`)

This folder contains the Visio template and macros for importing and exporting data to visualize the layout.

**Key Files:**
*   `PRG layout.vsdm`: The macro-enabled Visio drawing file used as the template or final output.
*   **Import/Export Modules (.bas):**
    *   `ImportDefault.bas` & `ExportDefault.bas`: General modules for moving data between Visio and another source (likely Excel).
    *   `ImportWorkloadNewSheet.bas` & `ImportWorkloadSheetPowerBI.bas`: Specialized import functions for handling workload data.

---

## 3. How to Use (High-Level)

1.  Open the `visio-excel-objectdata-and-macros.xlsm` file.
2.  Input layout coordinates and workload data into the specified sheets.
3.  Run the calculation macros from within Excel to generate cost matrices.
4.  Open the `PRG layout.vsdm` file in Visio.
5.  Use the Visio macros to import the layout and cost data for visualization.
