# Crossdock Layout Optimization Toolkit

This repository contains a suite of VBA-powered tools designed to perform a data-driven optimization of a crossdock or warehouse layout. The system uses Microsoft Visio for visualization and Microsoft Excel as the computational engine to analyze and propose new, more efficient layouts based on real-world workload data and physical constraints.

The core philosophy of this toolkit is a two-cycle optimization process, allowing for an initial static optimization followed by a refined optimization based on dynamic feedback from an external simulation program (e.g., Witness).

## System Architecture

The toolkit is comprised of four key files that work together. Understanding the role of each is critical to using the system correctly.

*   **`Layout.vsdm` (Visio File)**
    *   The **visual source of truth** for the initial layout.
    *   Contains the macro to **export** the default layout data into the Excel files.
    *   Contains the macro to **import** the final, optimized layout for visualization.

*   **`ObjectData.xlsm` (Excel - The Digital Twin)**
    *   The **master database and main analytical engine** of the project.
    *   The "Layout" sheet contains a detailed, row-by-row representation of every object in the Visio drawing.
    *   Contains the core `LayoutOptimizer`, `MatrixAnalyzers`, and `AnalysisTools` macros.
    *   All major calculations and optimizations are performed within this file.

*   **`InputData.xlsm` (Excel - The Control Panel)**
    *   The primary **user input and control file**.
    *   This is where the user manually inputs or imports `Workload` and `Max_Buffer` data.
    *   Contains the `RecalculateAreaWidths` macro, which uses business logic to propose new rack sizes based on simulation feedback.

*   **`Data_CD.xlsm` (Excel - The Bridge to Simulation)**
    *   Acts as the clean data-interchange file between the Excel toolkit and the external simulation software (Witness).
    *   It receives the final, scaled distance matrix from `ObjectData.xlsm`.
    *   It serves as the source for initial `Workload` and `Max_Buffer` data.

---

## The A-Z Optimization Workflow

This is the step-by-step user guide for running a full, two-cycle optimization.

### Phase 1: Initial Data Export & Preparation

1.  **Visio: Export Initial Layout**
    *   **File:** `Layout.vsdm`
    *   **Action:** Run the `ExportLayoutuDoExcelu_Finalni_s_Dokumentaci` macro.
    *   **Result:** Populates `ObjectData.xlsm` and `InputData.xlsm` with the default layout data.

2.  **User Action: Prepare Initial Workload Data**
    *   **File:** `Data_CD.xlsm`
    *   **Action:** Manually populate the **"Workload"** sheet with initial `Area` (ID), `Workload`, and `Max_Buffer` data.

3.  **InputData: Import Initial Workload**
    *   **File:** `InputData.xlsm`
    *   **Action:** Run the `ImportWorkloadAndBufferData` macro.
    *   **Result:** Pulls the data from `Data_CD.xlsm` into the control file.

4.  **ObjectData: Sync with InputData**
    *   **File:** `ObjectData.xlsm`
    *   **Action:** Run the `UpdateFromInputData` macro.
    *   **Result:** Pulls the prepared data from `InputData.xlsm` into the master analysis file.

### Phase 2: First Optimization Cycle

5.  **ObjectData: Run First Optimization**
    *   **File:** `ObjectData.xlsm`
    *   **Action:** Run the `RunFirstCycle_Placement` macro.
    *   **Result:** Calculates the first optimized layout based on workload and **original** rack sizes. Populates the `New_...` coordinate columns.

6.  **ObjectData: Export for Simulation**
    *   **File:** `ObjectData.xlsm`
    *   **Action:** Run the `ExportMatrixForSimulation` macro.
    *   **Result:** Creates and exports the scaled distance matrix to `Data_CD.xlsm`, ready for simulation.

### Phase 3: Simulation & Second Cycle Preparation

7.  **User Action: Run Simulation & Update Data**
    *   **File:** `Data_CD.xlsm`
    *   **Action:** Use the "MaticeVzdalenosti" sheet as input for Witness. After the simulation, update the `Workload` and `Max_Buffer` columns in the "Workload" sheet with the simulation's feedback.

8.  **InputData: Import Simulation Feedback**
    *   **File:** `InputData.xlsm`
    *   **Action:** Run the `ImportWorkloadAndBufferData` macro again.
    *   **Result:** Pulls the new, simulation-driven data from `Data_CD.xlsm`.

9.  **InputData: Recalculate Rack Widths**
    *   **File:** `InputData.xlsm`
    *   **Action:** Run the `RecalculateAreaWidths` macro.
    *   **Result:** Uses the new `Max_Buffer` values to automatically assign new, tiered widths to the `New_Width` column.

10. **ObjectData: Sync for Final Optimization**
    *   **File:** `ObjectData.xlsm`
    *   **Action:** Run the `UpdateFromInputData` macro again.
    *   **Result:** Pulls the final, recalculated `New_Width` values into the master file.

### Phase 4: Final Optimization and Visualization

11. **ObjectData: Run Second Optimization**
    *   **File:** `ObjectData.xlsm`
    *   **Action:** Run the `RunSecondCycle_Placement` macro.
    *   **Result:** Re-runs the optimization algorithm using the **new, smarter rack sizes**, producing the definitive layout.

12. **ObjectData: Run Final Analysis**
    *   **File:** `ObjectData.xlsm`
    *   **Action:** Run the `RunFinalAnalysis` (or `CalculateAllLayoutCosts`) macro.
    *   **Result:** Generates the final `Cost_Calculation` report, quantifying the efficiency gains.

13. **Visio: Import Final Layout**
    *   **File:** `Layout.vsdm`
    *   **Action:** Run the `ImportLayout_KROK_1_NakreslitVse` macro.
    *   **Result:** Draws the final, fully optimized layout into Visio for review.

---

## Core Concepts & Logic

*   **The Two-Cycle Process:** This toolkit is built on a Plan-Do-Check-Act cycle.
    *   **Cycle 1 (Plan/Do):** Creates a statically optimized layout based on workload and distance.
    *   **Cycle 2 (Check/Act):** Uses dynamic feedback from a simulation to refine the layout by adjusting physical rack sizes, addressing real-world bottlenecks like congestion.

*   **Placement Algorithm:** The `LayoutOptimizer` uses a "greedy" algorithm. It sorts all movable "Areas" by workload (highest first) and places them one by one, always searching for the valid spot closest to the "INBOUND" area.

*   **The Capacity Paradox (`#UNPLACED` Errors):** It is possible for the optimizer to report that it cannot place a rack, even if the sum of all rack widths is less than the total width of the zones. This is not a bug; it is a correct calculation caused by:
    1.  **Space Fragmentation:** The total empty space is broken up across multiple, separate zones.
    2.  **The "Cost" of Gaps:** The algorithm must reserve mandatory safety gaps between racks, which consumes additional space not accounted for in the simple sum of widths.
    3.  **The Greedy Algorithm:** Placing high-priority items in the "best" spots can sometimes create smaller, unusable slivers of empty space.

## Setup & Prerequisites

1.  **Software:** Microsoft Office suite with Visio and Excel, including the VBA development environment.
2.  **File Paths:** All macros containing file paths (`Workbooks.Open`) use hardcoded constants at the top of the Sub. These must be updated to match the file locations on your system or SharePoint. For maximum reliability, it is recommended to use local OneDrive sync paths instead of direct SharePoint URLs.
3.  **Class Modules:** The `ObjectData.xlsm` file requires two **Class Modules** to be present: `AreaDef` and `ZoneDef`.
