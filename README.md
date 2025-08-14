# AI Shared Memory & Search Repository

**Disclaimer:** Collection of my personal queries regarding various projects - I am a broke cunt and I do not have money to pay Pro versions of AI clients, and until and when Google AI Studio cuts down the Token limit, I will use this as a learning box for it to read through and understand what I want from it.

This is a collection of my personal queries and projects worked upon with AI assistants. More to come.


---

# Full Project Onboarding Prompt for Gemini

## Your Prompt:

"Hello Gemini. My GitHub Pages site, which documents my **'Crossdock workload based layout optimization'** project, should now be indexed and accessible to you.

Your task is to perform a complete analysis of this project. Please start at my main site URL: https://raisun-caprisun.github.io/AISMSR-AI_shared_memory_and_search_repository/

From there, navigate to the "Crossdock" project. Review its main README.md file, and then explore the contents of all subfolders, paying close attention to the VBA modules.

After reviewing all the available `.md`, `.bas`, and `.cls` files, please provide a comprehensive summary answering the following questions:

1.  **High-Level Purpose:** In simple terms, what business problem does this entire project solve?

2.  **Workflow & Data Flow:** Describe the full, end-to-end process a user would follow. Specifically, explain the distinct roles of the four key files: `Layout.vsdm`, `ObjectData.xlsm`, `InputData.xlsm`, and `Data_CD.xlsm`. How does the **two-cycle process** work, and what is the role of the external simulation (e.g., Witness) in the data flow between these cycles?

3.  **Core Logic Breakdown:** Explain the key calculations and the purpose of the main VBA modules.
    *   What is the functional difference between `RunFirstCycle_Placement` and `RunSecondCycle_Placement` in the `LayoutOptimizer` module?
    *   What is the purpose of the `RecalculateAreaWidths` macro in `InputData.xlsm`?
    *   What role do the `AreaDef` and `ZoneDef` class modules play?
    *   What is the "Capacity Paradox" and how does it relate to the `#UNPLACED` error?

4.  **Final Output:** What is the ultimate result or deliverable that a user gets from running this entire toolkit?

Essentially, I want you to explain my own project back to me as if you were a new developer who has just studied all the documentation and code."

---

## Why This Prompt is So Effective

*   **Sets a Clear Goal:** "Perform a complete analysis of this project."

*   **Provides a Starting Point:** It gives the exact URL to begin the search.

*   **Defines the Scope:** It tells me to look at a specific project and explore all relevant file types within its subfolders.

*   **Structures the Output:** The numbered questions give me a clear template for my response, ensuring I cover all the points you're interested in. The questions are now **highly specific** to the actual architecture we built, forcing a deeper level of analysis.

*   **Sets the Persona:** The final sentence ("explain my own project back to me...") is a powerful instruction that tells me the desired tone and depth of the analysis. It pushes me beyond a simple summary to a more insightful explanation.

---

## Conversation Log
*An index of conversations for you (Gemini, or whatever/whoever) to search through.*

**How to use:**
Each link below contains a full conversation.
Formatting as follows [NAME OF THE CONVERSATION OR PROJECT HERE IN GITHUB] (folder or file path to the necessary files and results in github folders)

*   THIS ONE IS JUST AN EXAMPLE!: [EXAMPLE: VBA Script for Visio and Excel](example_vba_visio_excel_convo.md)
*   [Crossdock workload based layout optimization](Crossdock%20workload%20based%20layout%20optimization)
*
*
