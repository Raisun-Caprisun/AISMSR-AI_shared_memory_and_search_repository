'=======================================================================================
'
'   MASTER AUTOMATION SCRIPT for Crossdock Layout Optimization
'   Version: 3.1 (Final - Automated Visio Save)
'
'=======================================================================================
Option Explicit

' --- This block ensures the script runs in the command-line host (cscript) ---
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
If InStr(UCase(WScript.FullName), "CSCRIPT.EXE") = 0 Then
    WshShell.Run "cscript.exe """ & WScript.ScriptFullName & """", 1, False
    WScript.Quit
End If

' --- Announce the start and get the project folder from the user ---
MsgBox "Welcome to the Crossdock Optimization Toolkit." & vbCrLf & vbCrLf & "You will now be asked to select the folder containing your project files (Visio, ObjectData, etc.).", vbInformation, "Crossdock Optimization."

Dim shell, folder, scriptPath
Set shell = CreateObject("Shell.Application")
Set folder = shell.BrowseForFolder(0, "Please select the project folder:", 0)

If (Not folder Is Nothing) Then
    scriptPath = folder.Self.Path & "\"
Else
    MsgBox "No folder was selected. The script will now exit.", vbExclamation, "Operation Cancelled."
    WScript.Quit
End If
Set folder = Nothing
Set shell = Nothing

MsgBox "Starting the full optimization process for the folder:" & vbCrLf & scriptPath & vbCrLf & vbCrLf & "A command window will now open to show progress. Please do not close it.", vbInformation, "Crossdock Optimization."

' --- Get handles to required objects ---
Dim fso, excelApp, visioApp
Dim visioFilePath, objectDataFilePath, inputDataFilePath
Dim visioDoc, excelWb

Set fso = CreateObject("Scripting.FileSystemObject")
visioFilePath = scriptPath & "Layout.vsdm"
objectDataFilePath = scriptPath & "ObjectData.xlsm"
inputDataFilePath = scriptPath & "InputData.xlsm"

On Error Resume Next

'=======================================================================================
'   PHASE 1: INITIAL DATA EXPORT & FIRST OPTIMIZATION
'=======================================================================================
WScript.Echo "--- Crossdock Optimization Tool ---" & vbCrLf & vbCrLf & "--- Enjoy the silence edition ---" & vbCrLf & vbCrLf & "--- COPYRIGHT 2025: Jakub Andar & Roman Korpos, all rights denied. :^) ---"& vbCrLf & vbCrLf
WScript.Echo "--- PHASE 1: STARTING ---" & vbCrLf & vbCrLf 
WScript.Echo "NOTE: Step 1.5-2 (Manual Data Prep in Data_CD.xlsm) is assumed to be complete. Always make sure for now that Data_CD is populated!" & vbCrLf & vbCrLf & "This message will be gone once Witness is also included in the automation." & vbCrLf

' --- Step 1: Export from Visio ---
WScript.Echo "Step 1/12: Exporting default layout from Visio... You might want to grab a coffee, these twelve steps will take a minute or two."
Set visioApp = CreateObject("Visio.Application")
visioApp.Visible = False
Set visioDoc = visioApp.Documents.Open(visioFilePath)
visioDoc.ExecuteLine "ExportLayoutuDoExcelu_Finalni_s_Dokumentaci"
visioDoc.Save ' *** ADDED: Save the document to prevent prompts ***
visioDoc.Close
visioApp.Quit
Set visioDoc = Nothing
Set visioApp = Nothing

' --- Step 3: Import Initial Data into InputData.xlsm ---
WScript.Echo "Step 2/12: Importing initial workload into InputData.xlsm..."
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
Set excelWb = excelApp.Workbooks.Open(inputDataFilePath)
excelApp.Run "'" & excelWb.Name & "'!ImportWorkloadAndBufferData", True
excelWb.Close True
Set excelWb = Nothing

' --- Step 4: Sync Data to ObjectData.xlsm ---
WScript.Echo "Step 3/12: Syncing data to ObjectData.xlsm..."
Set excelWb = excelApp.Workbooks.Open(objectDataFilePath)
excelApp.Run "'" & excelWb.Name & "'!UpdateFromInputData", True

' --- Step 5: Run First Cycle Optimization ---
WScript.Echo "Step 4/12: Running First Cycle placement optimization..."
excelApp.Run "'" & excelWb.Name & "'!RunFirstCycle_Placement", True

' --- Step 6: Generate All Matrices for Analysis ---
WScript.Echo "Step 5/12: Generating all four distance matrices..."
excelApp.Run "'" & excelWb.Name & "'!GenerateAllMatrices", True

' --- Step 7: Export Matrix for Simulation ---
WScript.Echo "Step 6/12: Exporting distance matrix for simulation..."  & vbCrLf & vbCrLf &  "Speaking of matrix, did you know, that after Matrix Reloaded launch, Ducati received tons of orders for black-painted bikes, but couldn't fulfill them? They had only red color available." & vbCrLf & vbCrLf & "A classic blunder with high volume of supplies but not being able to reflect trends and changes. Totally not Lean..."
excelApp.Run "'" & excelWb.Name & "'!ExportMatrixForSimulation", True
excelWb.Close True
Set excelWb = Nothing
excelApp.Quit
Set excelApp = Nothing

'=======================================================================================
'   PAUSE FOR SIMULATION (MANUAL STEP)
'=======================================================================================
WScript.Echo vbCrLf & "--- PAUSE: WAITING FOR USER ACTION ---"
MsgBox "Phase 1 Complete." & vbCrLf & vbCrLf & "Please perform the following steps:" & vbCrLf & "1. Run your Witness simulation." & vbCrLf & "2. Update the 'Workload' sheet in Data_CD.xlsm with the new results." & vbCrLf & "3. Save and close Data_CD.xlsm." & vbCrLf & vbCrLf & "Click OK to begin Phase 2.", vbInformation, "Action Required: Run Simulation"

'=======================================================================================
'   PHASE 2: FINAL OPTIMIZATION AND VISUALIZATION
'=======================================================================================
WScript.Echo vbCrLf & "--- PHASE 2: STARTING ---" & vbCrLf

' --- Step 8: Import Simulation Feedback ---
WScript.Echo "Step 7/12: Importing simulation feedback into InputData.xlsm..."
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
Set excelWb = excelApp.Workbooks.Open(inputDataFilePath)
excelApp.Run "'" & excelWb.Name & "'!ImportWorkloadAndBufferData", True

' --- Step 9: Recalculate Rack Widths ---
WScript.Echo "Step 8/12: Recalculating new area widths..."
excelApp.Run "'" & excelWb.Name & "'!RecalculateAreaWidths", True
excelWb.Close True
Set excelWb = Nothing

' --- Step 10: Sync Final Data ---
WScript.Echo "Step 9/12: Syncing final data to ObjectData.xlsm..."
Set excelWb = excelApp.Workbooks.Open(objectDataFilePath)
excelApp.Run "'" & excelWb.Name & "'!UpdateFromInputData", True

' --- Step 11: Run Second Cycle Optimization ---
WScript.Echo "Step 10/12: Running Second Cycle placement optimization..."
excelApp.Run "'" & excelWb.Name & "'!RunSecondCycle_Placement", True

' --- Step 12: Run Final Analysis ---
WScript.Echo "Step 11/12: Running final cost analysis..."
excelApp.Run "'" & excelWb.Name & "'!RunFinalAnalysis", True
excelWb.Close True
Set excelWb = Nothing
excelApp.Quit
Set excelApp = Nothing

' --- Step 13: Import Final Layout to Visio ---
WScript.Echo "Step 12/12: Drawing final optimized layout in Visio..." & vbCrLf & vbCrLf & "Autodestruction in 3... 2... 1.. Just kidding!" & vbCrLf
Set visioApp = CreateObject("Visio.Application")
visioApp.Visible = True
Set visioDoc = visioApp.Documents.Open(visioFilePath)
visioDoc.ExecuteLine "ImportLayout_KROK_1_NakreslitVse"
visioDoc.Save ' *** ADDED: Save the newly drawn layout automatically ***

Set visioDoc = Nothing
Set visioApp = Nothing

'=======================================================================================
'      FINAL CLEANUP AND COMPLETION MESSAGE
'=======================================================================================
WScript.Echo vbCrLf & "--- AUTOMATION COMPLETE ---"
MsgBox "Automation Complete!" & vbCrLf & vbCrLf & "The final optimized layout has been generated and saved in Visio.", vbInformation, "Crossdock Optimization"

Set fso = Nothing
Set WshShell = Nothing