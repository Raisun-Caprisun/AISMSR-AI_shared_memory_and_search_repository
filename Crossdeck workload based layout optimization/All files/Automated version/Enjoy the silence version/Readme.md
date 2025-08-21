## ALL CODES OVERVIEW BELOW - FULL CODE BLOCKS OF MACROS.
---

**VISIO - ExportDefault macro**

' =========================================================================================
'
'         FINÁLNÍ SKRIPTY PRO EXPORT LAYOUTU Z VISIA DO EXCELU
'         Verze: 10.1 (Automation-Ready - Bounding Box Calculation RESTORED)
'
' =========================================================================================
Public Sub ExportLayoutuDoExcelu_Finalni_s_Dokumentaci()
    ' --- Declarations ---
    Dim xlApp As Object, xlWbMain As Object, wsMain As Object
    Dim xlWbInput As Object, wsInput As Object
    Dim visShape As Visio.Shape, layoutPage As Visio.Page
    
    ' *** AUTOMATION FIX: Use Relative Paths ***
    Dim docPath As String: docPath = ThisDocument.Path
    Dim mainFilePath As String: mainFilePath = docPath & "ObjectData.xlsm"
    Dim inputDataFilePath As String: inputDataFilePath = docPath & "InputData.xlsm"

    ' --- Header Definitions ---
    Dim mainHeaders As Variant
    mainHeaders = Array("ID", "Name", "Text", "Layer", "Color (RGB)", "CenterX", "CenterY", "Width", "Height", "Angle", "Z-Order", "BBox_Left_X", "BBox_Right_X", "BBox_Bottom_Y", "BBox_Top_Y", "Workload", "New_Width", "New_Center_X", "New_Center_Y", "New_BBox_Left_X", "New_BBox_Right_X", "New_BBox_Bottom_Y", "New_BBox_Top_Y")
    Dim inputDataHeaders As Variant
    inputDataHeaders = Array("ID", "Text", "Layer", "Workload", "New_Width", "Max_Buffer")

    ' --- Bounding Box Variables ---
    Dim bBoxLeftInches As Double, bBoxBottomInches As Double, bBoxRightInches As Double, bBoxTopInches As Double
    Const PALEC_NA_MM As Double = 25.4

    ' --- Find "Layout" Page ---
    Dim p As Visio.Page
    For Each p In ThisDocument.Pages
        If LCase(p.Name) = "layout" Or LCase(p.NameU) = "layout" Then
            Set layoutPage = p
            Exit For
        End If
    Next p
    If layoutPage Is Nothing Then Debug.Print "ERROR: Page 'Layout' not found.": Exit Sub
    If Not Application.ActiveWindow.Page Is layoutPage Then Application.ActiveWindow.Page = layoutPage
    
    ' --- Connect to and Prepare Excel ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then Debug.Print "Could not start Excel.": Exit Sub
    
    xlApp.Visible = False
    Set xlWbMain = xlApp.Workbooks.Open(mainFilePath)
    Set xlWbInput = xlApp.Workbooks.Open(inputDataFilePath)
    If xlWbMain Is Nothing Or xlWbInput Is Nothing Then
        Debug.Print "ERROR: Could not open one of the Excel files. Check paths.": GoTo Cleanup
    End If
    
    Set wsMain = xlWbMain.Worksheets("Layout")
    Set wsInput = xlWbInput.Worksheets(1)
    
    ' Prep main sheet
    Dim lastColMain As Integer: lastColMain = UBound(mainHeaders) + 1
    wsMain.Range("A2:" & wsMain.Cells(50000, lastColMain).Address).ClearContents
    wsMain.Range("A1").Resize(1, lastColMain).Value = mainHeaders
    
    ' Prep input data sheet
    Dim lastColInput As Integer: lastColInput = UBound(inputDataHeaders) + 1
    wsInput.Cells.ClearContents
    wsInput.Range("A1").Resize(1, lastColInput).Value = inputDataHeaders

    ' --- Main Export Loop ---
    Dim i As Long: i = 2
    For Each visShape In ActivePage.Shapes
        ' --- Write Standard Properties ---
        wsMain.Cells(i, "A").Value = visShape.CellsU("Prop.objID").ResultIU
        wsInput.Cells(i, "A").Value = visShape.CellsU("Prop.objID").ResultIU
        wsMain.Cells(i, "C").Value = visShape.Text
        wsInput.Cells(i, "B").Value = visShape.Text
        On Error Resume Next
        wsMain.Cells(i, "D").Value = visShape.Layer(1).Name
        wsInput.Cells(i, "C").Value = visShape.Layer(1).Name
        On Error GoTo 0
        wsMain.Cells(i, "B").Value = visShape.Name
        wsMain.Cells(i, "E").Value = visShape.CellsU("FillForegnd").Result(visColor)
        wsMain.Cells(i, "F").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Result("mm")
        wsMain.Cells(i, "G").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Result("mm")
        wsMain.Cells(i, "H").Value = visShape.CellsU("Width").Result("mm")
        wsMain.Cells(i, "I").Value = visShape.CellsU("Height").Result("mm")
        wsMain.Cells(i, "J").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormAngle).Result("deg")
        wsMain.Cells(i, "K").Value = i - 1 ' Z-Order
        wsInput.Cells(i, "E").Value = visShape.CellsU("Width").Result("mm")
        
        ' *** RESTORED: Bounding Box Calculation and Writing ***
        visShape.BoundingBox visBBoxUprightWH, bBoxLeftInches, bBoxBottomInches, bBoxRightInches, bBoxTopInches
        wsMain.Cells(i, "L").Value = bBoxLeftInches * PALEC_NA_MM
        wsMain.Cells(i, "M").Value = bBoxRightInches * PALEC_NA_MM
        wsMain.Cells(i, "N").Value = bBoxBottomInches * PALEC_NA_MM
        wsMain.Cells(i, "O").Value = bBoxTopInches * PALEC_NA_MM
        ' ******************************************************
        
        i = i + 1
    Next visShape
    
    ' --- Finalize and Cleanup ---
    wsMain.Columns.AutoFit
    wsInput.Columns.AutoFit
    xlWbMain.Close SaveChanges:=True
    xlWbInput.Close SaveChanges:=True

Cleanup:
    If Not xlApp Is Nothing Then xlApp.Quit
    Set wsMain = Nothing: Set wsInput = Nothing
    Set xlWbMain = Nothing: Set xlWbInput = Nothing
    Set xlApp = Nothing
End Sub

---
**VISIO - ImportWorkloadSheetPowerBI macro**
' =========================================================================================
'
'         FINÁLNÍ SKRIPTY PRO IMPORT A AKTUALIZACI LAYOUTU Z EXCELU
'         Verze: 14.0 (Automation-Ready with Relative Paths)
'
' =========================================================================================
Public Sub ImportLayout_KROK_1_NakreslitVse()
    ' --- PARAMETER FOR ZOOM ---
    Const finalZoomLevel As Double = 0.25

    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim targetPage As Visio.Page
    
    ' *** AUTOMATION FIX: Use Relative Paths ***
    ' Assumes ObjectData.xlsm is in the SAME FOLDER as this Visio document.
    Dim filePath As String
    filePath = ThisDocument.Path & "ObjectData.xlsm"
    
    Const MM_NA_PALEC As Double = 1 / 25.4

    Set targetPage = Application.ActiveDocument.Pages(1)
    
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then Exit Sub
    
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(filePath)
    If xlWb Is Nothing Then GoTo Cleanup
    
    Set ws = xlWb.Worksheets("Layout")
    
    Dim i As Long
    For i = targetPage.Shapes.Count To 1 Step -1: targetPage.Shapes(i).Delete: Next i

    Dim lastRow As Long: lastRow = ws.UsedRange.Rows.Count
    If lastRow <= 1 Then GoTo Cleanup
    
    Dim inboundShape As Visio.Shape
    
    ' --- Main import loop ---
    For i = 2 To lastRow
        Dim layerName As Variant: layerName = ws.Cells(i, "D").Value
        Dim width_val As Variant, centerX_val As Variant, centerY_val As Variant
        
        If Trim(LCase(CStr(layerName))) Like "area*" Then
            width_val = ws.Cells(i, "Q").Value
            centerX_val = ws.Cells(i, "R").Value
            centerY_val = ws.Cells(i, "S").Value
        Else
            width_val = ws.Cells(i, "H").Value
            centerX_val = ws.Cells(i, "F").Value
            centerY_val = ws.Cells(i, "G").Value
        End If
        
        Dim height_val As Variant: height_val = ws.Cells(i, "I").Value
        
        If IsNumeric(centerX_val) And IsNumeric(centerY_val) And IsNumeric(width_val) And IsNumeric(height_val) Then
            Dim visShape As Visio.Shape
            Set visShape = targetPage.DrawRectangle((CDbl(centerX_val) - CDbl(width_val) / 2) * MM_NA_PALEC, _
                                                     (CDbl(centerY_val) - CDbl(height_val) / 2) * MM_NA_PALEC, _
                                                     (CDbl(centerX_val) + CDbl(width_val) / 2) * MM_NA_PALEC, _
                                                     (CDbl(centerY_val) + CDbl(height_val) / 2) * MM_NA_PALEC)
            
            Dim shapeText As Variant: shapeText = ws.Cells(i, "C").Value
            visShape.Text = CStr(shapeText)
            
            ' Apply properties
            visShape.CellsU("Angle").Result("deg") = CDbl(Nz(ws.Cells(i, "J").Value))
            visShape.CellsU("FillForegnd").Result(visColor) = CLng(Nz(ws.Cells(i, "E").Value))
            visShape.CellsU("Char.Size").Result("pt") = 36
            
            If LCase(CStr(shapeText)) = "inbound" Then Set inboundShape = visShape

            ' Add to layer
            If Trim(CStr(layerName)) <> "" Then
                Dim visLayer As Visio.Layer
                On Error Resume Next
                Set visLayer = targetPage.Layers.ItemU(CStr(layerName))
                On Error GoTo 0
                If visLayer Is Nothing Then Set visLayer = targetPage.Layers.Add(CStr(layerName))
                visLayer.Add visShape, 1
            End If
            
            visShape.BringToFront
        End If
    Next i
    
    ' Hide the "Zones" layer
    On Error Resume Next
    targetPage.Layers.ItemU("Zones").CellsC(visLayerVisible).ResultIU = 0
    On Error GoTo 0
    
    ' Center view and zoom
    If Not inboundShape Is Nothing Then
        Application.ActiveWindow.CenterViewOnShape inboundShape, visCenterView
        Application.ActiveWindow.Zoom = finalZoomLevel
    End If

Cleanup:
    If Not xlWb Is Nothing Then xlWb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set ws = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
End Sub

Private Function Nz(ByVal Value As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
    If IsNull(Value) Or IsEmpty(Value) Or IsError(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function

---
**EXCEL - InputData.xlsm - ImportDataFromData_CD macro**
'---------------------------------------------------------------------------------------
' Module: DataImporters
' Version: 2.1 - Silent Mode Enabled
' Description: This macro runs from "InputData.xlsm" and pulls data from the "Workload"
'              sheet in "Data_CD.xlsm". It is now fully portable for automation and
'              can be run silently.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub ImportWorkloadAndBufferData(Optional ByVal silent As Boolean = False)
    ' *** AUTOMATION-READY: Uses a Relative Path ***
    Dim sourceFilePath As String
    sourceFilePath = ThisWorkbook.Path & "\Data_CD.xlsm"
    
    Const sourceSheetName As String = "Workload"
    
    ' --- SETUP ---
    Dim sourceWb As Workbook, sourceSheet As Worksheet
    Dim destWb As Workbook, destSheet As Worksheet
    
    Set destWb = ThisWorkbook
    Set destSheet = destWb.Worksheets(1) ' Assumes data is on the first sheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Open the source workbook (Data_CD.xlsm) ---
    On Error Resume Next
    Set sourceWb = Workbooks.Open(sourceFilePath, ReadOnly:=True)
    If sourceWb Is Nothing Then
        If Not silent Then MsgBox "Failed to open the source data file at the specified path:" & vbCrLf & sourceFilePath, vbCritical
        GoTo Cleanup
    End If
    
    On Error Resume Next
    Set sourceSheet = sourceWb.Worksheets(sourceSheetName)
    If sourceSheet Is Nothing Then
        If Not silent Then MsgBox "The sheet named '" & sourceSheetName & "' was not found in the source file.", vbCritical
        sourceWb.Close SaveChanges:=False
        GoTo Cleanup
    End If
    On Error GoTo 0
    
    ' --- 1. Load data from Data_CD.xlsm into a Dictionary ---
    Dim dataToImport As Object
    Set dataToImport = CreateObject("Scripting.Dictionary")
    
    Dim lastSourceRow As Long
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastSourceRow ' Assuming headers are in row 1
        Dim keyID As Variant: keyID = sourceSheet.Cells(r, "A").Value
        
        If Not IsEmpty(keyID) And IsNumeric(keyID) Then
            Dim keyIDLong As Long: keyIDLong = CLng(keyID)
            
            If Not dataToImport.Exists(keyIDLong) Then
                Dim workloadValue As Variant: workloadValue = sourceSheet.Cells(r, "B").Value
                Dim bufferValue As Variant: bufferValue = sourceSheet.Cells(r, "C").Value
                dataToImport.Add keyIDLong, Array(workloadValue, bufferValue)
            End If
        End If
    Next r
    
    ' --- Close the source workbook ---
    sourceWb.Close SaveChanges:=False
    
    ' --- 2. Find the required columns in THIS destination file ---
    Dim colDestID As Long: colDestID = FindHeaderColumn(destSheet, "ID")
    Dim colDestWorkload As Long: colDestWorkload = FindHeaderColumn(destSheet, "Workload")
    Dim colDestBuffer As Long: colDestBuffer = FindHeaderColumn(destSheet, "Max_Buffer")
    
    If colDestID = 0 Or colDestWorkload = 0 Or colDestBuffer = 0 Then
        If Not silent Then MsgBox "Could not find one or more required columns ('ID', 'Workload', 'Max_Buffer') in this file.", vbCritical
        GoTo Cleanup
    End If
    
    ' --- 3. Update the "Workload" and "Max_Buffer" columns in this sheet ---
    Dim lastDestRow As Long
    lastDestRow = destSheet.Cells(destSheet.Rows.Count, colDestID).End(xlUp).Row
    
    Dim updatedCount As Long: updatedCount = 0
    For r = 2 To lastDestRow
        Dim destKeyID As Variant
        destKeyID = destSheet.Cells(r, colDestID).Value
        
        If Not IsEmpty(destKeyID) And IsNumeric(destKeyID) Then
            Dim destKeyIDLong As Long: destKeyIDLong = CLng(destKeyID)
            
            If dataToImport.Exists(destKeyIDLong) Then
                Dim dataArray As Variant
                dataArray = dataToImport(destKeyIDLong)
                
                destSheet.Cells(r, colDestWorkload).Value = dataArray(0) ' Workload
                destSheet.Cells(r, colDestBuffer).Value = dataArray(1)   ' Max_Buffer
                
                updatedCount = updatedCount + 1
            End If
        End If
    Next r
    
    ' *** SILENT MODE IMPLEMENTED ***
    If Not silent Then
        MsgBox "Import complete." & vbCrLf & vbCrLf & "Updated " & updatedCount & " rows from the Data_CD file.", vbInformation
    End If

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Set sourceSheet = Nothing
    Set sourceWb = Nothing
    Set destSheet = Nothing
    Set destWb = Nothing
    Set dataToImport = Nothing
End Sub


Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    ' Helper function to find a column number by its header text in the first row.
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function

---
**EXCEL - InputData.xlsm - WidthOptimizer Macro**

'---------------------------------------------------------------------------------------
' Module: WidthOptimizer
' Version: 1.1 - Silent Mode Enabled
' Description: This macro recalculates the 'New_Width' column for all "Areas" based
'              on a tiered system. It can now be run silently for automation.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub RecalculateAreaWidths(Optional ByVal silent As Boolean = False)
    ' --- CONFIGURATION ---
    Const TIER_1_WIDTH As Long = 7200
    Const TIER_2_WIDTH As Long = 4800
    Const TIER_3_WIDTH As Long = 2400
    
    Const TIER_1_PERCENT As Double = 0.03333
    Const TIER_2_PERCENT As Double = 0.21666
    
    Const MAX_TOTAL_WIDTH As Long = 184800
    
    ' --- SETUP ---
    Dim inputSheet As Worksheet
    Set inputSheet = ThisWorkbook.Worksheets(1) ' Assumes data is on the first sheet
    
    Application.ScreenUpdating = False
    
    ' --- Find all necessary columns by header name ---
    Dim cols As Object: Set cols = CreateObject("Scripting.Dictionary")
    cols.Add "Text", FindHeaderColumn(inputSheet, "Text")
    cols.Add "Layer", FindHeaderColumn(inputSheet, "Layer")
    cols.Add "New_Width", FindHeaderColumn(inputSheet, "New_Width")
    cols.Add "Max_Buffer", FindHeaderColumn(inputSheet, "Max_Buffer")
    
    If cols("Layer") = 0 Or cols("Max_Buffer") = 0 Or cols("New_Width") = 0 Then
        If Not silent Then MsgBox "Could not find required columns ('Layer', 'Max_Buffer', 'New_Width'). Please check the headers.", vbCritical
        Exit Sub
    End If
    
    ' --- 1. Read and Filter all "Areas" data ---
    Dim areasData As Collection
    Set areasData = LoadAreaData(inputSheet, cols)
    
    If areasData.Count = 0 Then
        If Not silent Then MsgBox "No rows with Layer = 'Areas' were found.", vbInformation
        Exit Sub
    End If
    
    ' --- 2. Sort the areas by Max_Buffer in descending order ---
    Dim sortedAreas As Collection
    Set sortedAreas = SortByMetric(areasData, "Max_Buffer", False) ' False for Descending
    
    ' --- 3. Calculate the number of slots for each tier ---
    Dim totalAreas As Long: totalAreas = sortedAreas.Count
    Dim numTier1 As Long: numTier1 = Round(totalAreas * TIER_1_PERCENT, 0)
    Dim numTier2 As Long: numTier2 = Round(totalAreas * TIER_2_PERCENT, 0)
    
    ' --- 4 & 5. Assign New Widths, Calculate Sum, and Write Back to Sheet ---
    Dim totalNewWidth As Double: totalNewWidth = 0
    Dim i As Long
    Dim currentArea As Object
    Dim newWidth As Long
    
    For i = 1 To totalAreas
        Set currentArea = sortedAreas(i)
        
        ' Assign width based on tiered position
        If i <= numTier1 Then
            newWidth = TIER_1_WIDTH
        ElseIf i <= (numTier1 + numTier2) Then
            newWidth = TIER_2_WIDTH
        Else
            newWidth = TIER_3_WIDTH
        End If
        
        ' Write the new width back to the correct row in the Excel sheet
        inputSheet.Cells(currentArea("Row"), cols("New_Width")).Value = newWidth
        
        ' Add to the running total
        totalNewWidth = totalNewWidth + newWidth
    Next i
    
    ' --- 6. Verify Constraint and Report to User ---
    If Not silent Then
        Dim msg As String
        msg = "Width recalculation complete." & vbCrLf & vbCrLf & _
              "Total Areas Processed: " & totalAreas & vbCrLf & _
              "Tier 1 Slots (@" & TIER_1_WIDTH & "mm): " & numTier1 & vbCrLf & _
              "Tier 2 Slots (@" & TIER_2_WIDTH & "mm): " & numTier2 & vbCrLf & _
              "Tier 3 Slots (@" & TIER_3_WIDTH & "mm): " & totalAreas - numTier1 - numTier2 & vbCrLf & vbCrLf & _
              "Calculated Total Width: " & totalNewWidth & vbCrLf & _
              "Constraint Limit: " & MAX_TOTAL_WIDTH & vbCrLf & vbCrLf
              
        If totalNewWidth <= MAX_TOTAL_WIDTH Then
            msg = msg & "Result: CONSTRAINT MET."
            MsgBox msg, vbInformation
        Else
            msg = msg & "Result: WARNING - CONSTRAINT FAILED."
            MsgBox msg, vbExclamation
        End If
    End If
    
    Application.ScreenUpdating = True
End Sub


Private Function LoadAreaData(ByVal ws As Worksheet, ByVal cols As Object) As Collection
    ' Reads all rows with Layer="Areas" into a collection of dictionary objects
    Set LoadAreaData = New Collection
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        If Trim(LCase(CStr(ws.Cells(r, cols("Layer")).Value))) Like "area*" Then
            Dim area As Object: Set area = CreateObject("Scripting.Dictionary")
            area.Add "Row", r ' Store original row number
            area.Add "Text", ws.Cells(r, cols("Text")).Value
            area.Add "Max_Buffer", CDbl(Nz(ws.Cells(r, cols("Max_Buffer")).Value))
            LoadAreaData.Add area
        End If
    Next r
End Function

Private Function SortByMetric(ByVal coll As Collection, ByVal sortKey As String, Optional ByVal ascending As Boolean = True) As Collection
    ' Sorts a collection of dictionary objects by a specified key
    If coll.Count <= 1 Then Set SortByMetric = coll: Exit Function
    
    Dim i As Long, j As Long, temp As Object
    Dim arr As New Collection
    For Each temp In coll: arr.Add temp: Next
    
    For i = 1 To arr.Count - 1
        For j = i + 1 To arr.Count
            Dim condition As Boolean
            If ascending Then
                condition = (arr(i)(sortKey) > arr(j)(sortKey))
            Else
                condition = (arr(i)(sortKey) < arr(j)(sortKey))
            End If
            
            If condition Then
                Set temp = arr(j)
                arr.Remove j
                arr.Add temp, before:=i
            End If
        Next j
    Next i
    Set SortByMetric = arr
End Function


' --- Self-Contained Helper Functions ---
Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function

Private Function Nz(ByVal Value As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
    If IsNull(Value) Or IsEmpty(Value) Or IsError(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function


---
**EXCEL - ObjectData.xlsm - GetWorkloadWidthFromInputData macro**

'---------------------------------------------------------------------------------------
' Module: DataUpdaters
' Version: 4.3 - Silent Mode Enabled
' Description: This version corrects the core logic to use the 'ID' column as the
'              unique key for matching rows, ensuring a reliable data transfer.
'              It can now be run silently for automation.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub UpdateFromInputData(Optional ByVal silent As Boolean = False)
    ' *** AUTOMATION-READY: Uses a Relative Path ***
    Dim inputDataPath As String
    inputDataPath = ThisWorkbook.Path & "\InputData.xlsm"
    
    Dim inputWb As Workbook, inputSheet As Worksheet
    Dim objectDataWb As Workbook, layoutSheet As Worksheet
    
    Set objectDataWb = ThisWorkbook
    Set layoutSheet = objectDataWb.Worksheets("Layout")
    If layoutSheet Is Nothing Then
        If Not silent Then MsgBox "The active workbook does not contain a 'Layout' sheet.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Set inputWb = Workbooks.Open(inputDataPath, ReadOnly:=True)
    If inputWb Is Nothing Then
        If Not silent Then MsgBox "Failed to open InputData file at path:" & vbCrLf & inputDataPath, vbCritical
        GoTo Cleanup
    End If
    Set inputSheet = inputWb.Worksheets(1)
    On Error GoTo 0
    
    ' --- Dynamically find all necessary columns in the source file ---
    Dim colInputID As Long: colInputID = FindHeaderColumn(inputSheet, "ID")
    Dim colInputLayer As Long: colInputLayer = FindHeaderColumn(inputSheet, "Layer")
    Dim colInputWorkload As Long: colInputWorkload = FindHeaderColumn(inputSheet, "Workload")
    Dim colInputNewWidth As Long: colInputNewWidth = FindHeaderColumn(inputSheet, "New_Width")
    
    If colInputID = 0 Or colInputLayer = 0 Or colInputWorkload = 0 Or colInputNewWidth = 0 Then
        If Not silent Then MsgBox "Could not find required columns ('ID', 'Layer', 'Workload', 'New_Width') in 'InputData.xlsm'.", vbCritical
        inputWb.Close SaveChanges:=False
        GoTo Cleanup
    End If

    ' --- Load data into dictionaries using ID as the key ---
    Dim workloadData As Object: Set workloadData = CreateObject("Scripting.Dictionary")
    Dim newWidthData As Object: Set newWidthData = CreateObject("Scripting.Dictionary")

    Dim lastInputRow As Long: lastInputRow = inputSheet.Cells(inputSheet.Rows.count, colInputID).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastInputRow
        Dim keyID As Variant
        keyID = inputSheet.Cells(r, colInputID).Value
        
        If Not IsEmpty(keyID) And IsNumeric(keyID) Then
            Dim keyIDLong As Long: keyIDLong = CLng(keyID)

            ' --- Logic for New_Width (applies to ALL layers) ---
            If Not newWidthData.Exists(keyIDLong) Then
                newWidthData.Add keyIDLong, CDbl(Nz(inputSheet.Cells(r, colInputNewWidth).Value))
            End If

            ' --- Logic for Workload (applies ONLY to "Areas") ---
            Dim layerName As String
            layerName = Trim(LCase(CStr(inputSheet.Cells(r, colInputLayer).Value)))
            If layerName Like "area*" Then
                If Not workloadData.Exists(keyIDLong) Then
                    workloadData.Add keyIDLong, CDbl(Nz(inputSheet.Cells(r, colInputWorkload).Value))
                End If
            End If
        End If
    Next r
    
    inputWb.Close SaveChanges:=False
    
    ' --- Find destination columns ---
    Dim colDestID As Long: colDestID = FindHeaderColumn(layoutSheet, "ID")
    Dim colDestLayer As Long: colDestLayer = FindHeaderColumn(layoutSheet, "Layer")
    Dim colDestWorkload As Long: colDestWorkload = FindHeaderColumn(layoutSheet, "Workload")
    Dim colDestNewWidth As Long: colDestNewWidth = FindHeaderColumn(layoutSheet, "New_Width")
    
    If colDestID = 0 Or colDestLayer = 0 Or colDestWorkload = 0 Or colDestNewWidth = 0 Then
        If Not silent Then MsgBox "Could not find required columns in this workbook's 'Layout' sheet.", vbCritical
        GoTo Cleanup
    End If
    
    ' --- Update columns in this sheet by matching ID ---
    Dim lastLayoutRow As Long: lastLayoutRow = layoutSheet.Cells(layoutSheet.Rows.count, colDestID).End(xlUp).Row
    
    Dim updatedCount As Long: updatedCount = 0
    For r = 2 To lastLayoutRow
        Dim layoutKeyID As Variant
        layoutKeyID = layoutSheet.Cells(r, colDestID).Value
        
        If Not IsEmpty(layoutKeyID) And IsNumeric(layoutKeyID) Then
            Dim layoutKeyIDLong As Long: layoutKeyIDLong = CLng(layoutKeyID)
            Dim didUpdate As Boolean: didUpdate = False
            
            Dim destLayerName As String
            destLayerName = Trim(LCase(CStr(layoutSheet.Cells(r, colDestLayer).Value)))
            
            If destLayerName Like "area*" Then
                If workloadData.Exists(layoutKeyIDLong) Then
                    layoutSheet.Cells(r, colDestWorkload).Value = workloadData(layoutKeyIDLong)
                Else
                    layoutSheet.Cells(r, colDestWorkload).Value = 0
                End If
                didUpdate = True
            End If
            
            If newWidthData.Exists(layoutKeyIDLong) Then
                layoutSheet.Cells(r, colDestNewWidth).Value = newWidthData(layoutKeyIDLong)
                didUpdate = True
            End If
            
            If didUpdate Then updatedCount = updatedCount + 1
        End If
    Next r
    
    objectDataWb.Save
    ' *** SILENT MODE IMPLEMENTED ***
    If Not silent Then
        MsgBox "Update complete." & vbCrLf & vbCrLf & updatedCount & " rows were updated from the InputData file.", vbInformation
    End If

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Set inputSheet = Nothing
    Set inputWb = Nothing
    Set layoutSheet = Nothing
    Set objectDataWb = Nothing
    Set workloadData = Nothing
    Set newWidthData = Nothing
End Sub


Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function

Private Function Nz(ByVal Value As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
    If IsNull(Value) Or IsEmpty(Value) Or IsError(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function

---
**EXCEL - ObjectData.xlsm - LayoutCostCalculator macro**

'---------------------------------------------------------------------------------------
' Module: AnalysisTools
' Version: 12.1 - Silent Mode Enabled
' Description: This version updates the analysis macros to be run silently for automation.
'---------------------------------------------------------------------------------------
Option Explicit

'========================================================================================
'      PUBLIC MASTER MACRO - This is the recommended macro to run.
'========================================================================================

Public Sub RunFinalAnalysis(Optional ByVal silent As Boolean = False)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Pass the silent flag down to the core calculation sub
    CalculateAllLayoutCosts showMsg:=False, beSilent:=silent

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Not silent Then
        MsgBox "The final Cost Calculation analysis is complete.", vbInformation
    End If
End Sub


'========================================================================================
'      CORE CALCULATION LOGIC
'========================================================================================

Public Sub CalculateAllLayoutCosts(Optional ByVal showMsg As Boolean = True, Optional ByVal beSilent As Boolean = False)
    ' --- This is the scale factor from your Visio drawing (10mm = 1m) ---
    Const SCALE_FACTOR_MM_PER_METER As Double = 10

    Dim layoutSheet As Worksheet, costSheet As Worksheet

    ' --- Setup and Sheet Preparation ---
    On Error Resume Next
    Set layoutSheet = ThisWorkbook.Worksheets("Layout")
    If layoutSheet Is Nothing Then
        If Not beSilent Then MsgBox "The 'Layout' worksheet was not found. Cannot proceed.", vbCritical
        Exit Sub
    End If

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Cost_Calculation").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set costSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    costSheet.Name = "Cost_Calculation"
    
    ' --- Calculate all four scenarios ---
    Dim results(1 To 4, 1 To 6) As Double
    Dim resultData As Variant

    resultData = CalculateSingleCost(layoutSheet, False, True, SCALE_FACTOR_MM_PER_METER)
    results(1, 1) = resultData(0): results(1, 2) = resultData(1): results(1, 3) = resultData(2) / 1000
    results(1, 4) = resultData(0) * 2: results(1, 5) = resultData(1) * 2: results(1, 6) = (resultData(2) / 1000) * 2
    
    resultData = CalculateSingleCost(layoutSheet, False, False, SCALE_FACTOR_MM_PER_METER)
    results(2, 1) = resultData(0): results(2, 2) = resultData(1): results(2, 3) = resultData(2) / 1000
    results(2, 4) = resultData(0) * 2: results(2, 5) = resultData(1) * 2: results(2, 6) = (resultData(2) / 1000) * 2
    
    resultData = CalculateSingleCost(layoutSheet, True, True, SCALE_FACTOR_MM_PER_METER)
    results(3, 1) = resultData(0): results(3, 2) = resultData(1): results(3, 3) = resultData(2) / 1000
    results(3, 4) = resultData(0) * 2: results(3, 5) = resultData(1) * 2: results(3, 6) = (resultData(2) / 1000) * 2

    resultData = CalculateSingleCost(layoutSheet, True, False, SCALE_FACTOR_MM_PER_METER)
    results(4, 1) = resultData(0): results(4, 2) = resultData(1): results(4, 3) = resultData(2) / 1000
    results(4, 4) = resultData(0) * 2: results(4, 5) = resultData(1) * 2: results(4, 6) = (resultData(2) / 1000) * 2

    ' --- Write results to the sheet with NEW, CLEARER HEADERS ---
    With costSheet
        .Range("A1:G1").Value = Array("Scenario", "Weighted Travel Cost (One-Way)", "Avg Travel per Item (m)", "Total Travel Distance (km)", "Weighted Travel Cost (Round-Trip)", "Avg Round-Trip per Item (m)", "Total Round-Trip Travel (km)")
        .Range("A2:A5").Value = Application.Transpose(Array("Default Layout - Euclidean", "Default Layout - Manhattan", "Optimized Layout - Euclidean", "Optimized Layout - Manhattan"))
        .Range("B2:G5").Value = results
        
        .Columns("A:G").AutoFit
        .Range("A1:G1").Font.Bold = True
        .Range("A2:A5").Font.Bold = True
        .Columns("B:G").NumberFormat = "#,##0.00"
        .Activate
        .Cells(1, 1).Select
    End With

    If showMsg And Not beSilent Then
        MsgBox "The Cost Calculation analysis is complete.", vbInformation
    End If
End Sub

Private Function CalculateSingleCost(ByVal layoutSheet As Worksheet, ByVal isOptimized As Boolean, ByVal useEuclidean As Boolean, ByVal scaleFactor As Double) As Variant
    Dim r As Long, lastRow As Long
    Dim totalCost As Double, totalWorkload As Double, totalWorkloadDistance As Double
    Dim colLayer As Long, colWorkload As Long, colX As Long, colY As Long
    
    colLayer = FindHeaderColumn(layoutSheet, "Layer")
    colWorkload = FindHeaderColumn(layoutSheet, "Workload")
    
    If isOptimized Then
        colX = FindHeaderColumn(layoutSheet, "New_Center_X"): colY = FindHeaderColumn(layoutSheet, "New_Center_Y")
    Else
        colX = FindHeaderColumn(layoutSheet, "CenterX"): colY = FindHeaderColumn(layoutSheet, "CenterY")
    End If
    
    lastRow = layoutSheet.Cells(layoutSheet.Rows.count, "A").End(xlUp).Row
    
    Dim inboundX As Double, inboundY As Double
    inboundX = GetInboundCoord(layoutSheet, lastRow, colLayer, colX)
    inboundY = GetInboundCoord(layoutSheet, lastRow, colLayer, colY)
    
    If inboundX = -1 Or colWorkload = 0 Or scaleFactor = 0 Then
        CalculateSingleCost = Array(0, 0, 0): Exit Function
    End If
    
    For r = 2 To lastRow
        If Trim(LCase(CStr(layoutSheet.Cells(r, colLayer).Value))) Like "area*" Then
            Dim workloadVal As Variant, areaXVal As Variant
            workloadVal = layoutSheet.Cells(r, colWorkload).Value
            areaXVal = layoutSheet.Cells(r, colX).Value
            Dim workload As Double
            workload = CDbl(Nz(workloadVal))
            If IsNumeric(areaXVal) And workload > 0 Then
                totalWorkload = totalWorkload + workload
                Dim areaX As Double, areaY As Double
                areaX = CDbl(areaXVal)
                areaY = CDbl(layoutSheet.Cells(r, colY).Value)
                Dim distance As Double, realDistance As Double
                If useEuclidean Then distance = Sqr((areaX - inboundX) ^ 2 + (areaY - inboundY) ^ 2) Else distance = Abs(areaX - inboundX) + Abs(areaY - inboundY)
                realDistance = distance / scaleFactor
                totalCost = totalCost + (workload * realDistance)
                totalWorkloadDistance = totalWorkloadDistance + (workload * realDistance)
            End If
        End If
    Next r
    
    Dim avgWeightedDistance As Double
    If totalWorkload > 0 Then avgWeightedDistance = totalCost / totalWorkload Else avgWeightedDistance = 0
    
    CalculateSingleCost = Array(totalCost, avgWeightedDistance, totalWorkloadDistance)
End Function


'========================================================================================
'      SELF-CONTAINED HELPER FUNCTIONS
'========================================================================================
Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function

Private Function GetInboundCoord(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal colLayer As Long, ByVal colCoord As Long) As Double
    Dim r As Long
    For r = 2 To lastRow
        If Trim(LCase(CStr(ws.Cells(r, colLayer).Value))) = "inbound" Then
            GetInboundCoord = CDbl(Nz(ws.Cells(r, colCoord).Value))
            Exit Function
        End If
    Next r
    GetInboundCoord = -1
End Function

Private Function Nz(ByVal Value As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
    If IsNull(Value) Or IsEmpty(Value) Or IsError(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function

---
**EXCEL - ObjectData.xlsm - MatrixAnalyzers macro**

'---------------------------------------------------------------------------------------
' Module: MatrixAnalyzers
' Version: 5.1 - Silent Mode Enabled
' Description: This version updates the public macros to be run silently for automation.
'---------------------------------------------------------------------------------------
Option Explicit

'========================================================================================
'      PUBLIC MACROS - These are the "buttons" you can run from the Macros menu.
'========================================================================================

Public Sub ExportMatrixForSimulation(Optional ByVal silent As Boolean = False)
    ' *** AUTOMATION-READY: Uses a Relative Path ***
    Dim destFilePath As String
    destFilePath = ThisWorkbook.Path & "\Data_CD.xlsm"
    
    Const sourceSheetName As String = "Matrix_Optimized_Euclidean"
    Const destSheetName As String = "MaticeVzdalenosti"
    Const DIVISOR As Double = 1000 ' To convert mm to m
    
    Dim sourceWb As Workbook: Set sourceWb = ThisWorkbook
    Dim sourceSheet As Worksheet, destWb As Workbook, destSheet As Worksheet
    Dim destFileName As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- 1. Verify that the source matrix exists ---
    On Error Resume Next
    Set sourceSheet = sourceWb.Worksheets(sourceSheetName)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        If Not silent Then MsgBox "The source sheet '" & sourceSheetName & "' was not found." & vbCrLf & "Please generate the optimized matrices first before running this export.", vbCritical
        GoTo Cleanup
    End If
    
    ' --- 2. Open the destination workbook ---
    On Error Resume Next
    Set destWb = Workbooks.Open(destFilePath)
    If destWb Is Nothing Then
        If Not silent Then MsgBox "Failed to open the destination file for the simulation:" & vbCrLf & destFilePath, vbCritical
        GoTo Cleanup
    End If
    destFileName = destWb.Name
    
    ' --- 3. Prepare the destination sheet ---
    On Error Resume Next
    Set destSheet = destWb.Worksheets(destSheetName)
    On Error GoTo 0
    If destSheet Is Nothing Then
        Set destSheet = destWb.Worksheets.Add
        destSheet.Name = destSheetName
    End If
    destSheet.Cells.Clear
    
    ' --- 4. Copy data with transformation ---
    Dim lastRow As Long, lastCol As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, 1).End(xlUp).Row
    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.count).End(xlToLeft).Column
    
    Dim r As Long, c As Long
    sourceSheet.Rows(1).Copy destSheet.Rows(1)
    sourceSheet.Columns(1).Copy destSheet.Columns(1)
    
    For r = 2 To lastRow
        For c = 2 To lastCol
            If IsNumeric(sourceSheet.Cells(r, c).Value) Then
                destSheet.Cells(r, c).Value = sourceSheet.Cells(r, c).Value / DIVISOR
            End If
        Next c
    Next r
    
    ' --- 5. Save and Close ---
    destSheet.Columns.AutoFit
    destWb.Close SaveChanges:=True
    
    If Not silent Then
        MsgBox "The distance matrix has been successfully exported and scaled for the simulation in '" & destFileName & "'.", vbInformation
    End If

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Public Sub GenerateAllMatrices(Optional ByVal silent As Boolean = False)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    GenerateMatrix_Master useOptimizedCoords:=True, useEuclideanCalc:=True, showMsg:=False
    GenerateMatrix_Master useOptimizedCoords:=True, useEuclideanCalc:=False, showMsg:=False
    GenerateMatrix_Master useOptimizedCoords:=False, useEuclideanCalc:=True, showMsg:=False
    GenerateMatrix_Master useOptimizedCoords:=False, useEuclideanCalc:=False, showMsg:=False
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Not silent Then
        MsgBox "All four distance matrices have been successfully generated.", vbInformation
    End If
End Sub

' The subs below do not need the silent parameter as they are only called by the master sub above
Public Sub GenerateOptimizedMatrix_Euclidean()
    GenerateMatrix_Master useOptimizedCoords:=True, useEuclideanCalc:=True
End Sub

Public Sub GenerateOptimizedMatrix_Manhattan()
    GenerateMatrix_Master useOptimizedCoords:=True, useEuclideanCalc:=False
End Sub

Public Sub GenerateDefaultMatrix_Euclidean()
    GenerateMatrix_Master useOptimizedCoords:=False, useEuclideanCalc:=True
End Sub

Public Sub GenerateDefaultMatrix_Manhattan()
    GenerateMatrix_Master useOptimizedCoords:=False, useEuclideanCalc:=False
End Sub

Private Sub GenerateMatrix_Master(ByVal useOptimizedCoords As Boolean, ByVal useEuclideanCalc As Boolean, Optional ByVal showMsg As Boolean = True)
    Dim layoutSheet As Worksheet, matrixSheet As Worksheet
    Dim sheetName As String, msg As String
    Dim colXName As String, colYName As String
    
    If useOptimizedCoords Then
        colXName = "New_Center_X": colYName = "New_Center_Y"
        If useEuclideanCalc Then sheetName = "Matrix_Optimized_Euclidean": msg = "Optimized Euclidean distance matrix created." Else sheetName = "Matrix_Optimized_Manhattan": msg = "Optimized Manhattan distance matrix created."
    Else
        colXName = "CenterX": colYName = "CenterY"
        If useEuclideanCalc Then sheetName = "Matrix_Default_Euclidean": msg = "Default Euclidean distance matrix created." Else sheetName = "Matrix_Default_Manhattan": msg = "Default Manhattan distance matrix created."
    End If
    
    Set layoutSheet = ThisWorkbook.Worksheets("Layout")
    If layoutSheet Is Nothing Then Exit Sub
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set matrixSheet = ThisWorkbook.Worksheets.Add(After:=layoutSheet)
    matrixSheet.Name = sheetName
    
    Dim colId As Long: colId = FindHeaderColumn(layoutSheet, "ID")
    Dim colLayer As Long: colLayer = FindHeaderColumn(layoutSheet, "Layer")
    Dim colX As Long: colX = FindHeaderColumn(layoutSheet, colXName)
    Dim colY As Long: colY = FindHeaderColumn(layoutSheet, colYName)
    
    If colId = 0 Or colLayer = 0 Or colX = 0 Or colY = 0 Then GoTo Cleanup
    
    Dim objects As Collection: Set objects = LoadFilteredObjects(layoutSheet, colId, colLayer, colX, colY)
    If objects.count = 0 Then GoTo Cleanup
    
    Dim sortedObjects As Collection: Set sortedObjects = SortObjectsByID(objects)
    
    matrixSheet.Cells.Clear
    Dim i As Long
    For i = 1 To sortedObjects.count
        matrixSheet.Cells(i + 1, 1).Value = sortedObjects(i)("ID")
        matrixSheet.Cells(1, i + 1).Value = sortedObjects(i)("ID")
    Next i
    
    Dim j As Long, distance As Double
    For i = 1 To sortedObjects.count
        For j = 1 To sortedObjects.count
            If useEuclideanCalc Then distance = Sqr((sortedObjects(j)("X") - sortedObjects(i)("X")) ^ 2 + (sortedObjects(j)("Y") - sortedObjects(i)("Y")) ^ 2) Else distance = Abs(sortedObjects(j)("X") - sortedObjects(i)("X")) + Abs(sortedObjects(j)("Y") - sortedObjects(i)("Y"))
            matrixSheet.Cells(i + 1, j + 1).Value = distance
        Next j
    Next i
    
    FormatMatrixSheet matrixSheet, sortedObjects.count
    If showMsg Then MsgBox msg, vbInformation

Cleanup:
    If showMsg Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
End Sub
Private Function LoadFilteredObjects(ByVal ws As Worksheet, ByVal cId As Long, ByVal cLyr As Long, ByVal cX As Long, ByVal cY As Long) As Collection
    Set LoadFilteredObjects = New Collection
    Dim r As Long, lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For r = 2 To lastRow
        Dim layerName As String: layerName = Trim(LCase(CStr(ws.Cells(r, cLyr).Value)))
        If (layerName = "inbound" Or layerName Like "area*") And IsNumeric(ws.Cells(r, cX).Value) And IsNumeric(ws.Cells(r, cY).Value) Then
            Dim objData As Object: Set objData = CreateObject("Scripting.Dictionary")
            objData.Add "ID", ws.Cells(r, cId).Value
            objData.Add "X", CDbl(ws.Cells(r, cX).Value): objData.Add "Y", CDbl(ws.Cells(r, cY).Value)
            LoadFilteredObjects.Add objData
        End If
    Next r
End Function
Private Function SortObjectsByID(ByVal coll As Collection) As Collection
    If coll.count <= 1 Then Set SortObjectsByID = coll: Exit Function
    Dim i As Long, j As Long, temp As Object, arr As New Collection, item As Object
    For Each item In coll: arr.Add item: Next item
    For i = 1 To arr.count - 1
        For j = i + 1 To arr.count
            If arr(i)("ID") > arr(j)("ID") Then
                Set temp = arr(j): arr.Remove j: arr.Add temp, before:=i
            End If
        Next j
    Next i
    Set SortObjectsByID = arr
End Function
Private Sub FormatMatrixSheet(ByVal ms As Worksheet, ByVal count As Long)
    ms.Columns("A:A").AutoFit
    ms.Rows("1:1").Font.Bold = True: ms.Columns("A:A").Font.Bold = True
    ms.Range(ms.Cells(2, 2), ms.Cells(count + 1, count + 1)).NumberFormat = "0"
    ms.Activate: ms.Cells(1, 1).Select
End Sub
Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function

---
**EXCEL - ObjectData.xlsm - Sorting macro**

'---------------------------------------------------------------------------------------
' Module: LayoutOptimizer
' Version: 3.6 - Silent Mode Enabled
' Description: This version can be run silently for automation. All message boxes
'              are now conditional based on the 'silent' parameter.
'---------------------------------------------------------------------------------------
Option Explicit

'========================================================================================
'      PUBLIC MACROS - RUN THESE FROM THE EXCEL INTERFACE
'========================================================================================

Public Sub RunFirstCycle_Placement(Optional ByVal silent As Boolean = False)
    Debug.Print "Running First Cycle Placement Optimization..."
    RunPlacementOptimization useNewWidth:=False, beSilent:=silent
    Debug.Print "First Cycle Placement complete."
    If Not silent Then
        MsgBox "First Cycle placement optimization is complete. Review the 'New_...' columns.", vbInformation
    End If
End Sub

Public Sub RunSecondCycle_Placement(Optional ByVal silent As Boolean = False)
    Debug.Print "Running Second Cycle Placement Optimization..."
    RunPlacementOptimization useNewWidth:=True, beSilent:=silent
    Debug.Print "Second Cycle Placement complete."
    If Not silent Then
        MsgBox "Second Cycle placement optimization is complete. Review the 'New_...' columns.", vbInformation
    End If
End Sub


'========================================================================================
'      CORE LOGIC - DO NOT RUN DIRECTLY
'========================================================================================

Private Sub RunPlacementOptimization(ByVal useNewWidth As Boolean, ByVal beSilent As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Layout")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- USER-CONFIGURABLE PARAMETERS ---
    Const gap_x_mm As Double = 1300
    Const gap_y_mm As Double = 100
    Const step_mm As Double = 150
    Const PREFERENCE_PENALTY As Double = 1000000
    Dim preferredZones As Variant
    preferredZones = Array("ZONE3", "ZONE4", "ZONE5")
    ' ----------------------------------------

    Dim cols As Object: Set cols = CreateObject("Scripting.Dictionary")
    cols.Add "Layer", FindHeaderColumn(ws, "Layer")
    cols.Add "Workload", FindHeaderColumn(ws, "Workload")
    cols.Add "Text", FindHeaderColumn(ws, "Text")
    cols.Add "OrigX", FindHeaderColumn(ws, "CenterX")
    cols.Add "OrigY", FindHeaderColumn(ws, "CenterY")
    cols.Add "OrigWidth", FindHeaderColumn(ws, "Width")
    cols.Add "OrigHeight", FindHeaderColumn(ws, "Height")
    cols.Add "NewWidth", FindHeaderColumn(ws, "New_Width")
    cols.Add "NewX", FindHeaderColumn(ws, "New_Center_X")
    cols.Add "NewY", FindHeaderColumn(ws, "New_Center_Y")
    cols.Add "NewL", FindHeaderColumn(ws, "New_BBox_Left_X")
    cols.Add "NewR", FindHeaderColumn(ws, "New_BBox_Right_X")
    cols.Add "NewB", FindHeaderColumn(ws, "New_BBox_Bottom_Y")
    cols.Add "NewT", FindHeaderColumn(ws, "New_BBox_Top_Y")
    cols.Add "OrigL", FindHeaderColumn(ws, "BBox_Left_X")
    cols.Add "OrigR", FindHeaderColumn(ws, "BBox_Right_X")
    cols.Add "OrigB", FindHeaderColumn(ws, "BBox_Bottom_Y")
    cols.Add "OrigT", FindHeaderColumn(ws, "BBox_Top_Y")
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, cols("Layer")).End(xlUp).Row
    
    Dim inboundX As Double, inboundY As Double
    inboundX = GetInboundCoord(ws, lastRow, cols("Layer"), cols("OrigX"))
    inboundY = GetInboundCoord(ws, lastRow, cols("Layer"), cols("OrigY"))
    If inboundX = -1 Then
        If Not beSilent Then MsgBox "'Inbound' area not found.", vbCritical
        GoTo Cleanup
    End If

    Dim allZones As Collection: Set allZones = LoadZones(ws, lastRow, cols)
    Dim areas As Collection: Set areas = LoadAreas(ws, lastRow, cols, useNewWidth)
    If areas.count = 0 Then
        If Not beSilent Then MsgBox "No 'Areas' found to optimize.", vbExclamation
        GoTo Cleanup
    End If

    Dim sortedAreas As Collection: Set sortedAreas = SortAreasByWorkload(areas)
    Dim placed As Collection: Set placed = New Collection
    
    ClearPreviousResults ws, lastRow, cols
    PreloadObstacles ws, lastRow, cols, placed

    Dim area As AreaDef
    For Each area In sortedAreas
        Dim finalCandidates As Collection: Set finalCandidates = New Collection
        Dim z As ZoneDef
        For Each z In allZones
            Dim minCX As Double: minCX = z.Left + area.width / 2 + gap_x_mm / 2
            Dim maxCX As Double: maxCX = z.Right - area.width / 2 - gap_x_mm / 2
            Dim minCY As Double: minCY = z.Bottom + area.height / 2 + gap_y_mm / 2
            Dim maxCY As Double: maxCY = z.Top - area.height / 2 - gap_y_mm / 2
            
            If minCX <= maxCX And minCY <= maxCY Then
                Dim bestPointInZone As Variant
                bestPointInZone = FindBestPointInZone_Grid(area, z, placed, inboundX, inboundY, gap_x_mm, gap_y_mm, step_mm)
                
                If Not IsEmpty(bestPointInZone) Then
                    Dim score As Double: score = Sqr((bestPointInZone(0) - inboundX) ^ 2 + (bestPointInZone(1) - inboundY) ^ 2)
                    If Not IsInArray(z.Name, preferredZones) Then score = score + PREFERENCE_PENALTY
                    finalCandidates.Add Array(bestPointInZone(0), bestPointInZone(1), score)
                End If
            End If
        Next z
        
        If finalCandidates.count > 0 Then
            Dim globalBest As Variant: globalBest = finalCandidates(1)
            Dim i As Long
            For i = 2 To finalCandidates.count
                If finalCandidates(i)(2) < globalBest(2) Then globalBest = finalCandidates(i)
            Next i
            
            Dim bestX As Double: bestX = globalBest(0)
            Dim bestY As Double: bestY = globalBest(1)
            
            ws.Cells(area.Row, cols("NewX")).Value = bestX
            ws.Cells(area.Row, cols("NewY")).Value = bestY
            
            Dim rect As Object: Set rect = CreateObject("Scripting.Dictionary")
            rect.Add "CenterX", bestX: rect.Add "CenterY", bestY
            rect.Add "Width", area.width: rect.Add "Height", area.height
            placed.Add rect
            
            ws.Cells(area.Row, cols("NewL")).Value = bestX - area.width / 2
            ws.Cells(area.Row, cols("NewR")).Value = bestX + area.width / 2
            ws.Cells(area.Row, cols("NewB")).Value = bestY - area.height / 2
            ws.Cells(area.Row, cols("NewT")).Value = bestY + area.height / 2
        Else
            ws.Cells(area.Row, cols("NewX")).Value = "#UNPLACED"
        End If
    Next area
    
    CopyStaticItemCoords ws, lastRow, cols

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'========================================================================================
'      HELPER FUNCTIONS
'========================================================================================

Private Sub PreloadObstacles(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal cols As Object, ByRef placed As Collection)
    Dim r As Long
    For r = 2 To lastRow
        Dim layerName As String: layerName = Trim(LCase(CStr(ws.Cells(r, cols("Layer")).Value)))
        Select Case layerName
            Case "area", "areas", "zone", "zones", "walk", "inbound"
            Case Else
                Dim obstacle As Object: Set obstacle = CreateObject("Scripting.Dictionary")
                obstacle.Add "CenterX", CDbl(Nz(ws.Cells(r, cols("OrigX")).Value, 0))
                obstacle.Add "CenterY", CDbl(Nz(ws.Cells(r, cols("OrigY")).Value, 0))
                obstacle.Add "Width", CDbl(Nz(ws.Cells(r, cols("OrigWidth")).Value, 0))
                obstacle.Add "Height", CDbl(Nz(ws.Cells(r, cols("OrigHeight")).Value, 0))
                placed.Add obstacle
        End Select
    Next r
End Sub

Private Sub CopyStaticItemCoords(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal cols As Object)
    Dim r As Long
    Dim layerName As String
    For r = 2 To lastRow
        layerName = Trim(LCase(CStr(ws.Cells(r, cols("Layer")).Value)))
        If Not (layerName Like "area*") Then
            If IsEmpty(ws.Cells(r, cols("NewX")).Value) Or Not IsNumeric(ws.Cells(r, cols("NewX")).Value) Then
                ws.Cells(r, cols("NewX")).Value = ws.Cells(r, cols("OrigX")).Value
                ws.Cells(r, cols("NewY")).Value = ws.Cells(r, cols("OrigY")).Value
                If cols("NewL") > 0 And cols("OrigL") > 0 Then ws.Cells(r, cols("NewL")).Value = ws.Cells(r, cols("OrigL")).Value
                If cols("NewR") > 0 And cols("OrigR") > 0 Then ws.Cells(r, cols("NewR")).Value = ws.Cells(r, cols("OrigR")).Value
                If cols("NewB") > 0 And cols("OrigB") > 0 Then ws.Cells(r, cols("NewB")).Value = ws.Cells(r, cols("OrigB")).Value
                If cols("NewT") > 0 And cols("OrigT") > 0 Then ws.Cells(r, cols("NewT")).Value = ws.Cells(r, cols("OrigT")).Value
            End If
        End If
    Next r
End Sub

Private Function LoadAreas(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal cols As Object, ByVal useNewWidth As Boolean) As Collection
    Dim areas As Collection: Set areas = New Collection
    Dim r As Long, layerName As String, a As AreaDef, widthCol As Long
    If useNewWidth And cols("NewWidth") > 0 Then widthCol = cols("NewWidth") Else widthCol = cols("OrigWidth")
    For r = 2 To lastRow
        layerName = Trim(LCase(CStr(ws.Cells(r, cols("Layer")).Value)))
        If layerName Like "area*" Then
            Set a = New AreaDef
            a.Row = r
            a.workload = CDbl(Nz(ws.Cells(r, cols("Workload")).Value, 0))
            a.width = CDbl(Nz(ws.Cells(r, widthCol).Value, 0))
            a.height = CDbl(Nz(ws.Cells(r, cols("OrigHeight")).Value, 0))
            a.Name = CStr(ws.Cells(r, cols("Text")).Value)
            a.centerX = CDbl(Nz(ws.Cells(r, cols("OrigX")).Value, 0))
            a.centerY = CDbl(Nz(ws.Cells(r, cols("OrigY")).Value, 0))
            areas.Add a
        End If
    Next r
    Set LoadAreas = areas
End Function

Private Function LoadZones(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal cols As Object) As Collection
    Dim zones As Collection: Set zones = New Collection
    Dim r As Long, layerName As String, z As ZoneDef
    For r = 2 To lastRow
        layerName = Trim(LCase(CStr(ws.Cells(r, cols("Layer")).Value)))
        If layerName Like "zone*" Then
            Set z = New ZoneDef
            z.Left = CDbl(Nz(ws.Cells(r, cols("OrigX")).Value, 0)) - CDbl(Nz(ws.Cells(r, cols("OrigWidth")).Value, 0)) / 2
            z.Right = CDbl(Nz(ws.Cells(r, cols("OrigX")).Value, 0)) + CDbl(Nz(ws.Cells(r, cols("OrigWidth")).Value, 0)) / 2
            z.Bottom = CDbl(Nz(ws.Cells(r, cols("OrigY")).Value, 0)) - CDbl(Nz(ws.Cells(r, cols("OrigHeight")).Value, 0)) / 2
            z.Top = CDbl(Nz(ws.Cells(r, cols("OrigY")).Value, 0)) + CDbl(Nz(ws.Cells(r, cols("OrigHeight")).Value, 0)) / 2
            z.Name = CStr(ws.Cells(r, cols("Text")).Value)
            z.centerX = CDbl(Nz(ws.Cells(r, cols("OrigX")).Value, 0))
            z.centerY = CDbl(Nz(ws.Cells(r, cols("OrigY")).Value, 0))
            zones.Add z
        End If
    Next r
    Set LoadZones = zones
End Function

Private Function FindBestPointInZone_Grid(ByVal area As AreaDef, ByVal zc As ZoneDef, ByVal placed As Collection, _
                                          ByVal inboundX As Double, ByVal inboundY As Double, _
                                          ByVal gapX As Double, ByVal gapY As Double, ByVal stepSize As Double) As Variant
    Dim search_minCX As Double: search_minCX = zc.Left + area.width / 2
    Dim search_maxCX As Double: search_maxCX = zc.Right - area.width / 2
    Dim search_minCY As Double: search_minCY = zc.Bottom + area.height / 2
    Dim search_maxCY As Double: search_maxCY = zc.Top - area.height / 2
    Dim gridPoints As Collection: Set gridPoints = New Collection
    Dim gridX As Double, gridY As Double, m As Long, center_y As Double: center_y = (search_minCY + search_maxCY) / 2
    m = 0
    Do
        Dim found_in_pass As Boolean: found_in_pass = False
        gridY = center_y + m * stepSize
        If gridY <= search_maxCY Then
            found_in_pass = True
            For gridX = search_minCX To search_maxCX Step stepSize
                If Not CheckOverlap(gridX, gridY, area.width, area.height, placed, gapX, gapY) Then gridPoints.Add Array(gridX, gridY, Sqr((gridX - inboundX) ^ 2 + (gridY - inboundY) ^ 2))
            Next gridX
        End If
        If m > 0 Then
            gridY = center_y - m * stepSize
            If gridY >= search_minCY Then
                found_in_pass = True
                For gridX = search_minCX To search_maxCX Step stepSize
                    If Not CheckOverlap(gridX, gridY, area.width, area.height, placed, gapX, gapY) Then gridPoints.Add Array(gridX, gridY, Sqr((gridX - inboundX) ^ 2 + (gridY - inboundY) ^ 2))
                Next gridX
            End If
        End If
        If gridPoints.count > 0 Then Exit Do
        If Not found_in_pass Then Exit Do
        m = m + 1
    Loop
    If gridPoints.count > 0 Then
        Dim bestPoint As Variant: bestPoint = gridPoints(1)
        Dim i As Long
        For i = 2 To gridPoints.count
            If gridPoints(i)(2) < bestPoint(2) Then bestPoint = gridPoints(i)
        Next i
        FindBestPointInZone_Grid = bestPoint
    Else
        FindBestPointInZone_Grid = Empty
    End If
End Function

Private Function CheckOverlap(centerX As Double, centerY As Double, width As Double, height As Double, placed As Collection, gapX As Double, gapY As Double) As Boolean
    Dim cLeft As Double: cLeft = centerX - width / 2
    Dim cRight As Double: cRight = centerX + width / 2
    Dim cBottom As Double: cBottom = centerY - height / 2
    Dim cTop As Double: cTop = centerY + height / 2
    Dim rect As Object
    For Each rect In placed
        Dim eLeft As Double, eRight As Double, eBottom As Double, eTop As Double
        eLeft = rect("CenterX") - rect("Width") / 2
        eRight = rect("CenterX") + rect("Width") / 2
        eBottom = rect("CenterY") - rect("Height") / 2
        eTop = rect("CenterY") + rect("Height") / 2
        Dim halfGapX As Double: halfGapX = gapX / 2
        Dim halfGapY As Double: halfGapY = gapY / 2
        If Not ((cRight + halfGapX) <= (eLeft - halfGapX) Or (cLeft - halfGapX) >= (eRight + halfGapX) Or (cTop + halfGapY) <= (eBottom - halfGapY) Or (cBottom - halfGapY) >= (eTop + halfGapY)) Then
            CheckOverlap = True
            Exit Function
        End If
    Next rect
    CheckOverlap = False
End Function

Private Function SortAreasByWorkload(ByVal areas As Collection) As Collection
    Dim sorted As Collection: Set sorted = New Collection
    If areas.count = 0 Then Set SortAreasByWorkload = sorted: Exit Function
    Dim i As Long, j As Long, temp As AreaDef, arr() As AreaDef, item As AreaDef
    ReDim arr(1 To areas.count)
    i = 1
    For Each item In areas
        Set arr(i) = item
        i = i + 1
    Next item
    For i = 2 To UBound(arr)
        Set temp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j).workload < temp.workload Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = temp
    Next i
    For i = 1 To UBound(arr): sorted.Add arr(i): Next i
    Set SortAreasByWorkload = sorted
End Function

Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function

Private Sub ClearPreviousResults(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal cols As Object)
    Dim rangeToClear As Range
    Set rangeToClear = ws.Range(ws.Cells(2, cols("NewX")), ws.Cells(lastRow, cols("NewT")))
    rangeToClear.ClearContents
End Sub

Public Function GetInboundCoord(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal colLayer As Long, ByVal colCoord As Long) As Double
    Dim r As Long
    For r = 2 To lastRow
        If Trim(LCase(CStr(ws.Cells(r, colLayer).Value))) = "inbound" Then
            GetInboundCoord = CDbl(Nz(ws.Cells(r, colCoord).Value, 0))
            Exit Function
        End If
    Next r
    GetInboundCoord = -1
End Function

Private Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim element As Variant
    On Error Resume Next
    For Each element In arr
        If element = stringToBeFound Then IsInArray = True: Exit Function
    Next element
    IsInArray = False
    On Error GoTo 0
End Function

Private Function Nz(ByVal Value As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
    If IsNull(Value) Or IsEmpty(Value) Or IsError(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function

---
**EXCEL - ObjectData.xlsm -  Class Module AreaDef**

'--------------------------------------------
' Class Module: AreaDef
' Description: Represents a movable area to be placed in a zone.
'--------------------------------------------
Option Explicit

Public Row As Long
Public Name As String
Public workload As Double
Public width As Double
Public height As Double
Public centerX As Double
Public centerY As Double

---
**EXCEL - ObjectData.xlsm -  Class Module RectangleDef**

'--------------------------------------------
' Class Module: RectangleDef
'--------------------------------------------
Public Left As Double
Public Right As Double
Public Bottom As Double
Public Top As Double

'--- ADD THESE TWO LINES ---
Public centerX As Double
Public centerY As Double
'---------------------------

Public AreaName As String

---
**EXCEL - ObjectData.xlsm - ClassModule ZoneDef**

'--------------------------------------------
' Class Module: ZoneDef
' Description: Represents a placement zone where areas can be located.
'--------------------------------------------
Option Explicit

Public Name As String
Public Left As Double
Public Right As Double
Public Bottom As Double
Public Top As Double
Public centerX As Double
Public centerY As Double

---
**The .vbs script for automation**

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

---

