Attribute VB_Name = "ExportDefault"
' =========================================================================================
'
'         FINÁLNÍ SKRIPTY PRO EXPORT LAYOUTU Z VISIA DO EXCELU
'         Verze: 9.0 (Silent Version with Expanded InputData Export)
'
' =========================================================================================
' POPIS ZMÌNY:
' Verze 9.0 rozšiøuje export do souboru "InputData.xlsm".
' 1. Pøidává sloupce "ID" a "Layer" pro lepší možnosti tøídìní a analýzy.
' 2. Mìní poøadí sloupcù na: ID, Text, Layer, Workload, New_Width.
' 3. Pøidává nový, prázdný sloupec "Max_Buffer" pro budoucí výpoèty.
' =========================================================================================

Public Sub ExportLayoutuDoExcelu_Finalni_s_Dokumentaci()
    ' --- Declarations for Excel and Visio objects ---
    Dim xlApp As Object
    Dim xlWbMain As Object, wsMain As Object
    Dim xlWbInput As Object, wsInput As Object ' Renamed for clarity
    Dim pvWindow As Object
    Dim visShape As Visio.Shape
    Dim layoutPage As Visio.Page
    
    ' --- File Paths ---
    Dim mainFilePath As String
    Dim inputDataFilePath As String ' Renamed for clarity
    mainFilePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/ObjectData.xlsm"
    inputDataFilePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/InputData.xlsm"

    ' --- Header Definitions ---
    Dim mainHeaders As Variant
    Dim inputDataHeaders As Variant
    
    ' Headers for the main file (ObjectData.xlsm) - unchanged
    mainHeaders = Array("ID", "Name", "Text", "Layer", "Color (RGB)", "CenterX", "CenterY", "Width", "Height", "Angle", "Z-Order", "BBox_Left_X", "BBox_Right_X", "BBox_Bottom_Y", "BBox_Top_Y", "Workload", "New_Width", "New_Center_X", "New_Center_Y", "New_BBox_Left_X", "New_BBox_Right_X", "New_BBox_Bottom_Y", "New_BBox_Top_Y")
    
    ' *** REVISED: New headers for InputData.xlsm in the correct order ***
    inputDataHeaders = Array("ID", "Text", "Layer", "Workload", "New_Width", "Max_Buffer")

    ' --- Other variables ---
    Dim i As Long
    Dim bBoxLeftInches As Double, bBoxBottomInches As Double, bBoxRightInches As Double, bBoxTopInches As Double
    Dim shapeWidth As Double
    Const PALEC_NA_MM As Double = 25.4
    
    ' --- Robustly find the "Layout" page ---
    Dim p As Visio.Page
    For Each p In ThisDocument.Pages
        If LCase(p.Name) = "layout" Or LCase(p.NameU) = "layout" Then
            Set layoutPage = p
            Exit For
        End If
    Next p
    
    If layoutPage Is Nothing Then
        Debug.Print "ERROR: Page named 'Layout' was not found in this Visio document. Export cannot continue."
        Exit Sub
    End If
    
    If Not Application.ActiveWindow.Page Is layoutPage Then
        Application.ActiveWindow.Page = layoutPage
    End If
    
    ' --- Connect to Excel ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then Debug.Print "Could not start Excel.": Exit Sub
    
    xlApp.Visible = True

    ' --- Process ObjectData.xlsm ---
    On Error Resume Next
    Set xlWbMain = xlApp.Workbooks.Open(mainFilePath)
    On Error GoTo 0
    If xlWbMain Is Nothing Then
        Debug.Print "ERROR: Could not open the main file: '" & mainFilePath & "'."
        xlApp.Quit: Set xlApp = Nothing: Exit Sub
    End If
    
    ' --- Process InputData.xlsm ---
    On Error Resume Next
    Set xlWbInput = xlApp.Workbooks.Open(inputDataFilePath)
    On Error GoTo 0
    If xlWbInput Is Nothing Then
        Debug.Print "ERROR: Could not open the InputData file: '" & inputDataFilePath & "'."
        xlWbMain.Close SaveChanges:=False
        xlApp.Quit: Set xlApp = Nothing: Exit Sub
    End If

    ' --- Handle Protected View ---
    If xlApp.ProtectedViewWindows.Count > 0 Then
        For Each pvWindow In xlApp.ProtectedViewWindows
            If pvWindow.Workbook.FullName = xlWbMain.FullName Or pvWindow.Workbook.FullName = xlWbInput.FullName Then
                pvWindow.Edit
            End If
        Next
        xlApp.Wait (Now + TimeValue("0:00:02"))
    End If
    
    ' --- Prepare worksheets ---
    Set wsMain = xlWbMain.Worksheets("Layout")
    Set wsInput = xlWbInput.Worksheets(1)
    
    ' Prepare main sheet
    wsMain.Activate
    Dim lastColMain As Integer: lastColMain = UBound(mainHeaders) + 1
    wsMain.Range("A2:" & wsMain.Cells(50000, lastColMain).Address).ClearContents
    wsMain.Range("A1").Resize(1, lastColMain).Value = mainHeaders
    wsMain.Range("P1:Q1").Interior.Color = RGB(146, 208, 80)
    wsMain.Range("R1:W1").Interior.Color = RGB(155, 194, 230)
    
    ' Prepare input data sheet
    wsInput.Activate
    Dim lastColInput As Integer: lastColInput = UBound(inputDataHeaders) + 1
    wsInput.Cells.ClearContents
    wsInput.Range("A1").Resize(1, lastColInput).Value = inputDataHeaders

    ' --- Main loop to export data to both files ---
    i = 2
    For Each visShape In ActivePage.Shapes
        visShape.BoundingBox visBBoxUprightWH, bBoxLeftInches, bBoxBottomInches, bBoxRightInches, bBoxTopInches
        shapeWidth = visShape.CellsU("Width").Result("mm")

        ' --- Write data to ObjectData.xlsm ---
        Dim currentID As Variant
        On Error Resume Next
        currentID = visShape.CellsU("Prop.objID").ResultIU
        wsMain.Cells(i, "A").Value = currentID
        wsMain.Cells(i, "C").Value = visShape.Text
        wsMain.Cells(i, "D").Value = visShape.Layer(1).Name
        On Error GoTo 0
        
        wsMain.Cells(i, "B").Value = visShape.Name
        wsMain.Cells(i, "E").Value = visShape.CellsU("FillForegnd").Result(visColor)
        wsMain.Cells(i, "F").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Result("mm")
        wsMain.Cells(i, "G").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Result("mm")
        wsMain.Cells(i, "H").Value = shapeWidth
        wsMain.Cells(i, "I").Value = visShape.CellsU("Height").Result("mm")
        wsMain.Cells(i, "J").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormAngle).Result("deg")
        wsMain.Cells(i, "K").Value = i - 1 ' Z-Order
        wsMain.Cells(i, "L").Value = bBoxLeftInches * PALEC_NA_MM
        wsMain.Cells(i, "M").Value = bBoxRightInches * PALEC_NA_MM
        wsMain.Cells(i, "N").Value = bBoxBottomInches * PALEC_NA_MM
        wsMain.Cells(i, "O").Value = bBoxTopInches * PALEC_NA_MM
        
        ' *** REVISED: Write expanded data to InputData.xlsm in the new order ***
        wsInput.Cells(i, "A").Value = currentID ' Column A: ID
        wsInput.Cells(i, "B").Value = visShape.Text ' Column B: Text
        On Error Resume Next
        wsInput.Cells(i, "C").Value = visShape.Layer(1).Name ' Column C: Layer
        On Error GoTo 0
        ' Column D ("Workload") remains blank
        wsInput.Cells(i, "E").Value = shapeWidth ' Column E: New_Width
        ' Column F ("Max_Buffer") remains blank
        
        i = i + 1
    Next visShape
    
    ' --- Finalize and clean up ---
    wsMain.Columns("A:" & Split(wsMain.Cells(1, lastColMain).Address, "$")(1)).AutoFit
    xlWbMain.Save
    
    wsInput.Columns("A:" & Split(wsInput.Cells(1, lastColInput).Address, "$")(1)).AutoFit
    xlWbInput.Save
    
    xlWbMain.Close SaveChanges:=False
    xlWbInput.Close SaveChanges:=False
    
    Set pvWindow = Nothing
    Set wsMain = Nothing
    Set wsInput = Nothing
    Set xlWbMain = Nothing
    Set xlWbInput = Nothing
    Set xlApp = Nothing
    
    Debug.Print "Export to both files completed successfully with the new InputData structure."

End Sub
