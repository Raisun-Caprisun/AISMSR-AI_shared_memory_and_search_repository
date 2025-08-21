Attribute VB_Name = "ExportDefault"
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
