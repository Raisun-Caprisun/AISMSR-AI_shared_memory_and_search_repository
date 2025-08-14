Attribute VB_Name = "ImportWorkloadSheetPowerBI"
' =========================================================================================
'
'         FIN¡LNÕ SKRIPTY PRO IMPORT A AKTUALIZACI LAYOUTU Z EXCELU
'         Verze: 13.1 (DefinitivnÌ verze - Oprava chyby "Type Mismatch")
'
' =========================================================================================
' POPIS ZMÃNY:
' Tato verze opravuje chybu "Type Mismatch" a vracÌ se k robustnÌmu ËtenÌ dat.
' 1. Vöechny hodnoty z Excelu se nejprve naËÌtajÌ do flexibilnÌch 'Variant' promÏnn˝ch.
' 2. Logika If/Else pro v˝bÏr sou¯adnic nynÌ pouze urËuje, ze kter˝ch sloupc˘ se m· ËÌst.
' 3. St·vajÌcÌ kontrola IsNumeric zajiöùuje, ûe se kreslÌ pouze tvary s platn˝mi ËÌseln˝mi daty.
' =========================================================================================
Public Sub ImportLayout_KROK_1_NakreslitVse()
    ' --- PARAMETER FOR ZOOM ---
    Const finalZoomLevel As Double = 0.25 ' 0.25 = 25%

    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim targetPage As Visio.Page
    Dim visShape As Visio.Shape
    Dim inboundShape As Visio.Shape
    Dim filePath As String
    Dim lastRow As Long, i As Long
    
    ' Restored original SharePoint path as requested
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/ObjectData.xlsm"
    Const MM_NA_PALEC As Double = 1 / 25.4

    Set targetPage = Application.ActiveDocument.Pages(1)
    
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then Exit Sub
    
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(filePath)
    If xlWb Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
        Exit Sub
    End If
    
    Set ws = xlWb.Worksheets("Layout")
    
    For i = targetPage.Shapes.Count To 1 Step -1: targetPage.Shapes(i).Delete: Next i

    lastRow = ws.UsedRange.Rows.Count
    If lastRow <= 1 Then GoTo Cleanup
    
    ' --- Main loop for data import ---
    For i = 2 To lastRow
        ' *** RESTORED: Read all data into flexible Variant variables first for safety ***
        Dim objID_val As Variant, shapeText As Variant, layerName As Variant
        Dim centerX_val As Variant, centerY_val As Variant, width_val As Variant, height_val As Variant
        Dim angle_deg_val As Variant, color_rgb_val As Variant
        
        ' --- Read common data that doesn't change ---
        objID_val = ws.Cells(i, "A").Value
        shapeText = ws.Cells(i, "C").Value
        layerName = ws.Cells(i, "D").Value
        color_rgb_val = ws.Cells(i, "E").Value
        height_val = ws.Cells(i, "I").Value
        angle_deg_val = ws.Cells(i, "J").Value

        ' *** CRITICAL FIX: Smartly decide which source columns to read from ***
        If Trim(LCase(CStr(layerName))) Like "area*" Then
            ' For "Areas", read from the NEW, optimized columns
            width_val = ws.Cells(i, "Q").Value    ' New_Width
            centerX_val = ws.Cells(i, "R").Value  ' New_Center_X
            centerY_val = ws.Cells(i, "S").Value  ' New_Center_Y
        Else
            ' For ALL OTHER layers (Zones, Walls, etc.), read from the ORIGINAL, static columns
            width_val = ws.Cells(i, "H").Value    ' Width
            centerX_val = ws.Cells(i, "F").Value  ' CenterX
            centerY_val = ws.Cells(i, "G").Value  ' CenterY
        End If
        
        ' *** RESTORED: Now, safely check if the chosen values are numeric before proceeding ***
        If IsNumeric(centerX_val) And IsNumeric(centerY_val) And IsNumeric(width_val) And IsNumeric(height_val) Then
            
            ' --- Convert safe values to specific types for calculation ---
            Dim centerX_mm As Double, centerY_mm As Double, width_mm As Double, height_mm As Double
            centerX_mm = CDbl(centerX_val)
            centerY_mm = CDbl(centerY_val)
            width_mm = CDbl(width_val)
            height_mm = CDbl(height_val)
            
            Dim centerX_inch As Double, centerY_inch As Double, width_inch As Double, height_inch As Double
            centerX_inch = centerX_mm * MM_NA_PALEC
            centerY_inch = centerY_mm * MM_NA_PALEC
            width_inch = width_mm * MM_NA_PALEC
            height_inch = height_mm * MM_NA_PALEC

            Set visShape = targetPage.DrawRectangle(centerX_inch - (width_inch / 2), centerY_inch - (height_inch / 2), centerX_inch + (width_inch / 2), centerY_inch + (height_inch / 2))
            
            visShape.Text = CStr(shapeText)
            visShape.CellsU("Angle").Result("deg") = CDbl(Nz(angle_deg_val))
            visShape.CellsU("FillForegnd").Result(visColor) = CLng(Nz(color_rgb_val))
            visShape.CellsU("Char.Size").Result("pt") = 36
            
            If LCase(CStr(shapeText)) = "inbound" Then Set inboundShape = visShape

            If Not IsEmpty(objID_val) And Trim(CStr(objID_val)) <> "" Then
                visShape.AddNamedRow visSectionProp, "objID", visTagDefault
                visShape.CellsU("Prop.objID.Label").FormulaU = """objID"""
                If IsNumeric(objID_val) Then visShape.CellsU("Prop.objID.Value").FormulaU = objID_val Else visShape.CellsU("Prop.objID.Value").FormulaU = """" & Replace(objID_val, """", """""") & """"
            End If
            
            If Trim(CStr(layerName)) <> "" Then
                Dim visLayer As Visio.Layer
                On Error Resume Next
                Set visLayer = targetPage.Layers.ItemU(CStr(layerName))
                On Error GoTo 0
                If visLayer Is Nothing Then Set visLayer = targetPage.Layers.Add(CStr(layerName))
                visLayer.Add visShape, 1
                Set visLayer = Nothing
            End If
            
            visShape.BringToFront
        End If
    Next i
    
    ' Hide the "Zones" layer after drawing everything
    On Error Resume Next
    Dim zonesLayer As Visio.Layer
    Set zonesLayer = targetPage.Layers.ItemU("Zones")
    If Not zonesLayer Is Nothing Then
        zonesLayer.CellsC(visLayerVisible).ResultIU = 0
    End If
    On Error GoTo 0
    
    ' Center view and set zoom level
    If Not inboundShape Is Nothing Then
        Application.ActiveWindow.CenterViewOnShape inboundShape, visCenterView
        Application.ActiveWindow.Zoom = finalZoomLevel
    End If

Cleanup:
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    On Error GoTo 0
    Set ws = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    Set visShape = Nothing: Set targetPage = Nothing: Set inboundShape = Nothing
End Sub

Private Function Nz(ByVal Value As Variant, Optional ByVal ValueIfNull As Variant = 0) As Variant
    ' Helper function included for completeness
    If IsNull(Value) Or IsEmpty(Value) Or IsError(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function
