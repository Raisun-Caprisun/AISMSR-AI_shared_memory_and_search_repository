Attribute VB_Name = "ImportWorkloadSheetPowerBI"
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

