Attribute VB_Name = "Sorting"
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

