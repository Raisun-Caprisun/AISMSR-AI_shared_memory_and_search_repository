Attribute VB_Name = "LayoutCostCalculator"
'---------------------------------------------------------------------------------------
' Module: LayoutCostCalculator
' Version: 5.0 - Final with Total Weighted Distance
' Popis: Vypoèítá celkové "náklady", prùmìrnou váženou vzdálenost a celkovou váženou vzdálenost.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub CalculateAllLayoutCosts()
    Dim layoutSheet As Worksheet, costSheet As Worksheet
    
    ' --- Nastavení ---
    On Error Resume Next
    Set layoutSheet = ThisWorkbook.Worksheets("Layout")
    If layoutSheet Is Nothing Then
        MsgBox "List 'Layout' nebyl nalezen.", vbCritical: Exit Sub
    End If
    
    Set costSheet = ThisWorkbook.Worksheets("Cost_Calculation")
    If costSheet Is Nothing Then
        Set costSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    End If
    On Error GoTo 0
    
    costSheet.Name = "Cost_Calculation"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    costSheet.Cells.Clear
    
    ' --- Výpoèet všech ètyø scénáøù ---
    Dim results(1 To 4, 1 To 3) As Double ' Array to hold Cost, AvgDist, TotalDist for each of 4 scenarios
    Dim resultData As Variant
    
    ' Scénáø 1: Default, Euclidean
    resultData = CalculateSingleCost(layoutSheet, isOptimized:=False, useEuclidean:=True)
    results(1, 1) = resultData(0): results(1, 2) = resultData(1) / 1000: results(1, 3) = resultData(2) / 1000000
    
    ' Scénáø 2: Default, Manhattan
    resultData = CalculateSingleCost(layoutSheet, isOptimized:=False, useEuclidean:=False)
    results(2, 1) = resultData(0): results(2, 2) = resultData(1) / 1000: results(2, 3) = resultData(2) / 1000000
    
    ' Scénáø 3: Adjusted, Euclidean
    resultData = CalculateSingleCost(layoutSheet, isOptimized:=True, useEuclidean:=True)
    results(3, 1) = resultData(0): results(3, 2) = resultData(1) / 1000: results(3, 3) = resultData(2) / 1000000

    ' Scénáø 4: Adjusted, Manhattan
    resultData = CalculateSingleCost(layoutSheet, isOptimized:=True, useEuclidean:=False)
    results(4, 1) = resultData(0): results(4, 2) = resultData(1) / 1000: results(4, 3) = resultData(2) / 1000000

    ' --- Zápis výsledkù na list ---
    With costSheet
        .Range("A1:D1").Value = Array("Method", "Total Cost (Workload*Distance)", "Average Distance per Workload Unit (m)", "Total Weighted Distance (km)")
        
        .Range("A2").Value = "Total cost default EUC"
        .Range("A3").Value = "Total cost default MAN"
        .Range("A4").Value = "Total cost adjusted sorting EUC"
        .Range("A5").Value = "Total cost adjusted sorting MAN"
        
        .Range("B2").Value = results(1, 1)
        .Range("C2").Value = results(1, 2)
        .Range("D2").Value = results(1, 3)
        
        .Range("B3").Value = results(2, 1)
        .Range("C3").Value = results(2, 2)
        .Range("D3").Value = results(2, 3)
        
        .Range("B4").Value = results(3, 1)
        .Range("C4").Value = results(3, 2)
        .Range("D4").Value = results(3, 3)
        
        .Range("B5").Value = results(4, 1)
        .Range("C5").Value = results(4, 2)
        .Range("D5").Value = results(4, 3)
        
        ' Formátování
        .Columns("A:D").AutoFit
        .Range("A1:D1").Font.Bold = True
        .Range("A2:A5").Font.Bold = True
        .Columns("B:D").NumberFormat = "0.00"
        .Activate
        .Cells(1, 1).Select
    End With

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Finální analýza nákladù a vzdáleností byla dokonèena.", vbInformation
End Sub

Private Function CalculateSingleCost(ByVal layoutSheet As Worksheet, ByVal isOptimized As Boolean, ByVal useEuclidean As Boolean) As Variant
    Dim r As Long
    Dim totalCost As Double: totalCost = 0
    Dim totalWorkload As Double: totalWorkload = 0
    Dim totalWorkloadDistance As Double: totalWorkloadDistance = 0 ' Nová promìnná

    ' --- Najdi sloupce ---
    Dim colLayer As Long: colLayer = FindHeaderColumn(layoutSheet, "Layer")
    Dim colWorkload As Long: colWorkload = FindHeaderColumn(layoutSheet, "Workload")
    Dim colX As Long, colY As Long
    
    If isOptimized Then
        colX = FindHeaderColumn(layoutSheet, "New_Center_X")
        colY = FindHeaderColumn(layoutSheet, "New_Center_Y")
    Else
        colX = FindHeaderColumn(layoutSheet, "CenterX")
        colY = FindHeaderColumn(layoutSheet, "CenterY")
    End If
    
    ' --- Naètení souøadnic Inbound ---
    Dim lastRow As Long: lastRow = layoutSheet.Cells(layoutSheet.Rows.count, "A").End(xlUp).Row
    Dim inboundX As Double: inboundX = GetInboundCoord(layoutSheet, lastRow, colLayer, colX)
    Dim inboundY As Double: inboundY = GetInboundCoord(layoutSheet, lastRow, colLayer, colY)
    
    If inboundX = -1 Or colWorkload = 0 Then
        CalculateSingleCost = Array(0, 0, 0)
        Exit Function
    End If
    
    ' --- Výpoèet nákladù a vzdáleností ---
    For r = 2 To lastRow
        If Trim(LCase(CStr(layoutSheet.Cells(r, colLayer).Value))) Like "area*" Then
            Dim workload As Double: workload = val(layoutSheet.Cells(r, colWorkload).Value)
            If IsNumeric(layoutSheet.Cells(r, colX).Value) Then
                totalWorkload = totalWorkload + workload
                
                If workload > 0 Then
                    Dim areaX As Double: areaX = CDbl(layoutSheet.Cells(r, colX).Value)
                    Dim areaY As Double: areaY = CDbl(layoutSheet.Cells(r, colY).Value)
                    Dim distance As Double
                    
                    If useEuclidean Then
                        distance = Sqr((areaX - inboundX) ^ 2 + (areaY - inboundY) ^ 2)
                    Else
                        distance = Abs(areaX - inboundX) + Abs(areaY - inboundY)
                    End If
                    
                    totalCost = totalCost + (workload * distance)
                    totalWorkloadDistance = totalWorkloadDistance + (workload * distance) ' Agregace vážené vzdálenosti
                End If
            End If
        End If
    Next r
    
    Dim avgWeightedDistance As Double
    If totalWorkload > 0 Then
        avgWeightedDistance = totalCost / totalWorkload
    Else
        avgWeightedDistance = 0
    End If
    
    ' Vrací pole se tøemi hodnotami
    CalculateSingleCost = Array(totalCost, avgWeightedDistance, totalWorkloadDistance)
End Function


' Tyto dvì funkce musí být VEØEJNÉ v jiném modulu (napø. v modulu "Sorting")
' Public Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
' Public Function GetInboundCoord(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal colLayer As Long, ByVal colCoord As Long) As Double
