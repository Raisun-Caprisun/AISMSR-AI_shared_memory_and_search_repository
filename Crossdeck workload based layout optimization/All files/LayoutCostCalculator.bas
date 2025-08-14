Attribute VB_Name = "LayoutCostCalculator"
'---------------------------------------------------------------------------------------
' Module: LayoutCostCalculator
' Version: 11.1 - Final with Improved Header Clarity
' Description: This version refines the headers in the final report to be more
'              intuitive and business-focused for end users and management.
'---------------------------------------------------------------------------------------
Option Explicit

'========================================================================================
'      PUBLIC MASTER MACRO - This is the recommended macro to run.
'========================================================================================

Public Sub RunFullAnalysis()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Run both analysis components
    CalculateAllLayoutCosts showMsg:=False
    ' GenerateAnalysisDashboard showMsg:=False ' Temporarily disabled as per user request

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "The full analysis, including the Cost Calculation, is complete.", vbInformation
End Sub


'========================================================================================
'      CORE CALCULATION LOGIC
'========================================================================================

Public Sub CalculateAllLayoutCosts(Optional ByVal showMsg As Boolean = True)
    ' --- This is the scale factor from your Visio drawing (10mm = 1m) ---
    Const SCALE_FACTOR_MM_PER_METER As Double = 10

    Dim layoutSheet As Worksheet, costSheet As Worksheet

    ' --- Setup and Sheet Preparation ---
    On Error Resume Next
    Set layoutSheet = ThisWorkbook.Worksheets("Layout")
    If layoutSheet Is Nothing Then
        MsgBox "The 'Layout' worksheet was not found. Cannot proceed.", vbCritical: Exit Sub
    End If

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Cost_Calculation").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set costSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    costSheet.Name = "Cost_Calculation"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- Calculate all four scenarios ---
    ' Array is sized for 4 scenarios and 6 result columns
    Dim results(1 To 4, 1 To 6) As Double
    Dim resultData As Variant

    ' --- Scenario 1: Default, Euclidean ---
    resultData = CalculateSingleCost(layoutSheet, False, True, SCALE_FACTOR_MM_PER_METER)
    results(1, 1) = resultData(0) ' One-Way Cost
    results(1, 2) = resultData(1) ' Avg Distance (m)
    results(1, 3) = resultData(2) / 1000 ' Total Distance (km)
    results(1, 4) = resultData(0) * 2 ' Round-Trip Cost
    results(1, 5) = resultData(1) * 2 ' Round-Trip Avg Distance (m)
    results(1, 6) = (resultData(2) / 1000) * 2 ' Round-Trip Total Distance (km)
    
    ' --- Scenario 2: Default, Manhattan ---
    resultData = CalculateSingleCost(layoutSheet, False, False, SCALE_FACTOR_MM_PER_METER)
    results(2, 1) = resultData(0): results(2, 2) = resultData(1): results(2, 3) = resultData(2) / 1000
    results(2, 4) = resultData(0) * 2: results(2, 5) = resultData(1) * 2: results(2, 6) = (resultData(2) / 1000) * 2
    
    ' --- Scenario 3: Optimized, Euclidean ---
    resultData = CalculateSingleCost(layoutSheet, True, True, SCALE_FACTOR_MM_PER_METER)
    results(3, 1) = resultData(0): results(3, 2) = resultData(1): results(3, 3) = resultData(2) / 1000
    results(3, 4) = resultData(0) * 2: results(3, 5) = resultData(1) * 2: results(3, 6) = (resultData(2) / 1000) * 2

    ' --- Scenario 4: Optimized, Manhattan ---
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

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "The Cost Calculation analysis is complete.", vbInformation
End Sub

Private Function CalculateSingleCost(ByVal layoutSheet As Worksheet, ByVal isOptimized As Boolean, ByVal useEuclidean As Boolean, ByVal scaleFactor As Double) As Variant
    Dim r As Long, lastRow As Long
    Dim totalCost As Double, totalWorkload As Double, totalWorkloadDistance As Double
    Dim colLayer As Long, colWorkload As Long, colX As Long, colY As Long
    
    ' --- Find necessary columns ---
    colLayer = FindHeaderColumn(layoutSheet, "Layer")
    colWorkload = FindHeaderColumn(layoutSheet, "Workload")
    
    If isOptimized Then
        colX = FindHeaderColumn(layoutSheet, "New_Center_X")
        colY = FindHeaderColumn(layoutSheet, "New_Center_Y")
    Else
        colX = FindHeaderColumn(layoutSheet, "CenterX")
        colY = FindHeaderColumn(layoutSheet, "CenterY")
    End If
    
    lastRow = layoutSheet.Cells(layoutSheet.Rows.count, "A").End(xlUp).Row
    
    ' --- Get Inbound coordinates ---
    Dim inboundX As Double, inboundY As Double
    inboundX = GetInboundCoord(layoutSheet, lastRow, colLayer, colX)
    inboundY = GetInboundCoord(layoutSheet, lastRow, colLayer, colY)
    
    ' --- Exit if essential data is missing ---
    If inboundX = -1 Or colWorkload = 0 Or scaleFactor = 0 Then
        CalculateSingleCost = Array(0, 0, 0)
        Exit Function
    End If
    
    ' --- Loop through all areas and calculate costs ---
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
                If useEuclidean Then
                    distance = Sqr((areaX - inboundX) ^ 2 + (areaY - inboundY) ^ 2)
                Else
                    distance = Abs(areaX - inboundX) + Abs(areaY - inboundY)
                End If
                
                realDistance = distance / scaleFactor
                totalCost = totalCost + (workload * realDistance)
                totalWorkloadDistance = totalWorkloadDistance + (workload * realDistance)
            End If
        End If
    Next r
    
    ' --- Calculate the final average distance ---
    Dim avgWeightedDistance As Double
    If totalWorkload > 0 Then
        avgWeightedDistance = totalCost / totalWorkload
    Else
        avgWeightedDistance = 0
    End If
    
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
