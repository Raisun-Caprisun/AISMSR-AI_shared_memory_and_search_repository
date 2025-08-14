Attribute VB_Name = "GetWorkloadWidthFromInputData"
'---------------------------------------------------------------------------------------
' Module: DataUpdaters
' Version: 3.5 - Final with Differentiated Data Import
' Description: This version correctly imports data based on layer.
'              - 'Workload' is updated ONLY for rows where Layer = "Areas".
'              - 'New_Width' is updated for ALL rows.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub UpdateFromInputData()
    ' --- Path to the SOURCE file. Using the direct URL. ---
    Const inputDataPath As String = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/InputData.xlsm"
    
    Dim inputWb As Workbook, inputSheet As Worksheet
    Dim objectDataWb As Workbook, layoutSheet As Worksheet
    
    Set objectDataWb = ThisWorkbook
    Set layoutSheet = objectDataWb.Worksheets("Layout")
    If layoutSheet Is Nothing Then
        MsgBox "The active workbook does not contain a 'Layout' sheet.", vbCritical: Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Open the source workbook with a parameter to try and force a refresh ---
    On Error Resume Next
    Set inputWb = Workbooks.Open(Filename:=inputDataPath, ReadOnly:=True, UpdateLinks:=0)
    If inputWb Is Nothing Then
        MsgBox "Failed to open the InputData source file at the specified path:" & vbCrLf & inputDataPath, vbCritical
        GoTo Cleanup
    End If
    Set inputSheet = inputWb.Worksheets(1)
    On Error GoTo 0
    
    ' --- Dynamically find all necessary columns in the source file ---
    Dim colInputText As Long: colInputText = FindHeaderColumn(inputSheet, "Text")
    Dim colInputLayer As Long: colInputLayer = FindHeaderColumn(inputSheet, "Layer")
    Dim colInputWorkload As Long: colInputWorkload = FindHeaderColumn(inputSheet, "Workload")
    Dim colInputNewWidth As Long: colInputNewWidth = FindHeaderColumn(inputSheet, "New_Width")
    
    If colInputText = 0 Or colInputLayer = 0 Or colInputWorkload = 0 Or colInputNewWidth = 0 Then
        MsgBox "Could not find one or more required columns ('Text', 'Layer', 'Workload', 'New_Width') in 'InputData.xlsm'.", vbCritical
        inputWb.Close SaveChanges:=False
        GoTo Cleanup
    End If

    ' *** REVISED: Use two separate dictionaries for separate logic ***
    Dim workloadData As Object: Set workloadData = CreateObject("Scripting.Dictionary")
    workloadData.CompareMode = vbTextCompare
    
    Dim newWidthData As Object: Set newWidthData = CreateObject("Scripting.Dictionary")
    newWidthData.CompareMode = vbTextCompare

    Dim lastInputRow As Long: lastInputRow = inputSheet.Cells(inputSheet.Rows.count, colInputText).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastInputRow
        Dim keyText As String
        keyText = Trim(CStr(inputSheet.Cells(r, colInputText).Value))
        
        If keyText <> "" Then
            ' --- Logic for New_Width (applies to ALL layers) ---
            If Not newWidthData.Exists(keyText) Then
                Dim newWidthValue As Double
                newWidthValue = CDbl(Nz(inputSheet.Cells(r, colInputNewWidth).Value))
                newWidthData.Add keyText, newWidthValue
            End If

            ' --- Logic for Workload (applies ONLY to "Areas") ---
            Dim layerName As String
            layerName = Trim(LCase(CStr(inputSheet.Cells(r, colInputLayer).Value)))
            
            If layerName Like "area*" Then
                If Not workloadData.Exists(keyText) Then
                    Dim workloadValue As Double
                    workloadValue = CDbl(Nz(inputSheet.Cells(r, colInputWorkload).Value))
                    workloadData.Add keyText, workloadValue
                End If
            End If
        End If
    Next r
    
    inputWb.Close SaveChanges:=False
    
    ' --- Find destination columns ---
    Dim colDestText As Long: colDestText = FindHeaderColumn(layoutSheet, "Text")
    Dim colDestWorkload As Long: colDestWorkload = FindHeaderColumn(layoutSheet, "Workload")
    Dim colDestNewWidth As Long: colDestNewWidth = FindHeaderColumn(layoutSheet, "New_Width")
    
    If colDestText = 0 Or colDestWorkload = 0 Or colDestNewWidth = 0 Then
        MsgBox "Could not find 'Text', 'Workload', or 'New_Width' columns in this workbook's 'Layout' sheet.", vbCritical
        GoTo Cleanup
    End If
    
    ' --- Update columns in this sheet using the separate dictionaries ---
    Dim lastLayoutRow As Long: lastLayoutRow = layoutSheet.Cells(layoutSheet.Rows.count, colDestText).End(xlUp).Row
    
    Dim updatedCount As Long: updatedCount = 0
    For r = 2 To lastLayoutRow
        Dim layoutKeyText As String
        layoutKeyText = Trim(CStr(layoutSheet.Cells(r, colDestText).Value))
        
        If layoutKeyText <> "" Then
            Dim didUpdate As Boolean: didUpdate = False

            ' Update Workload (only if it exists in the workload dictionary)
            If workloadData.Exists(layoutKeyText) Then
                layoutSheet.Cells(r, colDestWorkload).Value = workloadData(layoutKeyText)
                didUpdate = True
            Else
                layoutSheet.Cells(r, colDestWorkload).ClearContents
            End If

            ' Update New_Width (if it exists in the new width dictionary)
            If newWidthData.Exists(layoutKeyText) Then
                layoutSheet.Cells(r, colDestNewWidth).Value = newWidthData(layoutKeyText)
                didUpdate = True
            End If
            
            If didUpdate Then updatedCount = updatedCount + 1
        End If
    Next r
    
    objectDataWb.Save
    MsgBox "Update complete." & vbCrLf & vbCrLf & updatedCount & " rows were updated from the InputData file.", vbInformation

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
    ' Helper function to handle Null, Empty, or Error values gracefully.
    If IsNull(Value) Or IsEmpty(Value) Or IsError(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function
