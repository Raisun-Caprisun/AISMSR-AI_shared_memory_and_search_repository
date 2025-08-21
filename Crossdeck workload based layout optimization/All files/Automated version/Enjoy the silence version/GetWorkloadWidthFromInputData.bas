Attribute VB_Name = "GetWorkloadWidthFromInputData"
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
