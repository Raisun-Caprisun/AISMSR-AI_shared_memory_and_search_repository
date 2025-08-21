Attribute VB_Name = "MatrixAnalyzers"
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
