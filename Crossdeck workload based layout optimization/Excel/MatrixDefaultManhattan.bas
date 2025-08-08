Attribute VB_Name = "MatrixDefaultManhattan"
'---------------------------------------------------------------------------------------
' Module: MatrixGeneratorDefault_MAN
' Version: 1.0
' Popis: Vytvoøí Metrickou matici "po osách" z PÙVODNÍCH souøadnic na 5. listu.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub GenerateDefaultMatrix_Manhattan()
    Dim layoutSheet As Worksheet, matrixSheet As Worksheet
    
    On Error Resume Next
    Set layoutSheet = ThisWorkbook.Worksheets("Layout")
    If layoutSheet Is Nothing Then
        MsgBox "List 'Layout' nebyl nalezen.", vbCritical: Exit Sub
    End If
    
    ' Cílení na pátý list
    Do While ThisWorkbook.Worksheets.count < 5
        ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    Loop
    Set matrixSheet = ThisWorkbook.Worksheets(5)
    On Error GoTo 0
    
    matrixSheet.Name = "Matrix_Manhattan_Default"
    
    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    
    ' Cílení na PÙVODNÍ sloupce souøadnic
    Dim colId As Long:      colId = FindHeaderColumn(layoutSheet, "ID")
    Dim colLayer As Long:   colLayer = FindHeaderColumn(layoutSheet, "Layer")
    Dim colOrigX As Long:   colOrigX = FindHeaderColumn(layoutSheet, "CenterX")
    Dim colOrigY As Long:   colOrigY = FindHeaderColumn(layoutSheet, "CenterY")
    
    If colId = 0 Or colLayer = 0 Or colOrigX = 0 Or colOrigY = 0 Then
        MsgBox "Nebyly nalezeny povinné sloupce (ID, Layer, CenterX, CenterY).", vbCritical: GoTo Cleanup
    End If
    
    Dim objects As Collection: Set objects = LoadFilteredObjects(layoutSheet, colId, colLayer, colOrigX, colOrigY)
    If objects.count = 0 Then MsgBox "Nebyly nalezeny žádné objekty s platnými souøadnicemi.", vbInformation: GoTo Cleanup
    
    Dim sortedObjects As Collection: Set sortedObjects = SortObjectsByID(objects)
    
    matrixSheet.Cells.Clear
    Dim headerOffset As Integer: headerOffset = 2
    
    Dim i As Long, obj As Object
    For i = 1 To sortedObjects.count
        Set obj = sortedObjects(i)
        matrixSheet.Cells(i + headerOffset - 1, 1).Value = obj("ID")
        matrixSheet.Cells(1, i + headerOffset - 1).Value = obj("ID")
    Next i
    
    Dim j As Long, obj1 As Object, obj2 As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double, distance As Double
    For i = 1 To sortedObjects.count
        Set obj1 = sortedObjects(i): x1 = obj1("X"): y1 = obj1("Y")
        For j = 1 To sortedObjects.count
            Set obj2 = sortedObjects(j): x2 = obj2("X"): y2 = obj2("Y")
            distance = Abs(x2 - x1) + Abs(y2 - y1) ' Manhattan
            matrixSheet.Cells(i + headerOffset - 1, j + headerOffset - 1).Value = distance
        Next j
    Next i
    
    FormatMatrixSheet matrixSheet, sortedObjects.count
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
    MsgBox "Metrická matice PÙVODNÍHO layoutu 'po osách' byla úspìšnì vytvoøena.", vbInformation
End Sub


Private Function LoadFilteredObjects(ByVal ws As Worksheet, ByVal colId As Long, ByVal colLayer As Long, ByVal colX As Long, ByVal colY As Long) As Collection
    Dim objects As Collection: Set objects = New Collection
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        Dim layerName As String: layerName = Trim(LCase(CStr(ws.Cells(i, colLayer).Value)))
        If (layerName = "inbound" Or layerName Like "area*") Then
            If IsNumeric(ws.Cells(i, colX).Value) And IsNumeric(ws.Cells(i, colY).Value) Then
                Dim objData As Object: Set objData = CreateObject("Scripting.Dictionary")
                objData.Add "ID", ws.Cells(i, colId).Value
                objData.Add "X", CDbl(ws.Cells(i, colX).Value)
                objData.Add "Y", CDbl(ws.Cells(i, colY).Value)
                objects.Add objData
            End If
        End If
    Next i
    Set LoadFilteredObjects = objects
End Function

Private Sub FormatMatrixSheet(ByVal matrixSheet As Worksheet, ByVal objectCount As Long)
    matrixSheet.Columns("A:A").AutoFit
    matrixSheet.Rows("1:1").Font.Bold = True
    matrixSheet.Columns("A:A").Font.Bold = True
    matrixSheet.Range(matrixSheet.Cells(2, 2), matrixSheet.Cells(objectCount + 1, objectCount + 1)).NumberFormat = "0"
    matrixSheet.Activate
    matrixSheet.Cells(1, 1).Select
End Sub

Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function

Private Function SortObjectsByID(ByVal coll As Collection) As Collection
    If coll.count <= 1 Then Set SortObjectsByID = coll: Exit Function
    Dim i As Long, j As Long, temp As Object, arr As New Collection
    Dim item As Object
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
