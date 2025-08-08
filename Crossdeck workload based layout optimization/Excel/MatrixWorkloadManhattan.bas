Attribute VB_Name = "MatrixWorkloadManhattan"
'---------------------------------------------------------------------------------------
' Module: MatrixGeneratorManhattan
' Version: 1.0
' Popis: Vytvoøí matici vzdáleností "po osách" (Manhattan) na tøetím listu.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub GenerateDistanceMatrix_Manhattan()
    ' --- Nastavení ---
    Dim layoutSheet As Worksheet
    Dim matrixSheet As Worksheet
    
    On Error Resume Next
    Set layoutSheet = ThisWorkbook.Worksheets("Layout")
    If layoutSheet Is Nothing Then
        MsgBox "List 'Layout' nebyl nalezen. Ujistìte se, že první list má tento název.", vbCritical
        Exit Sub
    End If
    
    ' --- ZMÌNA: Cílení na tøetí list ---
    If ThisWorkbook.Worksheets.count < 3 Then
        ' Pøidá listy, dokud jich nejsou alespoò 3
        Do While ThisWorkbook.Worksheets.count < 3
            ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
        Loop
    End If
    Set matrixSheet = ThisWorkbook.Worksheets(3)
    On Error GoTo 0
    
    matrixSheet.Name = "Matrix_Manhattan"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Najdi sloupce ---
    Dim colId As Long:      colId = FindHeaderColumn(layoutSheet, "ID")
    Dim colLayer As Long:   colLayer = FindHeaderColumn(layoutSheet, "Layer")
    Dim colNewX As Long:    colNewX = FindHeaderColumn(layoutSheet, "New_Center_X")
    Dim colNewY As Long:    colNewY = FindHeaderColumn(layoutSheet, "New_Center_Y")
    
    If colId = 0 Or colLayer = 0 Or colNewX = 0 Or colNewY = 0 Then
        MsgBox "Nebyly nalezeny všechny povinné sloupce (ID, Layer, New_Center_X, New_Center_Y).", vbCritical
        GoTo Cleanup
    End If
    
    ' --- Naètení a FILTROVÁNÍ objektù ---
    Dim objects As Collection
    Set objects = New Collection
    
    Dim lastRow As Long
    lastRow = layoutSheet.Cells(layoutSheet.Rows.count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim layerName As String
        layerName = Trim(LCase(CStr(layoutSheet.Cells(i, colLayer).Value)))
        
        If (layerName = "inbound" Or layerName Like "area*") Then
            If IsNumeric(layoutSheet.Cells(i, colNewX).Value) And IsNumeric(layoutSheet.Cells(i, colNewY).Value) Then
                Dim objData As Object
                Set objData = CreateObject("Scripting.Dictionary")
                objData.Add "ID", layoutSheet.Cells(i, colId).Value
                objData.Add "X", CDbl(layoutSheet.Cells(i, colNewX).Value)
                objData.Add "Y", CDbl(layoutSheet.Cells(i, colNewY).Value)
                objects.Add objData
            End If
        End If
    Next i
    
    If objects.count = 0 Then
        MsgBox "Nebyly nalezeny žádné objekty 'Inbound' nebo 'Areas' s platnými souøadnicemi.", vbInformation
        GoTo Cleanup
    End If
    
    ' --- Seøazení objektù podle ID ---
    Dim sortedObjects As Collection
    Set sortedObjects = SortObjectsByID(objects)
    
    ' --- Vytvoøení matice ---
    matrixSheet.Cells.Clear
    
    Dim headerOffset As Integer
    headerOffset = 2
    
    Dim obj As Object
    For i = 1 To sortedObjects.count
        Set obj = sortedObjects(i)
        matrixSheet.Cells(i + headerOffset - 1, 1).Value = obj("ID")
        matrixSheet.Cells(1, i + headerOffset - 1).Value = obj("ID")
    Next i
    
    ' Vypoèítání a vyplnìní vzdáleností
    Dim j As Long
    Dim obj1 As Object, obj2 As Object
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim distance As Double
    
    For i = 1 To sortedObjects.count
        Set obj1 = sortedObjects(i)
        x1 = obj1("X")
        y1 = obj1("Y")
        
        For j = 1 To sortedObjects.count
            Set obj2 = sortedObjects(j)
            x2 = obj2("X")
            y2 = obj2("Y")
            
            ' --- ZMÌNA: Použití Metrické vzdálenosti "po osách" (Manhattan) ---
            distance = Abs(x2 - x1) + Abs(y2 - y1)
            
            matrixSheet.Cells(i + headerOffset - 1, j + headerOffset - 1).Value = distance
        Next j
    Next i
    
    ' Formátování
    matrixSheet.Columns("A:A").AutoFit
    matrixSheet.Rows("1:1").Font.Bold = True
    matrixSheet.Columns("A:A").Font.Bold = True
    matrixSheet.Range(matrixSheet.Cells(2, 2), matrixSheet.Cells(sortedObjects.count + 1, sortedObjects.count + 1)).NumberFormat = "0"
    
    matrixSheet.Activate
    matrixSheet.Cells(1, 1).Select
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Metrická matice vzdáleností 'po osách' byla úspìšnì vytvoøena na listu 'Matrix_Manhattan'.", vbInformation
End Sub

Private Function SortObjectsByID(ByVal coll As Collection) As Collection
    If coll.count <= 1 Then
        Set SortObjectsByID = coll
        Exit Function
    End If
    
    Dim i As Long, j As Long
    Dim temp As Object
    
    ' Simple bubble sort using a temporary collection
    Dim arr As New Collection
    Dim item As Object
    For Each item In coll
        arr.Add item
    Next item

    For i = 1 To arr.count - 1
        For j = i + 1 To arr.count
            If arr(i)("ID") > arr(j)("ID") Then
                Set temp = arr(j)
                arr.Remove j
                arr.Add temp, before:=i
            End If
        Next j
    Next i
    
    Set SortObjectsByID = arr
End Function

Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function
