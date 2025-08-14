Attribute VB_Name = "WidthOptimizer"
'---------------------------------------------------------------------------------------
' Module: WidthOptimizer
' Version: 1.0 - Tiered Width Recalculation
' Description: This macro recalculates the 'New_Width' column for all "Areas" based
'              on a tiered system driven by the 'Max_Buffer' value, while respecting
'              a total width constraint.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub RecalculateAreaWidths()
    ' --- CONFIGURATION ---
    Const TIER_1_WIDTH As Long = 7200
    Const TIER_2_WIDTH As Long = 4800
    Const TIER_3_WIDTH As Long = 2400
    
    Const TIER_1_PERCENT As Double = 0.03333
    Const TIER_2_PERCENT As Double = 0.21666
    
    Const MAX_TOTAL_WIDTH As Long = 184800
    
    ' --- SETUP ---
    Dim inputSheet As Worksheet
    Set inputSheet = ThisWorkbook.Worksheets(1) ' Assumes data is on the first sheet
    
    Application.ScreenUpdating = False
    
    ' --- Find all necessary columns by header name ---
    Dim cols As Object: Set cols = CreateObject("Scripting.Dictionary")
    cols.Add "Text", FindHeaderColumn(inputSheet, "Text")
    cols.Add "Layer", FindHeaderColumn(inputSheet, "Layer")
    cols.Add "New_Width", FindHeaderColumn(inputSheet, "New_Width")
    cols.Add "Max_Buffer", FindHeaderColumn(inputSheet, "Max_Buffer")
    
    If cols("Layer") = 0 Or cols("Max_Buffer") = 0 Or cols("New_Width") = 0 Then
        MsgBox "Could not find required columns ('Layer', 'Max_Buffer', 'New_Width'). Please check the headers.", vbCritical
        Exit Sub
    End If
    
    ' --- 1. Read and Filter all "Areas" data ---
    Dim areasData As Collection
    Set areasData = LoadAreaData(inputSheet, cols)
    
    If areasData.Count = 0 Then
        MsgBox "No rows with Layer = 'Areas' were found.", vbInformation
        Exit Sub
    End If
    
    ' --- 2. Sort the areas by Max_Buffer in descending order ---
    Dim sortedAreas As Collection
    Set sortedAreas = SortByMetric(areasData, "Max_Buffer", False) ' False for Descending
    
    ' --- 3. Calculate the number of slots for each tier ---
    Dim totalAreas As Long: totalAreas = sortedAreas.Count
    Dim numTier1 As Long: numTier1 = Round(totalAreas * TIER_1_PERCENT, 0)
    Dim numTier2 As Long: numTier2 = Round(totalAreas * TIER_2_PERCENT, 0)
    
    ' --- 4 & 5. Assign New Widths, Calculate Sum, and Write Back to Sheet ---
    Dim totalNewWidth As Double: totalNewWidth = 0
    Dim i As Long
    Dim currentArea As Object
    Dim newWidth As Long
    
    For i = 1 To totalAreas
        Set currentArea = sortedAreas(i)
        
        ' Assign width based on tiered position
        If i <= numTier1 Then
            newWidth = TIER_1_WIDTH
        ElseIf i <= (numTier1 + numTier2) Then
            newWidth = TIER_2_WIDTH
        Else
            newWidth = TIER_3_WIDTH
        End If
        
        ' Write the new width back to the correct row in the Excel sheet
        inputSheet.Cells(currentArea("Row"), cols("New_Width")).Value = newWidth
        
        ' Add to the running total
        totalNewWidth = totalNewWidth + newWidth
    Next i
    
    ' --- 6. Verify Constraint and Report to User ---
    Dim msg As String
    msg = "Width recalculation complete." & vbCrLf & vbCrLf & _
          "Total Areas Processed: " & totalAreas & vbCrLf & _
          "Tier 1 Slots (@" & TIER_1_WIDTH & "mm): " & numTier1 & vbCrLf & _
          "Tier 2 Slots (@" & TIER_2_WIDTH & "mm): " & numTier2 & vbCrLf & _
          "Tier 3 Slots (@" & TIER_3_WIDTH & "mm): " & totalAreas - numTier1 - numTier2 & vbCrLf & vbCrLf & _
          "Calculated Total Width: " & totalNewWidth & vbCrLf & _
          "Constraint Limit: " & MAX_TOTAL_WIDTH & vbCrLf & vbCrLf
          
    If totalNewWidth <= MAX_TOTAL_WIDTH Then
        msg = msg & "Result: CONSTRAINT MET."
        MsgBox msg, vbInformation
    Else
        msg = msg & "Result: WARNING - CONSTRAINT FAILED."
        MsgBox msg, vbExclamation
    End If
    
    Application.ScreenUpdating = True
End Sub


Private Function LoadAreaData(ByVal ws As Worksheet, ByVal cols As Object) As Collection
    ' Reads all rows with Layer="Areas" into a collection of dictionary objects
    Set LoadAreaData = New Collection
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        If Trim(LCase(CStr(ws.Cells(r, cols("Layer")).Value))) Like "area*" Then
            Dim area As Object: Set area = CreateObject("Scripting.Dictionary")
            area.Add "Row", r ' Store original row number
            area.Add "Text", ws.Cells(r, cols("Text")).Value
            area.Add "Max_Buffer", CDbl(Nz(ws.Cells(r, cols("Max_Buffer")).Value))
            LoadAreaData.Add area
        End If
    Next r
End Function

Private Function SortByMetric(ByVal coll As Collection, ByVal sortKey As String, Optional ByVal ascending As Boolean = True) As Collection
    ' Sorts a collection of dictionary objects by a specified key
    If coll.Count <= 1 Then Set SortByMetric = coll: Exit Function
    
    Dim i As Long, j As Long, temp As Object
    Dim arr As New Collection
    For Each temp In coll: arr.Add temp: Next
    
    For i = 1 To arr.Count - 1
        For j = i + 1 To arr.Count
            Dim condition As Boolean
            If ascending Then
                condition = (arr(i)(sortKey) > arr(j)(sortKey))
            Else
                condition = (arr(i)(sortKey) < arr(j)(sortKey))
            End If
            
            If condition Then
                Set temp = arr(j)
                arr.Remove j
                arr.Add temp, before:=i
            End If
        Next j
    Next i
    Set SortByMetric = arr
End Function


' --- Self-Contained Helper Functions ---
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

