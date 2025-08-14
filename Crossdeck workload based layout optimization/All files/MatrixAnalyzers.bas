Attribute VB_Name = "MatrixAnalyzers"
'---------------------------------------------------------------------------------------
' Module: MatrixAnalyzers
' Version: 4.3 - Corrected Automation Error on Close
' Description: This version fixes the "Automation Error" by storing the destination
'              workbook's name in a variable before the workbook object is closed.
'---------------------------------------------------------------------------------------
Option Explicit

'========================================================================================
'      PUBLIC MACRO FOR SIMULATION EXPORT
'========================================================================================

Public Sub ExportMatrixForSimulation()
    ' --- CONFIGURATION ---
    Const destFilePath As String = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/Data_CD.xlsm"
    Const sourceSheetName As String = "Matrix_Optimized_Euclidean"
    Const destSheetName As String = "MaticeVzdalenosti"
    Const DIVISOR As Double = 1000 ' To convert mm to m
    
    ' --- SETUP ---
    Dim sourceWb As Workbook: Set sourceWb = ThisWorkbook
    Dim sourceSheet As Worksheet
    Dim destWb As Workbook, destSheet As Worksheet
    Dim destFileName As String ' *** FIX: Variable to safely store the file name ***
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- 1. Verify that the source matrix exists ---
    On Error Resume Next
    Set sourceSheet = sourceWb.Worksheets(sourceSheetName)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        MsgBox "The source sheet '" & sourceSheetName & "' was not found." & vbCrLf & _
               "Please generate the optimized matrices first before running this export.", vbCritical
        GoTo Cleanup
    End If
    
    ' --- 2. Open the destination workbook ---
    On Error Resume Next
    Set destWb = Workbooks.Open(destFilePath)
    If destWb Is Nothing Then
        MsgBox "Failed to open the destination file for the simulation:" & vbCrLf & destFilePath, vbCritical
        GoTo Cleanup
    End If
    ' *** FIX: Store the name immediately after opening the workbook ***
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
    
    ' *** FIX: Use the stored file name variable in the message box ***
    MsgBox "The distance matrix has been successfully exported and scaled for the simulation in '" & destFileName & "'.", vbInformation

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'========================================================================================
'      EXISTING PUBLIC MACROS - No changes to these
'========================================================================================

Public Sub GenerateAllMatrices()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    GenerateMatrix_Master useOptimizedCoords:=True, useEuclideanCalc:=True, showMsg:=False
    GenerateMatrix_Master useOptimizedCoords:=True, useEuclideanCalc:=False, showMsg:=False
    GenerateMatrix_Master useOptimizedCoords:=False, useEuclideanCalc:=True, showMsg:=False
    GenerateMatrix_Master useOptimizedCoords:=False, useEuclideanCalc:=False, showMsg:=False
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "All four distance matrices have been successfully generated.", vbInformatio

