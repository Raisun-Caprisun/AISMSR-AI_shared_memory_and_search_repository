Attribute VB_Name = "ImportDataFromData_CD"
'---------------------------------------------------------------------------------------
' Module: DataImporters
' Version: 1.0 - Imports Workload and Buffer from Data_CD.xlsm
' Description: This macro runs from "InputData.xlsm" and pulls data from the "Workload"
'              sheet in "Data_CD.xlsm". It matches rows based on Area ID and updates
'              the "Workload" and "Max_Buffer" columns.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub ImportWorkloadAndBufferData()
    ' --- CONFIGURATION ---
    Const sourceFilePath As String = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/Data_CD.xlsm"
    Const sourceSheetName As String = "Workload"
    
    ' --- SETUP ---
    Dim sourceWb As Workbook, sourceSheet As Worksheet
    Dim destWb As Workbook, destSheet As Worksheet
    
    Set destWb = ThisWorkbook
    Set destSheet = destWb.Worksheets(1) ' Assumes data is on the first sheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Open the source workbook (Data_CD.xlsm) ---
    On Error Resume Next
    Set sourceWb = Workbooks.Open(sourceFilePath, ReadOnly:=True)
    If sourceWb Is Nothing Then
        MsgBox "Failed to open the source data file at the specified path:" & vbCrLf & sourceFilePath, vbCritical
        GoTo Cleanup
    End If
    
    On Error Resume Next
    Set sourceSheet = sourceWb.Worksheets(sourceSheetName)
    If sourceSheet Is Nothing Then
        MsgBox "The sheet named '" & sourceSheetName & "' was not found in the source file.", vbCritical
        sourceWb.Close SaveChanges:=False
        GoTo Cleanup
    End If
    On Error GoTo 0
    
    ' --- 1. Load data from Data_CD.xlsm into a Dictionary ---
    ' The Dictionary will store the Area ID as the key and an array with (Workload, Max_Buffer) as the value.
    Dim dataToImport As Object
    Set dataToImport = CreateObject("Scripting.Dictionary")
    
    Dim lastSourceRow As Long
    lastSourceRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastSourceRow ' Assuming headers are in row 1
        Dim keyID As Variant: keyID = sourceSheet.Cells(r, "A").Value
        
        ' Ensure the ID is not empty and is a number before processing
        If Not IsEmpty(keyID) And IsNumeric(keyID) Then
            Dim keyIDLong As Long: keyIDLong = CLng(keyID)
            
            If Not dataToImport.Exists(keyIDLong) Then
                ' Read values from columns B (Workload) and C (Max_Buffer)
                Dim workloadValue As Variant: workloadValue = sourceSheet.Cells(r, "B").Value
                Dim bufferValue As Variant: bufferValue = sourceSheet.Cells(r, "C").Value
                
                ' Add the data as an array to the dictionary
                dataToImport.Add keyIDLong, Array(workloadValue, bufferValue)
            End If
        End If
    Next r
    
    ' --- Close the source workbook now that we have its data in memory ---
    sourceWb.Close SaveChanges:=False
    
    ' --- 2. Find the required columns in THIS destination file ---
    Dim colDestID As Long: colDestID = FindHeaderColumn(destSheet, "ID")
    Dim colDestWorkload As Long: colDestWorkload = FindHeaderColumn(destSheet, "Workload")
    Dim colDestBuffer As Long: colDestBuffer = FindHeaderColumn(destSheet, "Max_Buffer")
    
    If colDestID = 0 Or colDestWorkload = 0 Or colDestBuffer = 0 Then
        MsgBox "Could not find one or more required columns ('ID', 'Workload', 'Max_Buffer') in this file.", vbCritical
        GoTo Cleanup
    End If
    
    ' --- 3. Update the "Workload" and "Max_Buffer" columns in this sheet ---
    Dim lastDestRow As Long
    lastDestRow = destSheet.Cells(destSheet.Rows.Count, colDestID).End(xlUp).Row
    
    Dim updatedCount As Long: updatedCount = 0
    For r = 2 To lastDestRow
        Dim destKeyID As Variant
        destKeyID = destSheet.Cells(r, colDestID).Value
        
        If Not IsEmpty(destKeyID) And IsNumeric(destKeyID) Then
            Dim destKeyIDLong As Long: destKeyIDLong = CLng(destKeyID)
            
            ' Check if a matching ID was found in the source data
            If dataToImport.Exists(destKeyIDLong) Then
                Dim dataArray As Variant
                dataArray = dataToImport(destKeyIDLong)
                
                ' Write the values to the respective columns
                destSheet.Cells(r, colDestWorkload).Value = dataArray(0) ' Workload
                destSheet.Cells(r, colDestBuffer).Value = dataArray(1)   ' Max_Buffer
                
                updatedCount = updatedCount + 1
            End If
        End If
    Next r
    
    MsgBox "Import complete." & vbCrLf & vbCrLf & "Updated " & updatedCount & " rows from the Data_CD file.", vbInformation

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Set sourceSheet = Nothing
    Set sourceWb = Nothing
    Set destSheet = Nothing
    Set destWb = Nothing
    Set dataToImport = Nothing
End Sub


Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    ' Helper function to find a column number by its header text in the first row.
    On Error Resume Next
    FindHeaderColumn = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0
End Function
