Attribute VB_Name = "ImportWorkloadSheetPowerBI"
' =========================================================================================
' =========================================================================================
'
'         FIN�LN� SKRIPTY PRO IMPORT A AKTUALIZACI LAYOUTU Z EXCELU
'         Verze: 10.0 (Definitivn� oprava - Rozd�len� na dva kroky pro 100% stabilitu)
'
' =========================================================================================
' INSTRUKCE PRO U�IVATELE:
'
' Krok 1: Spus�te makro "ImportLayout_KROK_1_NakreslitVse".
'         Toto makro sma�e prvn� str�nku a nakresl� na ni kompletn� nov� layout z Excelu.
'
' Krok 2: Po dokon�en� prvn�ho kroku spus�te makro "ImportLayout_KROK_2_ZobrazitVse".
'         Toto makro p�izp�sob� velikost str�nky a vycentruje pohled.
'
' D�vod rozd�len�: Odd�len� masivn�ch kreslic�ch operac� od operac� s u�ivatelsk�m
' rozhran�m (zoom, centrov�n�) je jedin� 100% spolehliv� zp�sob, jak zabr�nit chyb�m
' zp�soben�m nestabilitou aplikace Visio po hromadn�ch zm�n�ch.
' =========================================================================================


' =========================================================================================
' SKRIPT 1: Pouze import a kreslen�. TENTO SKRIPT SPUST� U�IVATEL JAKO PRVN�.
' =========================================================================================
Public Sub ImportLayout_KROK_1_NakreslitVse()
    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim pvWindow As Object ' Pro ProtectedViewWindow
    Dim targetPage As Visio.Page
    Dim visShape As Visio.Shape
    Dim filePath As String
    Dim lastRow As Long, i As Long
    
    ' --- Nastaven� ---
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/visio-excel-objectdata-and-macros.xlsm"
    Const MM_NA_PALEC As Double = 1 / 25.4

    ' --- C�len� na prvn� str�nku ---
    If Application.ActiveDocument.Pages.Count = 0 Then
        MsgBox "Aktivn� dokument neobsahuje ��dn� str�nky.", vbCritical
        Exit Sub
    End If
    Set targetPage = Application.ActiveDocument.Pages(1)
    
    ' --- Varov�n� u�ivatele p�ed smaz�n�m ---
    If MsgBox("Tato akce sma�e V�ECHNY tvary na prvn� str�nce ('" & targetPage.Name & "') a nahrad� je nov�m layoutem z Excelu." & vbCrLf & vbCrLf & "P�ejete si pokra�ovat?" & vbCrLf & vbCrLf & "Ujist�te se, ze jsou z�ny v Excelu se�azeny vzestupn� podle Z-Order!!!", _
              vbYesNo + vbExclamation, "Potvrzen� smaz�n�") = vbNo Then
        Exit Sub
    End If

    ' --- P�ipojen� k Excelu ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then MsgBox "Nepoda�ilo se spustit Excel.", vbCritical: Exit Sub
    
    xlApp.Visible = False

    On Error Resume Next
    Set xlWb = xlApp.Workbooks.Open(filePath)
    On Error GoTo 0
    If xlWb Is Nothing Then
        MsgBox "CHYBA: Soubor '" & filePath & "' se nepoda�ilo otev��t.", vbCritical
        xlApp.Quit: Set xlApp = Nothing: Exit Sub
    End If
    
    On Error Resume Next
    For Each pvWindow In xlApp.ProtectedViewWindows
        If pvWindow.Workbook.Name = xlWb.Name Then
            pvWindow.Edit
            Exit For
        End If
    Next
    On Error GoTo 0
    
    DoEvents
    Set ws = xlWb.Worksheets(1)
    
    ' --- Smaz�n� v�ech st�vaj�c�ch tvar� na c�lov� str�nce ---
    For i = targetPage.Shapes.Count To 1 Step -1
        targetPage.Shapes(i).Delete
    Next i

    ' --- Zji�t�n� rozsahu dat ---
    lastRow = ws.UsedRange.Rows.Count
    
    If lastRow <= 1 Then
        MsgBox "V Excel souboru nebyla nalezena ��dn� data k importu.", vbInformation
        GoTo Cleanup
    End If
    
    ' --- Hlavn� smy�ka pro import dat ---
    For i = 2 To lastRow
        Dim objID_val As Variant, shapeText As String, layerName As String
        Dim centerX_mm As Double, centerY_mm As Double, width_mm As Double, height_mm As Double
        Dim angle_deg As Double, color_rgb As Long
        
        On Error Resume Next
        objID_val = ws.Cells(i, "A").Value
        shapeText = ws.Cells(i, "C").Value
        layerName = ws.Cells(i, "D").Value
        color_rgb = ws.Cells(i, "E").Value
        width_mm = ws.Cells(i, "H").Value
        height_mm = ws.Cells(i, "I").Value
        angle_deg = ws.Cells(i, "J").Value
        centerX_mm = ws.Cells(i, "Q").Value
        centerY_mm = ws.Cells(i, "R").Value
        On Error GoTo 0
        
        If Not IsNumeric(centerX_mm) Or Not IsNumeric(centerY_mm) Then GoTo NextIteration
        
        Dim centerX_inch As Double, centerY_inch As Double, width_inch As Double, height_inch As Double
        centerX_inch = centerX_mm * MM_NA_PALEC
        centerY_inch = centerY_mm * MM_NA_PALEC
        width_inch = width_mm * MM_NA_PALEC
        height_inch = height_mm * MM_NA_PALEC

        Set visShape = targetPage.DrawRectangle(centerX_inch - (width_inch / 2), centerY_inch - (height_inch / 2), centerX_inch + (width_inch / 2), centerY_inch + (height_inch / 2))
        
        visShape.Text = shapeText
        visShape.CellsU("Angle").Result("deg") = angle_deg
        visShape.CellsU("FillForegnd").Result(visColor) = color_rgb
        visShape.CellsU("Char.Size").Result("pt") = 30

        If Not IsEmpty(objID_val) And Trim(CStr(objID_val)) <> "" Then
            visShape.AddNamedRow visSectionProp, "objID", visTagDefault
            visShape.CellsU("Prop.objID.Label").FormulaU = """objID"""
            If IsNumeric(objID_val) Then
                visShape.CellsU("Prop.objID.Value").FormulaU = objID_val
            Else
                visShape.CellsU("Prop.objID.Value").FormulaU = """" & Replace(objID_val, """", """""") & """"
            End If
        End If
        
        If Trim(layerName) <> "" Then
            Dim visLayer As Visio.Layer
            On Error Resume Next
            Set visLayer = targetPage.Layers.ItemU(layerName)
            On Error GoTo 0
            If visLayer Is Nothing Then Set visLayer = targetPage.Layers.Add(layerName)
            visLayer.Add visShape, 1
            Set visLayer = Nothing
        End If
        
        visShape.BringToFront

NextIteration:
    Next i

Cleanup:
    ' --- Robustn� �klid ---
    On Error Resume Next
    If Not xlWb Is Nothing Then
        xlWb.Close SaveChanges:=False
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If
    On Error GoTo 0
    
    Set pvWindow = Nothing
    Set ws = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set visShape = Nothing
    Set targetPage = Nothing
    
    MsgBox "DOKON�ENo: Layout byl �sp�n� nakreslen." & vbCrLf & vbCrLf & "Nyn� jej najdete v listu PowerBI.", vbInformation
End Sub

