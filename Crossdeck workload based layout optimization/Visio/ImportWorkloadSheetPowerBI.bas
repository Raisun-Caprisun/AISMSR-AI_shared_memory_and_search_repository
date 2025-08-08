Attribute VB_Name = "ImportWorkloadSheetPowerBI"
' =========================================================================================
' =========================================================================================
'
'         FIN¡LNÕ SKRIPTY PRO IMPORT A AKTUALIZACI LAYOUTU Z EXCELU
'         Verze: 10.0 (DefinitivnÌ oprava - RozdÏlenÌ na dva kroky pro 100% stabilitu)
'
' =========================================================================================
' INSTRUKCE PRO UéIVATELE:
'
' Krok 1: Spusùte makro "ImportLayout_KROK_1_NakreslitVse".
'         Toto makro smaûe prvnÌ str·nku a nakreslÌ na ni kompletnÌ nov˝ layout z Excelu.
'
' Krok 2: Po dokonËenÌ prvnÌho kroku spusùte makro "ImportLayout_KROK_2_ZobrazitVse".
'         Toto makro p¯izp˘sobÌ velikost str·nky a vycentruje pohled.
'
' D˘vod rozdÏlenÌ: OddÏlenÌ masivnÌch kreslicÌch operacÌ od operacÌ s uûivatelsk˝m
' rozhranÌm (zoom, centrov·nÌ) je jedin˝ 100% spolehliv˝ zp˘sob, jak zabr·nit chyb·m
' zp˘soben˝m nestabilitou aplikace Visio po hromadn˝ch zmÏn·ch.
' =========================================================================================


' =========================================================================================
' SKRIPT 1: Pouze import a kreslenÌ. TENTO SKRIPT SPUSTÕ UéIVATEL JAKO PRVNÕ.
' =========================================================================================
Public Sub ImportLayout_KROK_1_NakreslitVse()
    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim pvWindow As Object ' Pro ProtectedViewWindow
    Dim targetPage As Visio.Page
    Dim visShape As Visio.Shape
    Dim filePath As String
    Dim lastRow As Long, i As Long
    
    ' --- NastavenÌ ---
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/visio-excel-objectdata-and-macros.xlsm"
    Const MM_NA_PALEC As Double = 1 / 25.4

    ' --- CÌlenÌ na prvnÌ str·nku ---
    If Application.ActiveDocument.Pages.Count = 0 Then
        MsgBox "AktivnÌ dokument neobsahuje û·dnÈ str·nky.", vbCritical
        Exit Sub
    End If
    Set targetPage = Application.ActiveDocument.Pages(1)
    
    ' --- Varov·nÌ uûivatele p¯ed smaz·nÌm ---
    If MsgBox("Tato akce smaûe VäECHNY tvary na prvnÌ str·nce ('" & targetPage.Name & "') a nahradÌ je nov˝m layoutem z Excelu." & vbCrLf & vbCrLf & "P¯ejete si pokraËovat?" & vbCrLf & vbCrLf & "UjistÏte se, ze jsou zÛny v Excelu se¯azeny vzestupnÏ podle Z-Order!!!", _
              vbYesNo + vbExclamation, "PotvrzenÌ smaz·nÌ") = vbNo Then
        Exit Sub
    End If

    ' --- P¯ipojenÌ k Excelu ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then MsgBox "Nepoda¯ilo se spustit Excel.", vbCritical: Exit Sub
    
    xlApp.Visible = False

    On Error Resume Next
    Set xlWb = xlApp.Workbooks.Open(filePath)
    On Error GoTo 0
    If xlWb Is Nothing Then
        MsgBox "CHYBA: Soubor '" & filePath & "' se nepoda¯ilo otev¯Ìt.", vbCritical
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
    
    ' --- Smaz·nÌ vöech st·vajÌcÌch tvar˘ na cÌlovÈ str·nce ---
    For i = targetPage.Shapes.Count To 1 Step -1
        targetPage.Shapes(i).Delete
    Next i

    ' --- ZjiötÏnÌ rozsahu dat ---
    lastRow = ws.UsedRange.Rows.Count
    
    If lastRow <= 1 Then
        MsgBox "V Excel souboru nebyla nalezena û·dn· data k importu.", vbInformation
        GoTo Cleanup
    End If
    
    ' --- HlavnÌ smyËka pro import dat ---
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
    ' --- RobustnÌ ˙klid ---
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
    
    MsgBox "DOKON»ENo: Layout byl ˙spÏönÏ nakreslen." & vbCrLf & vbCrLf & "NynÌ jej najdete v listu PowerBI.", vbInformation
End Sub

