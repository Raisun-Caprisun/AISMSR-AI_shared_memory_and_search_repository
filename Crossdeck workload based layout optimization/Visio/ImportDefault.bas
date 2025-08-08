Attribute VB_Name = "ImportDefault"
' =========================================================================================
' FIN�LN� SKRIPT PRO IMPORT Z EXCELU DO VISIA
' Verze: 3.0 (�prava pro na�ten� vlastn�ho ID, oprava chyby 1004)
' Popis: Na�te layout na novou str�nku. Pro ka�d� objekt vytvo�� datov� pole "objID"
'        a napln� ho hodnotou ze sloupce A v Excelu.
' =========================================================================================
Public Sub ImportLayoutFromExcel_VlastniID_SharePoint()
    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim pvWindow As Object ' Pro ProtectedViewWindow
    Dim newVisPage As Visio.Page
    Dim visShape As Visio.Shape
    Dim filePath As String
    Dim lastRow As Long, i As Long
    
    ' --- Nastaven� cesty k souboru na SharePointu (p�vodn� a funk�n�) ---
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/visio-excel-objectdata-and-macros.xlsm"
    Const MM_NA_PALEC As Double = 1 / 25.4

    ' --- P�ipojen� k Excelu ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then MsgBox "Nepoda�ilo se spustit Excel.", vbCritical: Exit Sub
    
    xlApp.Visible = True ' Ponechat viditeln� pro lad�n�

    ' Otev�en� se�itu
    On Error Resume Next
    Set xlWb = xlApp.Workbooks.Open(filePath)
    On Error GoTo 0
    If xlWb Is Nothing Then
        MsgBox "CHYBA: Soubor '" & filePath & "' se nepoda�ilo otev��t.", vbCritical
        xlApp.Quit: Set xlApp = Nothing: Exit Sub
    End If
    
    ' --- Stabiliza�n� blok pro SharePoint ---
    On Error Resume Next
    For Each pvWindow In xlApp.ProtectedViewWindows
        If pvWindow.Workbook.Name = xlWb.Name Then
            pvWindow.Edit
            Exit For
        End If
    Next
    On Error GoTo 0
    
    DoEvents
    xlApp.Wait (Now + TimeValue("0:00:02"))
    DoEvents
    
    xlWb.Activate
    Set ws = xlWb.Worksheets(1)
    ws.Activate
    
    ' --- P��prava Visio str�nky ---
    Set newVisPage = Application.ActiveDocument.Pages.Add()
    newVisPage.Name = "Imported Layout " & Format(Now, "yyyy-MM-dd HH-mm-ss")
    newVisPage.Background = False

    ' Spolehliv� zji�t�n� posledn�ho ��dku s daty pomoc� UsedRange
    lastRow = ws.UsedRange.Rows.Count
    
    ' Kontrola, zda jsou v souboru n�jak� data k importu
    If lastRow <= 1 Then
        MsgBox "V Excel souboru nebyla nalezena ��dn� data k importu (krom� hlavi�ky).", vbInformation
        xlWb.Close SaveChanges:=False
        xlApp.Quit
        Exit Sub
    End If
    
    ' --- Hlavn� smy�ka pro import dat ---
    For i = 2 To lastRow
        ' Na�ten� hodnot z Excelu
        Dim objID_val As Variant
        Dim shapeText As String, layerName As String
        Dim centerX_mm As Double, centerY_mm As Double, width_mm As Double, height_mm As Double
        Dim angle_deg As Double, color_rgb As Long
        
        ' O�et�en� pro p��pad, �e by bu�ky byly pr�zdn�
        On Error Resume Next
        objID_val = ws.Cells(i, "A").Value ' <<< ZDE NA��T�ME HODNOTU ID
        shapeText = ws.Cells(i, "C").Value
        layerName = ws.Cells(i, "D").Value
        color_rgb = ws.Cells(i, "E").Value
        centerX_mm = ws.Cells(i, "F").Value
        centerY_mm = ws.Cells(i, "G").Value
        width_mm = ws.Cells(i, "H").Value
        height_mm = ws.Cells(i, "I").Value
        angle_deg = ws.Cells(i, "J").Value
        On Error GoTo 0
        
        ' P�evod hodnot z mm na palce
        Dim centerX_inch As Double, centerY_inch As Double, width_inch As Double, height_inch As Double
        centerX_inch = centerX_mm * MM_NA_PALEC
        centerY_inch = centerY_mm * MM_NA_PALEC
        width_inch = width_mm * MM_NA_PALEC
        height_inch = height_mm * MM_NA_PALEC

        ' Vytvo�en� objektu na nov� str�nce
        Set visShape = newVisPage.DrawRectangle(centerX_inch - (width_inch / 2), _
                                               centerY_inch - (height_inch / 2), _
                                               centerX_inch + (width_inch / 2), _
                                               centerY_inch + (height_inch / 2))
                                            
        ' Nastaven� z�kladn�ch vlastnost�
        visShape.Text = shapeText
        visShape.CellsU("Angle").Result("deg") = angle_deg
        visShape.CellsU("FillForegnd").Result(visColor) = color_rgb
        
        ' =========================================================================
        ' <<< ZDE JE KL��OV� ZM�NA >>>
        ' P�id�n� a nastaven� vlastn�ho ID (objID)
        ' =========================================================================
        If Not IsEmpty(objID_val) And Trim(CStr(objID_val)) <> "" Then
            ' P�id� datov� ��dek s n�zvem "objID", pokud je�t� neexistuje.
            visShape.AddNamedRow visSectionProp, "objID", visTagDefault
            
            ' Nastav� popisek, kter� je vid�t v okn� "Data obrazce"
            visShape.CellsU("Prop.objID.Label").FormulaU = """objID"""
            
            ' Zap�e samotnou hodnotu na�tenou z Excelu
            If IsNumeric(objID_val) Then
                visShape.CellsU("Prop.objID.Value").FormulaU = objID_val
            Else
                visShape.CellsU("Prop.objID.Value").FormulaU = """" & Replace(objID_val, """", """""") & """"
            End If
        End If
        ' =========================================================================

        ' O�et�en� vrstvy
        If Trim(layerName) <> "" Then
            Dim visLayer As Visio.Layer
            On Error Resume Next
            Set visLayer = newVisPage.Layers.ItemU(layerName)
            On Error GoTo 0
            
            If visLayer Is Nothing Then
                Set visLayer = newVisPage.Layers.Add(layerName)
            End If
            
            visLayer.Add visShape, 1
            Set visLayer = Nothing
        End If
    Next i
    
    ' --- Uklizen� ---
    xlWb.Close SaveChanges:=False
    xlApp.Quit
    
    Set pvWindow = Nothing
    Set ws = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set visShape = Nothing
    Set newVisPage = Nothing
    
    MsgBox "Layout byl �sp�n� naimportov�n. Vlastn� ID byla p�i�azena.", vbInformation
End Sub
'### Shrnut� kl��ov�ch �prav:
'1.  **Na�ten� ID z Excelu:** Hned na za��tku smy�ky p�ibyl ��dek `objID_val = ws.Cells(i, "A").Value`, kter� na�te hodnotu ze sloupce "A".
'2.  **Vytvo�en� a napln�n� datov�ho pole:** Za nastaven�m z�kladn�ch vlastnost� (barva, �hel) je nov� blok k�du. Ten zkontroluje, zda bylo n�jak� ID na�teno, a pokud ano:
 '   *   `visShape.AddNamedRow ...`: Vytvo�� pro dan� objekt datov� pole `objID`, pokud je�t� neexistuje.
  '  *   `visShape.CellsU("Prop.objID.Label")...`: Nastav� viditeln� popisek v okn� "Data obrazce".
   ' *   `visShape.CellsU("Prop.objID.Value")...`: Zap�e do pole samotnou hodnotu, kterou na�etl z Excelu.
'
'Tento skript nyn� p�esn� zrcadl� v� exportovac� skript a umo��uje kompletn� "round-trip" va�ich dat.


