Attribute VB_Name = "ImportDefault"
' =========================================================================================
' FINÁLNÍ SKRIPT PRO IMPORT Z EXCELU DO VISIA
' Verze: 3.0 (Úprava pro naètení vlastního ID, oprava chyby 1004)
' Popis: Naète layout na novou stránku. Pro každý objekt vytvoøí datové pole "objID"
'        a naplní ho hodnotou ze sloupce A v Excelu.
' =========================================================================================
Public Sub ImportLayoutFromExcel_VlastniID_SharePoint()
    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim pvWindow As Object ' Pro ProtectedViewWindow
    Dim newVisPage As Visio.Page
    Dim visShape As Visio.Shape
    Dim filePath As String
    Dim lastRow As Long, i As Long
    
    ' --- Nastavení cesty k souboru na SharePointu (pùvodní a funkèní) ---
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/visio-excel-objectdata-and-macros.xlsm"
    Const MM_NA_PALEC As Double = 1 / 25.4

    ' --- Pøipojení k Excelu ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then MsgBox "Nepodaøilo se spustit Excel.", vbCritical: Exit Sub
    
    xlApp.Visible = True ' Ponechat viditelné pro ladìní

    ' Otevøení sešitu
    On Error Resume Next
    Set xlWb = xlApp.Workbooks.Open(filePath)
    On Error GoTo 0
    If xlWb Is Nothing Then
        MsgBox "CHYBA: Soubor '" & filePath & "' se nepodaøilo otevøít.", vbCritical
        xlApp.Quit: Set xlApp = Nothing: Exit Sub
    End If
    
    ' --- Stabilizaèní blok pro SharePoint ---
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
    
    ' --- Pøíprava Visio stránky ---
    Set newVisPage = Application.ActiveDocument.Pages.Add()
    newVisPage.Name = "Imported Layout " & Format(Now, "yyyy-MM-dd HH-mm-ss")
    newVisPage.Background = False

    ' Spolehlivé zjištìní posledního øádku s daty pomocí UsedRange
    lastRow = ws.UsedRange.Rows.Count
    
    ' Kontrola, zda jsou v souboru nìjaká data k importu
    If lastRow <= 1 Then
        MsgBox "V Excel souboru nebyla nalezena žádná data k importu (kromì hlavièky).", vbInformation
        xlWb.Close SaveChanges:=False
        xlApp.Quit
        Exit Sub
    End If
    
    ' --- Hlavní smyèka pro import dat ---
    For i = 2 To lastRow
        ' Naètení hodnot z Excelu
        Dim objID_val As Variant
        Dim shapeText As String, layerName As String
        Dim centerX_mm As Double, centerY_mm As Double, width_mm As Double, height_mm As Double
        Dim angle_deg As Double, color_rgb As Long
        
        ' Ošetøení pro pøípad, že by buòky byly prázdné
        On Error Resume Next
        objID_val = ws.Cells(i, "A").Value ' <<< ZDE NAÈÍTÁME HODNOTU ID
        shapeText = ws.Cells(i, "C").Value
        layerName = ws.Cells(i, "D").Value
        color_rgb = ws.Cells(i, "E").Value
        centerX_mm = ws.Cells(i, "F").Value
        centerY_mm = ws.Cells(i, "G").Value
        width_mm = ws.Cells(i, "H").Value
        height_mm = ws.Cells(i, "I").Value
        angle_deg = ws.Cells(i, "J").Value
        On Error GoTo 0
        
        ' Pøevod hodnot z mm na palce
        Dim centerX_inch As Double, centerY_inch As Double, width_inch As Double, height_inch As Double
        centerX_inch = centerX_mm * MM_NA_PALEC
        centerY_inch = centerY_mm * MM_NA_PALEC
        width_inch = width_mm * MM_NA_PALEC
        height_inch = height_mm * MM_NA_PALEC

        ' Vytvoøení objektu na nové stránce
        Set visShape = newVisPage.DrawRectangle(centerX_inch - (width_inch / 2), _
                                               centerY_inch - (height_inch / 2), _
                                               centerX_inch + (width_inch / 2), _
                                               centerY_inch + (height_inch / 2))
                                            
        ' Nastavení základních vlastností
        visShape.Text = shapeText
        visShape.CellsU("Angle").Result("deg") = angle_deg
        visShape.CellsU("FillForegnd").Result(visColor) = color_rgb
        
        ' =========================================================================
        ' <<< ZDE JE KLÍÈOVÁ ZMÌNA >>>
        ' Pøidání a nastavení vlastního ID (objID)
        ' =========================================================================
        If Not IsEmpty(objID_val) And Trim(CStr(objID_val)) <> "" Then
            ' Pøidá datový øádek s názvem "objID", pokud ještì neexistuje.
            visShape.AddNamedRow visSectionProp, "objID", visTagDefault
            
            ' Nastaví popisek, který je vidìt v oknì "Data obrazce"
            visShape.CellsU("Prop.objID.Label").FormulaU = """objID"""
            
            ' Zapíše samotnou hodnotu naètenou z Excelu
            If IsNumeric(objID_val) Then
                visShape.CellsU("Prop.objID.Value").FormulaU = objID_val
            Else
                visShape.CellsU("Prop.objID.Value").FormulaU = """" & Replace(objID_val, """", """""") & """"
            End If
        End If
        ' =========================================================================

        ' Ošetøení vrstvy
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
    
    ' --- Uklizení ---
    xlWb.Close SaveChanges:=False
    xlApp.Quit
    
    Set pvWindow = Nothing
    Set ws = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set visShape = Nothing
    Set newVisPage = Nothing
    
    MsgBox "Layout byl úspìšnì naimportován. Vlastní ID byla pøiøazena.", vbInformation
End Sub
'### Shrnutí klíèových úprav:
'1.  **Naètení ID z Excelu:** Hned na zaèátku smyèky pøibyl øádek `objID_val = ws.Cells(i, "A").Value`, který naète hodnotu ze sloupce "A".
'2.  **Vytvoøení a naplnìní datového pole:** Za nastavením základních vlastností (barva, úhel) je nový blok kódu. Ten zkontroluje, zda bylo nìjaké ID naèteno, a pokud ano:
 '   *   `visShape.AddNamedRow ...`: Vytvoøí pro daný objekt datové pole `objID`, pokud ještì neexistuje.
  '  *   `visShape.CellsU("Prop.objID.Label")...`: Nastaví viditelný popisek v oknì "Data obrazce".
   ' *   `visShape.CellsU("Prop.objID.Value")...`: Zapíše do pole samotnou hodnotu, kterou naèetl z Excelu.
'
'Tento skript nyní pøesnì zrcadlí váš exportovací skript a umožòuje kompletní "round-trip" vašich dat.


