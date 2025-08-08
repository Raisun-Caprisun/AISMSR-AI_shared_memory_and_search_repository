Attribute VB_Name = "ExportDefault"
' =========================================================================================
' =========================================================================================
'
'         FINÁLNÍ SKRIPT PRO EXPORT LAYOUTU Z VISIA DO EXCELU
'         Verze: 6.0 (Stabilní verze s .ResultIU a detailní dokumentací)
'
' =========================================================================================
' PØED SPUŠTÌNÍM TOHOTO SKRIPTU SE UJISTÌTE, ŽE JSOU SPLNÌNY NÁSLEDUJÍCÍ PODMÍNKY:
' -----------------------------------------------------------------------------------------
'
' Tento skript exportuje data obrazcù (shapes) z aktivní stránky Visia do Excelu. Aby
' fungoval správnì a spolehlivì, musí být zdrojová data ve Visiu peèlivì pøipravena.
'
' === KLÍÈOVÁ PØÍPRAVA PRO VŠECHNY OBJEKTY VE VISIU ===
'
' 1. EXISTENCE DATOVÉHO POLE:
'    Každý objekt, který má být exportován s vlastním ID, musí mít definované
'    datové pole (Shape Data). Pøidává se pøes pravý klik -> Data -> Definovat data obrazce.
'
' 2. SPRÁVNÝ INTERNÍ NÁZEV POLE:
'    Interní "Název" (Name) datového pole musí být PØESNÌ "objID". Nestaèí, aby byl
'    jen "Popisek" (Label) nastaven na "objID". Interní název je to, co skript používá
'    pro identifikaci pole. Zkontrolujte v dialogu "Definovat data obrazce".
'
' 3. SPRÁVNÝ DATOVÝ TYP:
'    "Typ" (Type) datového pole "objID" musí být nastaven na "Èíslo" (Number).
'    Toto zajišuje, že skript mùže hodnotu správnì interpretovat jako celé èíslo.
'
' 4. PØIØAZENÁ HODNOTA:
'    Každému objektu musí být v poli "objID" pøiøazena èíselná hodnota.
'
' === LIMITACE A DÙLEŽITÁ UPOZORNÌNÍ ===
'
' 1. LIMIT HODNOTY ID: max. 32 767
'    Tento skript používá metodu .ResultIU, která naèítá hodnotu jako datový typ
'    "Integer". Tento typ má maximální hodnotu 32 767. Pokud jakýkoliv objekt
'    bude mít v "objID" hodnotu vyšší než 32767, skript tuto hodnotu nenaète
'    a buòka v Excelu pro toto ID zùstane prázdná! Dá se upravit pouzitim Result.Str, ale je treba upravit skript - doporucuji Google AI Studio.
'
' 2. CESTA K SOUBORU:
'    Ujistìte se, že promìnná "filePath" obsahuje správnou a funkèní cestu
'    k cílovému .xlsm souboru na SharePointu nebo lokálním disku.
'
' 3. CÍLOVÝ LIST:
'    Skript zapisuje data na první list (index 1) v Excel sešitu.
'
' 4. MAZÁNÍ DAT:
'    Pøed každým exportem skript vymaže veškerý obsah v rozsahu A2:P50000 na
'    cílovém listu. Ujistìte se, že v tomto rozsahu nemáte žádná data, která
'    chcete zachovat.
'
' =========================================================================================
' =========================================================================================

Public Sub ExportLayoutuDoExcelu_Finalni_s_Dokumentaci()
    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim pvWindow As Object ' Pro ProtectedViewWindow
    Dim visShape As Visio.Shape
    Dim filePath As String
    Dim i As Long
    
    ' Promìnné pro Bounding Box
    Dim bBoxLeftInches As Double, bBoxBottomInches As Double, bBoxRightInches As Double, bBoxTopInches As Double
    
    ' --- Nastavení cesty k souboru na SharePointu ---
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/visio-excel-objectdata-and-macros.xlsm"

    ' --- Pøipojení k Excelu ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then MsgBox "Nepodaøilo se spustit Excel.", vbCritical: Exit Sub
    
    xlApp.Visible = True

    On Error Resume Next
    Set xlWb = xlApp.Workbooks.Open(filePath)
    On Error GoTo 0
    If xlWb Is Nothing Then
        MsgBox "CHYBA: Soubor '" & filePath & "' se nepodaøilo otevøít.", vbCritical
        xlApp.Quit: Set xlApp = Nothing: Exit Sub
    End If
    
    ' Ošetøení Chránìného zobrazení
    If xlApp.ProtectedViewWindows.Count > 0 Then
        For Each pvWindow In xlApp.ProtectedViewWindows
            If pvWindow.Workbook.FullName = xlWb.FullName Then
                pvWindow.Edit
                xlApp.Wait (Now + TimeValue("0:00:02"))
                Exit For
            End If
        Next
    End If
    
    Set ws = xlWb.Worksheets(1)
    xlWb.Activate
    ws.Activate

    ' --- Vymazání starých dat ---
    ws.Range("A2:P50000").ClearContents
    
    ' --- Pøíprava hlavièky ---
    ws.Range("A1:P1").Value = Array("ID", "Name", "Text", "Layer", "Color (RGB)", "CenterX", "CenterY", "Width", "Height", "Angle", "Z-Order", "BBox_Left_X", "BBox_Right_X", "BBox_Bottom_Y", "BBox_Top_Y", "Workload")
    
    ' --- Hlavní smyèka pro export dat ---
    i = 2
    For Each visShape In ActivePage.Shapes
        ' Získání BoundingBox
        visShape.BoundingBox visBBoxUprightWH, bBoxLeftInches, bBoxBottomInches, bBoxRightInches, bBoxTopInches

        ' Naètení vlastního ID pomocí metody .ResultIU. Vyžaduje perfektnì èistá data.
        On Error Resume Next
        ws.Cells(i, "A").Value = visShape.CellsU("Prop.objID").ResultIU
        On Error GoTo 0
        
        ws.Cells(i, "B").Value = visShape.Name
        ws.Cells(i, "C").Value = visShape.Text
        On Error Resume Next
        ws.Cells(i, "D").Value = visShape.Layer(1).Name
        On Error GoTo 0
        
        ws.Cells(i, "E").Value = visShape.CellsU("FillForegnd").Result(visColor)
        
        ws.Cells(i, "F").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Result("mm")
        ws.Cells(i, "G").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Result("mm")
        
        ws.Cells(i, "H").Value = visShape.CellsU("Width").Result("mm")
        ws.Cells(i, "I").Value = visShape.CellsU("Height").Result("mm")
        
        ws.Cells(i, "J").Value = visShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormAngle).Result("deg")
        
        ws.Cells(i, "K").Value = i - 1
        
        Const PALEC_NA_MM As Double = 25.4
        ws.Cells(i, "L").Value = bBoxLeftInches * PALEC_NA_MM
        ws.Cells(i, "M").Value = bBoxRightInches * PALEC_NA_MM
        ws.Cells(i, "N").Value = bBoxBottomInches * PALEC_NA_MM
        ws.Cells(i, "O").Value = bBoxTopInches * PALEC_NA_MM
        
        i = i + 1
    Next visShape
    
    ' --- Uklizení ---
    ws.Columns("A:P").AutoFit
    xlWb.Save
    xlWb.Close
    xlApp.Quit
    
    Set pvWindow = Nothing
    Set ws = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    MsgBox "Export byl úspìšnì dokonèen. Ujistìte se, že všechna ID byla správnì zapsána.", vbInformation
End Sub
