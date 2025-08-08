Attribute VB_Name = "ExportDefault"
' =========================================================================================
' =========================================================================================
'
'         FIN�LN� SKRIPT PRO EXPORT LAYOUTU Z VISIA DO EXCELU
'         Verze: 6.0 (Stabiln� verze s .ResultIU a detailn� dokumentac�)
'
' =========================================================================================
' P�ED SPU�T�N�M TOHOTO SKRIPTU SE UJIST�TE, �E JSOU SPLN�NY N�SLEDUJ�C� PODM�NKY:
' -----------------------------------------------------------------------------------------
'
' Tento skript exportuje data obrazc� (shapes) z aktivn� str�nky Visia do Excelu. Aby
' fungoval spr�vn� a spolehliv�, mus� b�t zdrojov� data ve Visiu pe�liv� p�ipravena.
'
' === KL��OV� P��PRAVA PRO V�ECHNY OBJEKTY VE VISIU ===
'
' 1. EXISTENCE DATOV�HO POLE:
'    Ka�d� objekt, kter� m� b�t exportov�n s vlastn�m ID, mus� m�t definovan�
'    datov� pole (Shape Data). P�id�v� se p�es prav� klik -> Data -> Definovat data obrazce.
'
' 2. SPR�VN� INTERN� N�ZEV POLE:
'    Intern� "N�zev" (Name) datov�ho pole mus� b�t P�ESN� "objID". Nesta��, aby byl
'    jen "Popisek" (Label) nastaven na "objID". Intern� n�zev je to, co skript pou��v�
'    pro identifikaci pole. Zkontrolujte v dialogu "Definovat data obrazce".
'
' 3. SPR�VN� DATOV� TYP:
'    "Typ" (Type) datov�ho pole "objID" mus� b�t nastaven na "��slo" (Number).
'    Toto zaji��uje, �e skript m��e hodnotu spr�vn� interpretovat jako cel� ��slo.
'
' 4. P�I�AZEN� HODNOTA:
'    Ka�d�mu objektu mus� b�t v poli "objID" p�i�azena ��seln� hodnota.
'
' === LIMITACE A D�LE�IT� UPOZORN�N� ===
'
' 1. LIMIT HODNOTY ID: max. 32 767
'    Tento skript pou��v� metodu .ResultIU, kter� na��t� hodnotu jako datov� typ
'    "Integer". Tento typ m� maxim�ln� hodnotu 32 767. Pokud jak�koliv objekt
'    bude m�t v "objID" hodnotu vy��� ne� 32767, skript tuto hodnotu nena�te
'    a bu�ka v Excelu pro toto ID z�stane pr�zdn�! D� se upravit pouzitim Result.Str, ale je treba upravit skript - doporucuji Google AI Studio.
'
' 2. CESTA K SOUBORU:
'    Ujist�te se, �e prom�nn� "filePath" obsahuje spr�vnou a funk�n� cestu
'    k c�lov�mu .xlsm souboru na SharePointu nebo lok�ln�m disku.
'
' 3. C�LOV� LIST:
'    Skript zapisuje data na prvn� list (index 1) v Excel se�itu.
'
' 4. MAZ�N� DAT:
'    P�ed ka�d�m exportem skript vyma�e ve�ker� obsah v rozsahu A2:P50000 na
'    c�lov�m listu. Ujist�te se, �e v tomto rozsahu nem�te ��dn� data, kter�
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
    
    ' Prom�nn� pro Bounding Box
    Dim bBoxLeftInches As Double, bBoxBottomInches As Double, bBoxRightInches As Double, bBoxTopInches As Double
    
    ' --- Nastaven� cesty k souboru na SharePointu ---
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/visio-excel-objectdata-and-macros.xlsm"

    ' --- P�ipojen� k Excelu ---
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then MsgBox "Nepoda�ilo se spustit Excel.", vbCritical: Exit Sub
    
    xlApp.Visible = True

    On Error Resume Next
    Set xlWb = xlApp.Workbooks.Open(filePath)
    On Error GoTo 0
    If xlWb Is Nothing Then
        MsgBox "CHYBA: Soubor '" & filePath & "' se nepoda�ilo otev��t.", vbCritical
        xlApp.Quit: Set xlApp = Nothing: Exit Sub
    End If
    
    ' O�et�en� Chr�n�n�ho zobrazen�
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

    ' --- Vymaz�n� star�ch dat ---
    ws.Range("A2:P50000").ClearContents
    
    ' --- P��prava hlavi�ky ---
    ws.Range("A1:P1").Value = Array("ID", "Name", "Text", "Layer", "Color (RGB)", "CenterX", "CenterY", "Width", "Height", "Angle", "Z-Order", "BBox_Left_X", "BBox_Right_X", "BBox_Bottom_Y", "BBox_Top_Y", "Workload")
    
    ' --- Hlavn� smy�ka pro export dat ---
    i = 2
    For Each visShape In ActivePage.Shapes
        ' Z�sk�n� BoundingBox
        visShape.BoundingBox visBBoxUprightWH, bBoxLeftInches, bBoxBottomInches, bBoxRightInches, bBoxTopInches

        ' Na�ten� vlastn�ho ID pomoc� metody .ResultIU. Vy�aduje perfektn� �ist� data.
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
    
    ' --- Uklizen� ---
    ws.Columns("A:P").AutoFit
    xlWb.Save
    xlWb.Close
    xlApp.Quit
    
    Set pvWindow = Nothing
    Set ws = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    MsgBox "Export byl �sp�n� dokon�en. Ujist�te se, �e v�echna ID byla spr�vn� zaps�na.", vbInformation
End Sub
