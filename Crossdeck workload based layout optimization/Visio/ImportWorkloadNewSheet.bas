Attribute VB_Name = "ImportWorkloadNewSheet"
' =========================================================================================
' SKRIPT PRO IMPORT Z EXCELU DO VISIA (Z NOVÝCH SLOUPCÙ)
' Verze: 3.0
' Popis: Naète layout z Excelu na novou stránku. Tento skript je upraven tak, aby
'        èetl souøadnice støedu (CenterX, CenterY) z nových sloupcù Q a R.
' =========================================================================================
Public Sub ImportLayout_From_New_Columns()
    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim pvWindow As Object ' Pro ProtectedViewWindow
    Dim newVisPage As Visio.Page
    Dim visShape As Visio.Shape
    Dim filePath As String
    Dim lastRow As Long, i As Long
    
    ' --- Nastavení ---
    filePath = "https://trw1-my.sharepoint.com/personal/roman_korpos_zf_com/Documents/Desktop/visio-excel-objectdata-and-macros.xlsm"
    Const MM_NA_PALEC As Double = 1 / 25.4

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
    
    ' --- Stabilizaèní blok ---
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

    ' --- Zjištìní rozsahu dat ---
    lastRow = ws.UsedRange.Rows.Count
    
    If lastRow <= 1 Then
        MsgBox "V Excel souboru nebyla nalezena žádná data k importu.", vbInformation
        xlWb.Close SaveChanges:=False
        xlApp.Quit
        Exit Sub
    End If
    
    ' --- Hlavní smyèka pro import dat ---
    For i = 2 To lastRow
        Dim shapeText As String, layerName As String
        Dim centerX_mm As Double, centerY_mm As Double, width_mm As Double, height_mm As Double
        Dim angle_deg As Double, color_rgb As Long
        
        ' Naètení hodnot z pùvodních sloupcù
        On Error Resume Next
        shapeText = ws.Cells(i, "C").Value
        layerName = ws.Cells(i, "D").Value
        color_rgb = ws.Cells(i, "E").Value
        width_mm = ws.Cells(i, "H").Value
        height_mm = ws.Cells(i, "I").Value
        angle_deg = ws.Cells(i, "J").Value
        
        ' --- ZMÌNA: Naètení souøadnic støedu z NOVÝCH sloupcù ---
        centerX_mm = ws.Cells(i, "Q").Value ' Ètení z New_Center_X
        centerY_mm = ws.Cells(i, "R").Value ' Ètení z New_Center_Y
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
                                            
        ' Nastavení vlastností
        visShape.Text = shapeText
        visShape.CellsU("Angle").Result("deg") = angle_deg
        visShape.CellsU("FillForegnd").Result(visColor) = color_rgb
        
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
    
    MsgBox "Layout byl úspìšnì naimportován z nových souøadnicových sloupcù.", vbInformation
End Sub
