import os
import sys
import win32com.client
from datetime import datetime, timedelta

def rgb_to_int(r, g, b):
    """Excel uses 0xBBGGRR format (BGR integer)."""
    return r + (g << 8) + (b << 16)

def create_pallet_workbook():
    workspace = os.path.abspath(os.getcwd())
    target_path = os.path.join(workspace, "Palet_Karti_Sistemi.xlsm")
    old_target_path = os.path.join(workspace, "Palet_Karti_Yedek_Sistemi.xlsm")
    original = os.path.join(workspace, "Book1.xlsm")
    backup_book1 = os.path.join(workspace, "Book1_Yedek.xlsm")

    # If Book1.xlsm exists, backup
    if os.path.exists(original) and not os.path.exists(backup_book1):
        try:
            import shutil
            shutil.copy2(original, backup_book1)
            print(f"Orijinal Book1.xlsm yedeklendi: {backup_book1}")
        except Exception as e:
            print(f"Yedekleme uyarisi: {e}")

    # Remove target file if exists (checking for file locks)
    for path_to_clean in [target_path, old_target_path]:
        if os.path.exists(path_to_clean):
            try:
                os.remove(path_to_clean)
            except PermissionError:
                print(f"HATA: '{os.path.basename(path_to_clean)}' dosyasi su anda baska bir programda (muhtemelen Excel) acik!")
                print("Lutfen dosyayi kapatin ve komutu tekrar calistirin.")
                return False
            except Exception as e:
                print(f"Eski dosya temizleme uyarisi ({path_to_clean}): {e}")

    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Add()

        # Ensure exactly 2 sheets
        while wb.Sheets.Count < 2:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        while wb.Sheets.Count > 2:
            wb.Sheets(wb.Sheets.Count).Delete()

        ws_input = wb.Sheets(1)
        ws_input.Name = "VERI_GIRIS"

        ws_cards = wb.Sheets(2)
        ws_cards.Name = "PALET_KARTLARI"

        # ----------------------------------------------------
        # 1. SETUP VERI_GIRIS SHEET
        # ----------------------------------------------------
        ws_input.Tab.Color = rgb_to_int(45, 106, 79)

        # Column widths
        ws_input.Columns("A").ColumnWidth = 3
        ws_input.Columns("B").ColumnWidth = 26
        ws_input.Columns("C").ColumnWidth = 24
        ws_input.Columns("D").ColumnWidth = 24
        ws_input.Columns("E").ColumnWidth = 4
        ws_input.Columns("F").ColumnWidth = 28
        ws_input.Columns("G").ColumnWidth = 28
        ws_input.Columns("H").ColumnWidth = 4

        # Title Header
        title_rng = ws_input.Range("B2:G2")
        title_rng.Merge()
        title_rng.Value = "📦 PALET KARTI YÖNETİM & BASKI SİSTEMİ"
        title_rng.Font.Name = "Segoe UI"
        title_rng.Font.Size = 16
        title_rng.Font.Bold = True
        title_rng.Font.Color = rgb_to_int(255, 255, 255)
        title_rng.Interior.Color = rgb_to_int(27, 67, 50)
        title_rng.HorizontalAlignment = -4108
        title_rng.VerticalAlignment = -4108
        ws_input.Rows(2).RowHeight = 35

        sub_rng = ws_input.Range("B3:G3")
        sub_rng.Merge()
        sub_rng.Value = "Palet bilgilerini ve basılacak adet sayısını girin, ardından 'Kartları Göster' veya 'Yazdır' butonuna tıklayın."
        sub_rng.Font.Name = "Segoe UI"
        sub_rng.Font.Size = 10
        sub_rng.Font.Italic = True
        sub_rng.Font.Color = rgb_to_int(85, 85, 85)
        sub_rng.HorizontalAlignment = -4108
        ws_input.Rows(3).RowHeight = 20

        # Section 1 Header: Parti & Ürün Bilgileri
        sec1 = ws_input.Range("B5:D5")
        sec1.Merge()
        sec1.Value = "📋 PARTİ VE ÜRÜN BİLGİLERİ"
        sec1.Font.Name = "Segoe UI"
        sec1.Font.Size = 11
        sec1.Font.Bold = True
        sec1.Font.Color = rgb_to_int(255, 255, 255)
        sec1.Interior.Color = rgb_to_int(45, 106, 79)
        sec1.HorizontalAlignment = -4131
        ws_input.Rows(5).RowHeight = 24

        today_str = datetime.now().strftime("%Y-%m-%d")
        deliv_str = (datetime.now() + timedelta(days=14)).strftime("%Y-%m-%d")

        fields = [
            (6, "Müşteri / Firma Adı:", ""),
            (7, "Ürün Türü / Cinsi:", ""),
            (8, "Çeşit / Model:", ""),
            (9, "Parti / Lot Numarası:", "LOT-2026/001"),
            (10, "Sipariş / Sevk No:", "SIP-1001"),
            (11, "Üretim / Hazırlık Tarihi:", today_str),
            (12, "Tahmini Teslim / Sevk Tarihi:", deliv_str),
            (13, "Kasa / Koli İçi Adet (Göz):", 216),
            (14, "Paletteki Koli / Kasa Adedi:", 15),
            (15, "Palet Başına Toplam Adet:", "=C13*C14"),
            (16, "Kalite / Kontrol Durumu:", "KONTROL EDİLDİ - UYGUN"),
            (17, "Sorumlu Personel:", "")
        ]

        for row_idx, label, default_val in fields:
            ws_input.Rows(row_idx).RowHeight = 22
            # Label
            lbl_cell = ws_input.Cells(row_idx, 2)
            lbl_cell.Value = label
            lbl_cell.Font.Name = "Segoe UI"
            lbl_cell.Font.Size = 10
            lbl_cell.Font.Bold = True
            lbl_cell.Font.Color = rgb_to_int(34, 34, 34)
            lbl_cell.Interior.Color = rgb_to_int(240, 247, 242)
            lbl_cell.Borders.Color = rgb_to_int(208, 229, 213)

            # Value cell
            val_rng = ws_input.Range(f"C{row_idx}:D{row_idx}")
            val_rng.Merge()
            val_rng.Value = default_val
            val_rng.Font.Name = "Segoe UI"
            val_rng.Font.Size = 10
            val_rng.Font.Bold = (row_idx in [6, 7, 8, 15])
            val_rng.Interior.Color = rgb_to_int(255, 255, 255)
            val_rng.Borders.Color = rgb_to_int(208, 229, 213)
            val_rng.HorizontalAlignment = -4131

            if row_idx in [11, 12]:
                val_rng.NumberFormat = "yyyy-mm-dd"
            elif row_idx in [13, 14]:
                val_rng.NumberFormat = "#,##0"
                val_rng.HorizontalAlignment = -4108
            elif row_idx == 15:
                val_rng.NumberFormat = '#,##0" Adet / Palet"'
                val_rng.Font.Color = rgb_to_int(27, 106, 45)
                val_rng.Interior.Color = rgb_to_int(232, 248, 237)
                val_rng.HorizontalAlignment = -4108

        # Data Validation (Dropdown) for Kasa / Koli İçi Adet (C13)
        try:
            val_box = ws_input.Range("C13:D13").Validation
            val_box.Delete()
            val_box.Add(3, 1, 1, "45, 77, 104, 150, 200, 216, 288, 384")
            val_box.IgnoreBlank = True
            val_box.InCellDropdown = True
        except Exception as e:
            print(f"Veri dogrulama (C13) uyarisi: {e}")

        # Data Validation (Dropdown) for Kalite / Kontrol Durumu (C16)
        try:
            val_status = ws_input.Range("C16:D16").Validation
            val_status.Delete()
            val_status.Add(3, 1, 1, "KONTROL EDİLDİ - UYGUN, BİRİNCİ KALİTE, STANDART SEVKİYAT, İLAÇLI - KORUMALI, ÖZEL SİPARİŞ")
            val_status.IgnoreBlank = True
            val_status.InCellDropdown = True
        except Exception as e:
            print(f"Veri dogrulama (C16) uyarisi: {e}")

        # Section 2: Basılacak Kart Adedi
        sec2 = ws_input.Range("B19:D19")
        sec2.Merge()
        sec2.Value = "⚡ BASILACAK PALET (KART) SAYISI"
        sec2.Font.Name = "Segoe UI"
        sec2.Font.Size = 11
        sec2.Font.Bold = True
        sec2.Font.Color = rgb_to_int(255, 255, 255)
        sec2.Interior.Color = rgb_to_int(30, 74, 140)
        sec2.HorizontalAlignment = -4131
        ws_input.Rows(19).RowHeight = 24

        ws_input.Rows(20).RowHeight = 32
        lbl_card = ws_input.Cells(20, 2)
        lbl_card.Value = "Toplam Palet (Kart) Sayısı:"
        lbl_card.Font.Name = "Segoe UI"
        lbl_card.Font.Size = 11
        lbl_card.Font.Bold = True
        lbl_card.Interior.Color = rgb_to_int(230, 240, 250)
        lbl_card.Borders.Color = rgb_to_int(153, 188, 232)

        cnt_rng = ws_input.Range("C20:D20")
        cnt_rng.Merge()
        cnt_rng.Value = 12
        cnt_rng.Font.Name = "Segoe UI"
        cnt_rng.Font.Size = 16
        cnt_rng.Font.Bold = True
        cnt_rng.Font.Color = rgb_to_int(198, 40, 40)
        cnt_rng.Interior.Color = rgb_to_int(204, 229, 255)
        cnt_rng.HorizontalAlignment = -4108
        cnt_rng.VerticalAlignment = -4108
        cnt_rng.Borders.Weight = -4138
        cnt_rng.Borders.Color = rgb_to_int(30, 74, 140)

        # Help Note below
        note_rng = ws_input.Range("B21:D21")
        note_rng.Merge()
        note_rng.Value = "💡 Buraya yazdığınız sayı kadar kart ekranda gösterilir ve yazıcıdan basılır."
        note_rng.Font.Name = "Segoe UI"
        note_rng.Font.Size = 9
        note_rng.Font.Italic = True
        note_rng.Font.Color = rgb_to_int(102, 102, 102)

        # Action Instruction Box in Column F:G
        info_header = ws_input.Range("F5:G5")
        info_header.Merge()
        info_header.Value = "🚀 HIZLI İŞLEM BUTONLARI"
        info_header.Font.Name = "Segoe UI"
        info_header.Font.Size = 11
        info_header.Font.Bold = True
        info_header.Font.Color = rgb_to_int(255, 255, 255)
        info_header.Interior.Color = rgb_to_int(27, 67, 50)
        info_header.HorizontalAlignment = -4108
        ws_input.Rows(5).RowHeight = 24

        # Description text under buttons
        desc_rng = ws_input.Range("F18:G21")
        desc_rng.Merge()
        desc_rng.Value = "📌 NASIL KULLANILIR?\n1. Sol taraftaki parti bilgilerini doldurun.\n2. Basılacak kart sayısını (C20) girin.\n3. 'KARTLARI GÖSTER' ile kontrol edin veya 'YAZDIR' ile A4 çıktısı alın."
        desc_rng.Font.Name = "Segoe UI"
        desc_rng.Font.Size = 9.5
        desc_rng.Font.Color = rgb_to_int(68, 68, 68)
        desc_rng.Interior.Color = rgb_to_int(247, 250, 248)
        desc_rng.Borders.Color = rgb_to_int(208, 229, 213)
        desc_rng.VerticalAlignment = -4160
        desc_rng.WrapText = True

        # ----------------------------------------------------
        # 2. SETUP PALET_KARTLARI SHEET
        # ----------------------------------------------------
        ws_cards.Tab.Color = rgb_to_int(30, 74, 140)

        # ----------------------------------------------------
        # 3. ADD VBA MACRO MODULE
        # ----------------------------------------------------
        vba_code = """Option Explicit

' ==========================================================
' PALET KARTI OTOMASYON MODÜLÜ
' ==========================================================

Public Sub Makro_KartlariGoster()
    Dim wsInput As Worksheet, wsCards As Worksheet
    Dim totalCards As Long, i As Long
    Dim startRow As Long, cardHeight As Long, cardSpacing As Long
    Dim custName As String, pType As String, variety As String
    Dim lotNo As String, orderNo As String, prodDate As String, delDate As String
    Dim violsPerPallet As Variant, cellCount As Variant, totalSeedlings As Variant
    Dim statusText As String, responsible As String
    Dim lastCardEndRow As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wsInput = ThisWorkbook.Sheets("VERI_GIRIS")
    Set wsCards = ThisWorkbook.Sheets("PALET_KARTLARI")
    
    totalCards = Val(wsInput.Range("C20").Value)
    If totalCards < 1 Then
        MsgBox "Lütfen en az 1 adet palet kartı sayısı giriniz!", vbExclamation, "Geçersiz Sayı"
        GoTo CleanExit
    End If
    If totalCards > 500 Then
        If MsgBox(totalCards & " adet kart oluşturulacak. Devam etmek istiyor musunuz?", vbQuestion + vbYesNo, "Kart Sayısı Onayı") = vbNo Then
            GoTo CleanExit
        End If
    End If
    
    custName = Trim(wsInput.Range("C6").Value)
    pType = Trim(wsInput.Range("C7").Value)
    variety = Trim(wsInput.Range("C8").Value)
    lotNo = Trim(wsInput.Range("C9").Value)
    orderNo = Trim(wsInput.Range("C10").Value)
    
    If IsDate(wsInput.Range("C11").Value) Then
        prodDate = Format(wsInput.Range("C11").Value, "dd.mm.yyyy")
    Else
        prodDate = Trim(wsInput.Range("C11").Text)
    End If
    
    If IsDate(wsInput.Range("C12").Value) Then
        delDate = Format(wsInput.Range("C12").Value, "dd.mm.yyyy")
    Else
        delDate = Trim(wsInput.Range("C12").Text)
    End If
    
    cellCount = wsInput.Range("C13").Value
    violsPerPallet = wsInput.Range("C14").Value
    totalSeedlings = wsInput.Range("C15").Value
    statusText = Trim(wsInput.Range("C16").Value)
    responsible = Trim(wsInput.Range("C17").Value)
    
    wsCards.Cells.Clear
    wsCards.ResetAllPageBreaks
    
    ' Page Setup
    With wsCards.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With
    
    wsCards.Columns("A").ColumnWidth = 2
    wsCards.Columns("B").ColumnWidth = 14
    wsCards.Columns("C").ColumnWidth = 18
    wsCards.Columns("D").ColumnWidth = 14
    wsCards.Columns("E").ColumnWidth = 14
    wsCards.Columns("F").ColumnWidth = 14
    wsCards.Columns("G").ColumnWidth = 16
    wsCards.Columns("H").ColumnWidth = 2
    
    cardHeight = 13
    cardSpacing = 2
    
    For i = 1 To totalCards
        startRow = (i - 1) * (cardHeight + cardSpacing) + 2
        Call BuildSingleCard(wsCards, startRow, i, totalCards, custName, pType, variety, lotNo, orderNo, prodDate, delDate, cellCount, violsPerPallet, totalSeedlings, statusText, responsible)
        
        If (i Mod 2 = 0) And (i < totalCards) Then
            wsCards.HPageBreaks.Add Before:=wsCards.Rows(startRow + cardHeight + cardSpacing)
        End If
    Next i
    
    lastCardEndRow = (totalCards - 1) * (cardHeight + cardSpacing) + 2 + cardHeight
    wsCards.PageSetup.PrintArea = "$B$1:$G$" & lastCardEndRow
    
    wsCards.Activate
    wsCards.Range("B2").Select
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Bir hata oluştu: " & Err.Description, vbCritical, "Hata"
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub BuildSingleCard(ws As Worksheet, r As Long, cardNo As Long, totalCards As Long, _
                            custName As String, pType As String, variety As String, _
                            lotNo As String, orderNo As String, prodDate As String, delDate As String, _
                            cellCount As Variant, violsPerPallet As Variant, totalSeedlings As Variant, _
                            statusText As String, responsible As String)
    Dim rngCard As Range
    Dim rHeader As Long, rSub As Long, r1 As Long, r2 As Long, r3 As Long, r4 As Long, r5 As Long, rFooter As Long
    
    rHeader = r
    rSub = r + 1
    r1 = r + 2
    r2 = r + 4
    r3 = r + 6
    r4 = r + 8
    r5 = r + 10
    rFooter = r + 12
    
    ws.Rows(rHeader).RowHeight = 28
    ws.Rows(rSub).RowHeight = 22
    ws.Rows(r1).RowHeight = 15
    ws.Rows(r1 + 1).RowHeight = 22
    ws.Rows(r2).RowHeight = 15
    ws.Rows(r2 + 1).RowHeight = 22
    ws.Rows(r3).RowHeight = 15
    ws.Rows(r3 + 1).RowHeight = 22
    ws.Rows(r4).RowHeight = 15
    ws.Rows(r4 + 1).RowHeight = 24
    ws.Rows(r5).RowHeight = 15
    ws.Rows(r5 + 1).RowHeight = 22
    ws.Rows(rFooter + 1).RowHeight = 14
    
    With ws.Range(ws.Cells(rHeader, 2), ws.Cells(rHeader, 7))
        .Merge
        .Value = "📦 PALET TANITIM KARTI"
        .Font.Name = "Segoe UI"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(27, 67, 50)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(rSub, 2), ws.Cells(rSub, 5))
        .Merge
        .Value = "  LOT NO: " & lotNo & "  |  SİPARİŞ NO: " & orderNo
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(27, 67, 50)
        .Interior.Color = RGB(232, 245, 233)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(rSub, 6), ws.Cells(rSub, 7))
        .Merge
        .Value = "PALET NO: " & Format(cardNo, "00") & " / " & Format(totalCards, "00")
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(198, 40, 40)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Call FormatField(ws, r1, 2, 4, "MÜŞTERİ / FİRMA ADI", custName, True)
    Call FormatField(ws, r1, 5, 7, "ÜRÜN / TÜR", pType, True)
    
    Call FormatField(ws, r2, 2, 4, "ÇEŞİT / MODEL", variety, True)
    Call FormatField(ws, r2, 5, 7, "ÜRETİM / HAZIRLIK TARİHİ", prodDate, False)
    
    Call FormatField(ws, r3, 2, 3, "TESLİM / SEVK TARİHİ", delDate, False)
    Call FormatField(ws, r3, 4, 5, "BİRİM / KASA İÇİ ADET", cellCount & " Adet", False)
    Call FormatField(ws, r3, 6, 7, "PALETTEKİ KASA / KOLİ ADEDİ", violsPerPallet & " Koli/Kasa", False)
    
    With ws.Range(ws.Cells(r4, 2), ws.Cells(r4 + 1, 4))
        .Merge
        .Value = "TOPLAM ÜRÜN: " & Format(totalSeedlings, "#,##0") & " ADET"
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(27, 67, 50)
        .Interior.Color = RGB(216, 243, 220)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Call FormatField(ws, r4, 5, 7, "KALİTE & KONTROL DURUMU", statusText, False)
    
    With ws.Range(ws.Cells(r5, 2), ws.Cells(r5 + 1, 4))
        .Merge
        .Value = "SORUMLU: " & IIf(Len(responsible) > 0, responsible, "-")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(50, 50, 50)
        .Interior.Color = RGB(248, 249, 250)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range(ws.Cells(r5, 5), ws.Cells(r5 + 1, 7))
        .Merge
        .Value = "KALİTE ONAYI: [ ONAYLANDI ]"
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(27, 94, 32)
        .Interior.Color = RGB(232, 245, 233)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Set rngCard = ws.Range(ws.Cells(rHeader, 2), ws.Cells(r5 + 1, 7))
    With rngCard.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(27, 67, 50)
    End With
    With rngCard.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(27, 67, 50)
    End With
    With rngCard.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(27, 67, 50)
    End With
    With rngCard.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(27, 67, 50)
    End With
    
    If cardNo < totalCards Then
        With ws.Range(ws.Cells(rFooter + 1, 2), ws.Cells(rFooter + 1, 7))
            .Merge
            .Value = "- - - - - - - - - - - - - - - - - - - - ✂ MAKASLA KESİM ÇİZGİSİ ✂ - - - - - - - - - - - - - - - - - - - -"
            .Font.Name = "Segoe UI"
            .Font.Size = 8
            .Font.Italic = True
            .Font.Color = RGB(160, 160, 160)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
End Sub

Private Sub FormatField(ws As Worksheet, r As Long, cStart As Long, cEnd As Long, lbl As String, val As String, isBig As Boolean)
    With ws.Range(ws.Cells(r, cStart), ws.Cells(r, cEnd))
        .Merge
        .Value = lbl
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.Color = RGB(203, 213, 225)
    End With
    
    With ws.Range(ws.Cells(r + 1, cStart), ws.Cells(r + 1, cEnd))
        .Merge
        .Value = val
        .Font.Name = "Segoe UI"
        .Font.Size = IIf(isBig, 11, 10)
        .Font.Bold = True
        .Font.Color = RGB(15, 23, 42)
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.Color = RGB(203, 213, 225)
    End With
End Sub

Public Sub Makro_Yazdir()
    Call Makro_KartlariGoster
    ThisWorkbook.Sheets("PALET_KARTLARI").PrintPreview
End Sub

Public Sub Makro_PDF_Kaydet()
    Dim wsCards As Worksheet
    Dim pdfPath As String, timeStamp As String, desktopDir As String
    Dim objShell As Object
    
    Call Makro_KartlariGoster
    Set wsCards = ThisWorkbook.Sheets("PALET_KARTLARI")
    
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    desktopDir = objShell.SpecialFolders("Desktop")
    If Len(desktopDir) = 0 Then
        desktopDir = Environ("USERPROFILE") & "\\Desktop"
    End If
    
    timeStamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    pdfPath = desktopDir & "\\Palet_Kartlari_" & timeStamp & ".pdf"
    
    wsCards.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                                IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    If Err.Number = 0 Then
        MsgBox "PDF belgeniz başarıyla masaüstüne kaydedildi ve açılıyor:" & vbCrLf & pdfPath, vbInformation, "PDF Kaydedildi"
    Else
        MsgBox "PDF kaydedilirken bir hata oluştu: " & Err.Description, vbCritical, "Hata"
    End If
End Sub

Public Sub Makro_Temizle()
    Dim wsInput As Worksheet
    If MsgBox("Formdaki veriler sıfırlanacak. Emin misiniz?", vbQuestion + vbYesNo, "Formu Sıfırla") = vbYes Then
        Set wsInput = ThisWorkbook.Sheets("VERI_GIRIS")
        wsInput.Range("C6:D10").ClearContents
        wsInput.Range("C16:D17").ClearContents
        wsInput.Range("C13").Value = 216
        wsInput.Range("C14").Value = 15
        wsInput.Range("C20").Value = 1
        MsgBox "Form temizlendi!", vbInformation, "Bilgi"
    End If
End Sub
"""

        # Add VBA module to workbook
        vb_comp = wb.VBProject.VBComponents.Add(1)
        vb_comp.Name = "Modul_PaletKarti"
        vb_comp.CodeModule.AddFromString(vba_code)
        print("VBA Modülü (Modul_PaletKarti) eklendi.")

        # ----------------------------------------------------
        # 4. ADD ACTION BUTTONS ON VERI_GIRIS SHEET
        # ----------------------------------------------------
        buttons_config = [
            ("btnGen", "⚡ KARTLARI GÖSTER", "Makro_KartlariGoster", 7, rgb_to_int(45, 106, 79), 0xFFFFFF),
            ("btnPrint", "🖨️ YAZDIR / BASKI ÖNİZLEME", "Makro_Yazdir", 10, rgb_to_int(30, 74, 140), 0xFFFFFF),
            ("btnPDF", "📥 PDF OLARAK KAYDET", "Makro_PDF_Kaydet", 13, rgb_to_int(180, 80, 20), 0xFFFFFF),
            ("btnClear", "🔄 FORMU TEMİZLE", "Makro_Temizle", 16, rgb_to_int(90, 90, 90), 0xFFFFFF)
        ]

        for name, text, macro_name, row_pos, bg_color, fg_color in buttons_config:
            top_cell = ws_input.Cells(row_pos, 6)
            btn_left = top_cell.Left + 5
            btn_top = top_cell.Top + 2
            btn_width = ws_input.Range(f"F{row_pos}:G{row_pos}").Width - 10
            btn_height = 34

            shape = ws_input.Shapes.AddShape(5, btn_left, btn_top, btn_width, btn_height)
            shape.Name = name
            shape.TextFrame.Characters().Text = text
            shape.TextFrame.Characters().Font.Name = "Segoe UI"
            shape.TextFrame.Characters().Font.Size = 11
            shape.TextFrame.Characters().Font.Bold = True
            shape.TextFrame.Characters().Font.Color = fg_color
            shape.TextFrame.HorizontalAlignment = -4108
            shape.TextFrame.VerticalAlignment = -4108
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = bg_color
            shape.Line.Visible = False
            shape.OnAction = macro_name

        # ----------------------------------------------------
        # 5. INITIALIZE CARDS USING THE VBA MACRO (SINGLE SOURCE OF TRUTH)
        # ----------------------------------------------------
        print("VBA Makrosu çalıştırılarak ilk kartlar oluşturuluyor...")
        excel.Run(f"'{wb.Name}'!Makro_KartlariGoster")

        # Focus back on VERI_GIRIS
        ws_input.Activate()
        ws_input.Range("C6").Select()

        # Save workbooks
        wb.SaveAs(target_path, 52)
        print(f"Başarıyla kaydedildi: {target_path}")

        wb.SaveCopyAs(original)
        print(f"Kopya kaydedildi: {original}")

        wb.Close(False)
        print("İşlem tamamlandı!")
        return True

    except Exception as e:
        print(f"Hata: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return False
    finally:
        try:
            if wb is not None:
                del wb
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
                del excel
        except Exception:
            pass

if __name__ == "__main__":
    success = create_pallet_workbook()
    if success:
        print("EXCEL DOSYASI BAŞARIYLA OLUŞTURULDU!")
    else:
        print("EXCEL OLUŞTURMA BAŞARISIZ OLDU!")
        sys.exit(1)
