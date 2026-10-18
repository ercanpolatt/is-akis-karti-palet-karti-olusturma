import os
import sys
import win32com.client

def rgb_to_int(r, g, b):
    # Excel uses 0xBBGGRR format (BGR integer)
    return r + (g << 8) + (b << 16)

def format_field(ws, r, c_start, c_end, lbl, val, is_big=False):
    # Header label
    lbl_rng = ws.Range(ws.Cells(r, c_start), ws.Cells(r, c_end))
    lbl_rng.Merge()
    lbl_rng.Value = lbl
    lbl_rng.Font.Name = "Segoe UI"
    lbl_rng.Font.Size = 8
    lbl_rng.Font.Bold = True
    lbl_rng.Font.Color = rgb_to_int(100, 116, 139) # Slate grey
    lbl_rng.Interior.Color = rgb_to_int(241, 245, 249) # Soft light grey
    lbl_rng.HorizontalAlignment = -4108
    lbl_rng.VerticalAlignment = -4108
    lbl_rng.Borders.Color = rgb_to_int(203, 213, 225)

    # Value cell
    val_rng = ws.Range(ws.Cells(r + 1, c_start), ws.Cells(r + 1, c_end))
    val_rng.Merge()
    val_rng.Value = val
    val_rng.Font.Name = "Segoe UI"
    val_rng.Font.Size = 11 if is_big else 10
    val_rng.Font.Bold = True
    val_rng.Font.Color = rgb_to_int(15, 23, 42)
    val_rng.Interior.Color = rgb_to_int(255, 255, 255)
    val_rng.HorizontalAlignment = -4108
    val_rng.VerticalAlignment = -4108
    val_rng.Borders.Color = rgb_to_int(203, 213, 225)

def build_single_card_py(ws, r, card_no, total_cards, cust_name, f_type, variety, lot_no, order_no, sow_date, del_date, cell_count, viols_per_pallet, total_seedlings, status_text, responsible):
    r_header = r
    r_sub = r + 1
    r1 = r + 2
    r2 = r + 4
    r3 = r + 6
    r4 = r + 8
    r5 = r + 10
    r_footer = r + 12

    # Row heights
    ws.Rows(r_header).RowHeight = 28
    ws.Rows(r_sub).RowHeight = 22
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
    ws.Rows(r_footer + 1).RowHeight = 14

    # 1. Main Header
    h_rng = ws.Range(ws.Cells(r_header, 2), ws.Cells(r_header, 7))
    h_rng.Merge()
    h_rng.Value = "🌱 FİDE PALET TANITIM KARTI"
    h_rng.Font.Name = "Segoe UI"
    h_rng.Font.Size = 13
    h_rng.Font.Bold = True
    h_rng.Font.Color = rgb_to_int(255, 255, 255)
    h_rng.Interior.Color = rgb_to_int(27, 67, 50) # #1B4332 Dark Green
    h_rng.HorizontalAlignment = -4108
    h_rng.VerticalAlignment = -4108

    # 2. Subheader
    sub_left = ws.Range(ws.Cells(r_sub, 2), ws.Cells(r_sub, 5))
    sub_left.Merge()
    sub_left.Value = f"  LOT NO: {lot_no}  |  SİPARİŞ NO: {order_no}"
    sub_left.Font.Name = "Segoe UI"
    sub_left.Font.Size = 9.5
    sub_left.Font.Bold = True
    sub_left.Font.Color = rgb_to_int(27, 67, 50)
    sub_left.Interior.Color = rgb_to_int(232, 245, 233)
    sub_left.HorizontalAlignment = -4131
    sub_left.VerticalAlignment = -4108

    sub_right = ws.Range(ws.Cells(r_sub, 6), ws.Cells(r_sub, 7))
    sub_right.Merge()
    sub_right.Value = f"PALET NO: {card_no:02d} / {total_cards:02d}"
    sub_right.Font.Name = "Segoe UI"
    sub_right.Font.Size = 11
    sub_right.Font.Bold = True
    sub_right.Font.Color = rgb_to_int(255, 255, 255)
    sub_right.Interior.Color = rgb_to_int(198, 40, 40) # Red Badge
    sub_right.HorizontalAlignment = -4108
    sub_right.VerticalAlignment = -4108

    # Rows 1 to 3
    format_field(ws, r1, 2, 4, "MÜŞTERİ / FİRMA ADI", cust_name, True)
    format_field(ws, r1, 5, 7, "FİDE TÜRÜ (ÜRÜN)", f_type, True)

    format_field(ws, r2, 2, 4, "ÇEŞİT / VARYETE", variety, True)
    format_field(ws, r2, 5, 7, "TOHUM EKİM TARİHİ", sow_date, False)

    format_field(ws, r3, 2, 3, "TESLİM / ÇIKIŞ TARİHİ", del_date, False)
    format_field(ws, r3, 4, 5, "VİYOL GÖZ SAYISI", f"{cell_count} Gözlü", False)
    format_field(ws, r3, 6, 7, "PALETTEKİ VİYOL ADEDİ", f"{viols_per_pallet} Kasa/Viyol", False)

    # Row 4: Total Seedlings & Status
    seedling_box = ws.Range(ws.Cells(r4, 2), ws.Cells(r4 + 1, 4))
    seedling_box.Merge()
    seedling_box.Value = f"TOPLAM FİDE: {total_seedlings:,} ADET".replace(",", ".")
    seedling_box.Font.Name = "Segoe UI"
    seedling_box.Font.Size = 12
    seedling_box.Font.Bold = True
    seedling_box.Font.Color = rgb_to_int(27, 67, 50)
    seedling_box.Interior.Color = rgb_to_int(216, 243, 220)
    seedling_box.HorizontalAlignment = -4108
    seedling_box.VerticalAlignment = -4108

    format_field(ws, r4, 5, 7, "AŞILAMA & İLAÇLAMA DURUMU", status_text, False)

    # Row 5: Responsible & Quality
    resp_box = ws.Range(ws.Cells(r5, 2), ws.Cells(r5 + 1, 4))
    resp_box.Merge()
    resp_box.Value = f"SORUMLU: {responsible}"
    resp_box.Font.Name = "Segoe UI"
    resp_box.Font.Size = 9
    resp_box.Font.Bold = True
    resp_box.Font.Color = rgb_to_int(50, 50, 50)
    resp_box.Interior.Color = rgb_to_int(248, 249, 250)
    resp_box.HorizontalAlignment = -4108
    resp_box.VerticalAlignment = -4108

    qual_box = ws.Range(ws.Cells(r5, 5), ws.Cells(r5 + 1, 7))
    qual_box.Merge()
    qual_box.Value = "KALİTE ONAYI: [ ONAYLANDI ]"
    qual_box.Font.Name = "Segoe UI"
    qual_box.Font.Size = 9.5
    qual_box.Font.Bold = True
    qual_box.Font.Color = rgb_to_int(27, 94, 32)
    qual_box.Interior.Color = rgb_to_int(232, 245, 233)
    qual_box.HorizontalAlignment = -4108
    qual_box.VerticalAlignment = -4108

    # Outer Border (Medium Green)
    card_rng = ws.Range(ws.Cells(r_header, 2), ws.Cells(r5 + 1, 7))
    for edge in [7, 8, 9, 10]: # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
        card_rng.Borders(edge).LineStyle = 1 # xlContinuous
        card_rng.Borders(edge).Weight = -4138 # xlMedium
        card_rng.Borders(edge).Color = rgb_to_int(27, 67, 50)

    # Cut Line
    if card_no < total_cards:
        cut_rng = ws.Range(ws.Cells(r_footer + 1, 2), ws.Cells(r_footer + 1, 7))
        cut_rng.Merge()
        cut_rng.Value = "- - - - - - - - - - - - - - - - - - - - ✂ MAKASLA KESİM ÇİZGİSİ ✂ - - - - - - - - - - - - - - - - - - - -"
        cut_rng.Font.Name = "Segoe UI"
        cut_rng.Font.Size = 8
        cut_rng.Font.Italic = True
        cut_rng.Font.Color = rgb_to_int(160, 160, 160)
        cut_rng.HorizontalAlignment = -4108
        cut_rng.VerticalAlignment = -4108

def build_all_cards_py(ws_cards, total_cards=12):
    cust_name = "ÖRNEK TARIM İŞLETMESİ A.Ş."
    f_type = "DOMATES"
    variety = "TORRY F1"
    lot_no = "LOT-2026/089"
    order_no = "SIP-45802"
    sow_date = "21.08.2026"
    del_date = "10.09.2026"
    cell_count = 216
    viols_per_pallet = 15
    total_seedlings = cell_count * viols_per_pallet
    status_text = "AŞILI - KORUYUCU İLAÇLAMA YAPILDI"
    responsible = "Zir. Müh. Ahmet YILMAZ"

    card_height = 13
    card_spacing = 2

    ws_cards.ResetAllPageBreaks()
    for i in range(1, total_cards + 1):
        start_row = (i - 1) * (card_height + card_spacing) + 2
        build_single_card_py(ws_cards, start_row, i, total_cards, cust_name, f_type, variety, lot_no, order_no, sow_date, del_date, cell_count, viols_per_pallet, total_seedlings, status_text, responsible)
        if (i % 2 == 0) and (i < total_cards):
            ws_cards.HPageBreaks.Add(Before=ws_cards.Rows(start_row + card_height + card_spacing))

    last_card_end = (total_cards - 1) * (card_height + card_spacing) + 2 + card_height
    ws_cards.PageSetup.PrintArea = f"$B$1:$G${last_card_end}"

def create_pallet_workbook():
    workspace = os.path.abspath(os.getcwd())
    target_path = os.path.join(workspace, "Palet_Karti_Fide_Sistemi.xlsm")
    original = os.path.join(workspace, "Book1.xlsm")
    backup_book1 = os.path.join(workspace, "Book1_Yedek.xlsm")

    # If Book1.xlsm exists, backup
    if os.path.exists(original) and not os.path.exists(backup_book1):
        try:
            import shutil
            shutil.copy2(original, backup_book1)
            print(f"Original Book1.xlsm backed up as {backup_book1}")
        except Exception as e:
            print(f"Backup notice: {e}")

    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except Exception as e:
            print(f"Notice on removing existing target_path: {e}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Add()
        
        # Ensure we have 2 sheets
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
        title_rng.Value = "🌱 FİDE PALET KARTI YÖNETİM & BASKI SİSTEMİ"
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

        # Section 1 Header: Parti & Müşteri Bilgileri
        sec1 = ws_input.Range("B5:D5")
        sec1.Merge()
        sec1.Value = "📋 PARTİ VE FİDE BİLGİLERİ"
        sec1.Font.Name = "Segoe UI"
        sec1.Font.Size = 11
        sec1.Font.Bold = True
        sec1.Font.Color = rgb_to_int(255, 255, 255)
        sec1.Interior.Color = rgb_to_int(45, 106, 79)
        sec1.HorizontalAlignment = -4131
        ws_input.Rows(5).RowHeight = 24

        fields = [
            (6, "Müşteri / Firma Adı:", "ÖRNEK TARIM İŞLETMESİ A.Ş."),
            (7, "Fide Türü (Ürün):", "DOMATES"),
            (8, "Çeşit / Varyete:", "TORRY F1"),
            (9, "Parti / Lot Numarası:", "LOT-2026/089"),
            (10, "Sipariş / Fatura No:", "SIP-45802"),
            (11, "Tohum / Ekim Tarihi:", "2026-08-21"),
            (12, "Tahmini Teslim / Çıkış Tarihi:", "2026-09-10"),
            (13, "Viyol Tipi (Göz Sayısı):", 216),
            (14, "Paletteki Viyol (Kasa) Adedi:", 15),
            (15, "Palet Başına Toplam Fide:", "=C13*C14"),
            (16, "Aşılama / İlaçlama Durumu:", "AŞILI - KORUYUCU İLAÇLAMA YAPILDI"),
            (17, "Kalite & Kontrol Sorumlusu:", "Zir. Müh. Ahmet YILMAZ")
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
                val_rng.NumberFormat = '#,##0" Fide / Palet"'
                val_rng.Font.Color = rgb_to_int(27, 106, 45)
                val_rng.Interior.Color = rgb_to_int(232, 248, 237)
                val_rng.HorizontalAlignment = -4108

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
        # 2. SETUP PALET_KARTLARI (TEMPLATE SHEET)
        # ----------------------------------------------------
        ws_cards.Tab.Color = rgb_to_int(30, 74, 140)

        # Page Setup for PALET_KARTLARI
        ws_cards.PageSetup.Orientation = 1 # xlPortrait
        ws_cards.PageSetup.PaperSize = 9 # xlPaperA4
        ws_cards.PageSetup.LeftMargin = excel.InchesToPoints(0.4)
        ws_cards.PageSetup.RightMargin = excel.InchesToPoints(0.4)
        ws_cards.PageSetup.TopMargin = excel.InchesToPoints(0.4)
        ws_cards.PageSetup.BottomMargin = excel.InchesToPoints(0.4)
        ws_cards.PageSetup.HeaderMargin = 0
        ws_cards.PageSetup.FooterMargin = 0
        ws_cards.PageSetup.CenterHorizontally = True

        # Columns for Card Layout
        ws_cards.Columns("A").ColumnWidth = 2
        ws_cards.Columns("B").ColumnWidth = 14
        ws_cards.Columns("C").ColumnWidth = 18
        ws_cards.Columns("D").ColumnWidth = 14
        ws_cards.Columns("E").ColumnWidth = 14
        ws_cards.Columns("F").ColumnWidth = 14
        ws_cards.Columns("G").ColumnWidth = 16
        ws_cards.Columns("H").ColumnWidth = 2

        # Populate initial 12 cards
        print("Populating initial 12 cards on PALET_KARTLARI...")
        build_all_cards_py(ws_cards, 12)

        # ----------------------------------------------------
        # 3. ADD VBA MACROS TO WORKBOOK
        # ----------------------------------------------------
        vba_code = """Attribute VB_Name = "Modul_PaletKarti"
Option Explicit

' ==========================================================
' FİDE PALET KARTI OTOMASYON MODÜLÜ
' ==========================================================

Public Sub Makro_KartlariGoster()
    Dim wsInput As Worksheet, wsCards As Worksheet
    Dim totalCards As Long, i As Long
    Dim startRow As Long, cardHeight As Long, cardSpacing As Long
    Dim custName As String, fType As String, variety As String
    Dim lotNo As String, orderNo As String, sowDate As String, delDate As String
    Dim violsPerPallet As Variant, cellCount As Variant, totalSeedlings As Variant
    Dim statusText As String, responsible As String
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wsInput = ThisWorkbook.Sheets("VERI_GIRIS")
    Set wsCards = ThisWorkbook.Sheets("PALET_KARTLARI")
    
    totalCards = wsInput.Range("C20").Value
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
    fType = Trim(wsInput.Range("C7").Value)
    variety = Trim(wsInput.Range("C8").Value)
    lotNo = Trim(wsInput.Range("C9").Value)
    orderNo = Trim(wsInput.Range("C10").Value)
    sowDate = Format(wsInput.Range("C11").Value, "dd.mm.yyyy")
    delDate = Format(wsInput.Range("C12").Value, "dd.mm.yyyy")
    cellCount = wsInput.Range("C13").Value
    violsPerPallet = wsInput.Range("C14").Value
    totalSeedlings = wsInput.Range("C15").Value
    statusText = Trim(wsInput.Range("C16").Value)
    responsible = Trim(wsInput.Range("C17").Value)
    
    wsCards.Cells.Clear
    wsCards.ResetAllPageBreaks
    
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
        Call BuildSingleCard(wsCards, startRow, i, totalCards, custName, fType, variety, lotNo, orderNo, sowDate, delDate, cellCount, violsPerPallet, totalSeedlings, statusText, responsible)
        
        If (i Mod 2 = 0) And (i < totalCards) Then
            wsCards.HPageBreaks.Add Before:=wsCards.Rows(startRow + cardHeight + cardSpacing)
        End If
    Next i
    
    Dim lastCardEndRow As Long
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
                            custName As String, fType As String, variety As String, _
                            lotNo As String, orderNo As String, sowDate As String, delDate As String, _
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
        .Value = "🌱 FİDE PALET TANITIM KARTI"
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
    Call FormatField(ws, r1, 5, 7, "FİDE TÜRÜ (ÜRÜN)", fType, True)
    
    Call FormatField(ws, r2, 2, 4, "ÇEŞİT / VARYETE", variety, True)
    Call FormatField(ws, r2, 5, 7, "TOHUM EKİM TARİHİ", sowDate, False)
    
    Call FormatField(ws, r3, 2, 3, "TESLİM / ÇIKIŞ TARİHİ", delDate, False)
    Call FormatField(ws, r3, 4, 5, "VİYOL GÖZ SAYISI", cellCount & " Gözlü", False)
    Call FormatField(ws, r3, 6, 7, "PALETTEKİ VİYOL ADEDİ", violsPerPallet & " Kasa/Viyol", False)
    
    With ws.Range(ws.Cells(r4, 2), ws.Cells(r4 + 1, 4))
        .Merge
        .Value = "TOPLAM FİDE: " & Format(totalSeedlings, "#,##0") & " ADET"
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(27, 67, 50)
        .Interior.Color = RGB(216, 243, 220)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Call FormatField(ws, r4, 5, 7, "AŞILAMA & İLAÇLAMA DURUMU", statusText, False)
    
    With ws.Range(ws.Cells(r5, 2), ws.Cells(r5 + 1, 4))
        .Merge
        .Value = "SORUMLU: " & responsible
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
    Dim pdfPath As String, timeStamp As String
    
    Call Makro_KartlariGoster
    Set wsCards = ThisWorkbook.Sheets("PALET_KARTLARI")
    
    timeStamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    pdfPath = Environ("USERPROFILE") & "\\Desktop\\Fide_Palet_Kartlari_" & timeStamp & ".pdf"
    
    On Error Resume Next
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
        print("VBA Module Modul_PaletKarti added successfully.")

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

        ws_input.Activate()
        ws_input.Range("C20").Select()

        # Save as target_path (.xlsm)
        wb.SaveAs(target_path, 52)
        print(f"Saved: {target_path}")

        # Also save copy as Book1.xlsm
        wb.SaveCopyAs(original)
        print(f"Main copy updated: {original}")

        wb.Close(False)
        print("Done!")
        return True

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return False
    finally:
        excel.Quit()

if __name__ == "__main__":
    success = create_pallet_workbook()
    if success:
        print("EXCEL CREATION SUCCEEDED!")
    else:
        print("EXCEL CREATION FAILED!")
