'=============================================================================
' MODUL   : POIN_CABANG
' VERSI   : 1.0
' DESKRIPSI: Kalkulator Poin Cabang Wilayah Kepri
'            - 27 Cabang x 11 Kategori x 31 Hari
'            - Auto-build sheet Dashboard & Input per cabang
'            - Hitung poin positif x bobot + penalti nihil (-1)
' CARA PAKAI:
'   1. Buka Excel -> tekan Alt+F11
'   2. Insert -> Module -> Paste seluruh kode ini
'   3. Tutup VBE, kembali ke Excel
'   4. Tekan Alt+F8 -> pilih "BuatSemua" -> Run
'=============================================================================
Option Explicit

'─────────────────────────────────────────────────────────────────────────────
' KONSTANTA & DATA REFERENSI
'─────────────────────────────────────────────────────────────────────────────
Private Const SHEET_DASHBOARD As String = "DASHBOARD"
Private Const SHEET_REKAP     As String = "REKAP BULANAN"

' 27 Cabang
Private Function DaftarCabang() As Variant
    DaftarCabang = Array( _
        "10900 - KC BATAM IMAM BONJOL", _
        "10901 - KCP BATAM LUBUK BAJA", _
        "10902 - KCP BATAM RAJA ALI HAJI", _
        "10903 - KCP BATAM SEKUPANG", _
        "10904 - KCP BATAM INDUSTRIAL PARK", _
        "10905 - KC TANJUNGPINANG", _
        "10906 - KCP TANJUNG UBAN", _
        "10907 - KCP BATAM BANDARA HANG NADIM", _
        "10908 - KCP BATAM CENTER", _
        "10909 - KCP BATAM SP PLAZA", _
        "10910 - KCP BATAM KAWASAN INDUSTRI TUNAS", _
        "10911 - KCP BATAM TIBAN", _
        "10912 - KCP BATAM PANBIL", _
        "10913 - KCP TANJUNG BALAI KARIMUN", _
        "10914 - KCP KIJANG", _
        "10915 - KCP NATUNA", _
        "10916 - KCP BATAM KAWASAN INDUSTRI KABIL", _
        "10917 - KCP BINTAN CENTER", _
        "10918 - KCP BATAM FANINDO", _
        "10919 - KCP BATAM KEPRI MALL", _
        "10920 - KCP BATAM PALM SPRING", _
        "10922 - KCP BATAM BOTANIA", _
        "10924 - KCP BATAM GRAND NIAGA MAS", _
        "10925 - KCP BATAM BATU AMPAR", _
        "10926 - KCP BINTAN ALUMINA INDONESIA", _
        "10977 - KCP TANJUNG BATU", _
        "10980 - KCP BATAM TANJUNG PIAYU" _
    )
End Function

' 11 Kategori  [label, poin]
Private Function DaftarKategori() As Variant
    DaftarKategori = Array( _
        Array("MTB > 25 New CIF - EDC",               8), _
        Array("GIRO > 25 New CIF",                    8), _
        Array("KOPRA / TABREG > 10 / TRM",            4), _
        Array("AXA CC Retail",                        6), _
        Array("HVC > 100 jt",                        10), _
        Array("KSM < 100 / LVM Usaha",                4), _
        Array("New CIF < 25 jt (Gir-Tabis-Tabreg) / GMM", 2), _
        Array("KPR DTBO / KSM > 100 jt / CC Approve", 8), _
        Array("KKB / PKS Mitra ID",                   4), _
        Array("E-Commerce / NTP",                     8), _
        Array("Livin USAK / Payroll PMP",              1) _
    )
End Function

'─────────────────────────────────────────────────────────────────────────────
' ENTRY POINT — Jalankan macro ini untuk membangun semua sheet sekaligus
'─────────────────────────────────────────────────────────────────────────────
Sub BuatSemua()
    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationManual

    BuatSheetInput
    BuatSheetDashboard
    BuatSheetRekap

    Application.Calculation   = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Sheets(SHEET_DASHBOARD).Activate
    MsgBox "Selesai! Sheet berhasil dibuat:" & vbNewLine & _
           "  - " & SHEET_DASHBOARD & vbNewLine & _
           "  - " & SHEET_REKAP & vbNewLine & _
           "  - 27 sheet input cabang", _
           vbInformation, "Poin Cabang Kepri"
End Sub

'─────────────────────────────────────────────────────────────────────────────
' 1.  SHEET INPUT PER CABANG  (27 sheet, nama = kode cabang)
'─────────────────────────────────────────────────────────────────────────────
Sub BuatSheetInput()
    Dim cabang   As Variant : cabang   = DaftarCabang()
    Dim kategori As Variant : kategori = DaftarKategori()
    Dim ws       As Worksheet
    Dim i As Integer, d As Integer, k As Integer
    Dim r As Long, c As Long
    Dim shName As String
    Dim maxPoin As Integer : maxPoin = 0
    For k = 0 To UBound(kategori)
        maxPoin = maxPoin + kategori(k)(1)
    Next k

    For i = 0 To UBound(cabang)
        ' Nama sheet = kode 5 digit
        shName = Left(cabang(i), 5)

        ' Hapus sheet lama jika ada
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(shName).Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = shName

        '── Header cabang ────────────────────────────────────────────
        With ws
            .Tab.Color = RGB(0, 70, 127)          ' biru Mandiri

            ' Judul
            .Cells(1, 1).Value = cabang(i)
            With .Range("A1:AJ1")
                .Merge
                .Font.Bold = True
                .Font.Size = 12
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(0, 70, 127)
                .HorizontalAlignment = xlCenter
            End With

            ' Baris 2 - sub header
            .Cells(2, 1).Value = "INPUT NILAI HARIAN PER KATEGORI"
            With .Range("A2:AJ2")
                .Merge
                .Font.Italic = True
                .Font.Size = 9
                .Font.Color = RGB(200, 220, 255)
                .Interior.Color = RGB(0, 100, 160)
                .HorizontalAlignment = xlCenter
            End With

            ' ── Header kolom (baris 4) ──────────────────────────────
            .Cells(4, 1).Value = "TGL"
            .Cells(4, 1).Font.Bold = True
            .Cells(4, 1).Interior.Color = RGB(0, 70, 127)
            .Cells(4, 1).Font.Color = RGB(255, 255, 255)

            For k = 0 To UBound(kategori)
                c = k + 2
                .Cells(4, c).Value = kategori(k)(0)
                .Cells(4, c).Font.Bold = True
                .Cells(4, c).Interior.Color = RGB(0, 70, 127)
                .Cells(4, c).Font.Color = RGB(255, 255, 255)
                .Cells(4, c).WrapText = True

                ' Bobot di baris 5
                .Cells(5, c).Value = kategori(k)(1)
                .Cells(5, c).Font.Bold = True
                .Cells(5, c).Interior.Color = RGB(173, 216, 230)
                .Cells(5, c).HorizontalAlignment = xlCenter
            Next k

            ' Kolom tambahan: Total+, Total-, Saldo
            c = UBound(kategori) + 3
            .Cells(4, c).Value     = "TOTAL +"
            .Cells(4, c + 1).Value = "TOTAL -"
            .Cells(4, c + 2).Value = "SALDO"
            .Cells(5, c).Value     = "Positif"
            .Cells(5, c + 1).Value = "Penalti"
            .Cells(5, c + 2).Value = "Bersih"

            For Each cl In .Range(.Cells(4, c), .Cells(5, c + 2))
                cl.Font.Bold = True
                cl.Interior.Color = RGB(0, 90, 140)
                cl.Font.Color = RGB(255, 255, 255)
                cl.HorizontalAlignment = xlCenter
            Next cl

            ' Baris keterangan bobot
            .Cells(3, 1).Value = "Bobot:"
            .Cells(3, 1).Font.Italic = True
            .Cells(3, 1).Font.Size = 8

            ' ── Baris 31 tanggal (mulai baris 6) ───────────────────
            For d = 1 To 31
                r = d + 5   ' baris 6 = tgl 1

                ' Kolom A = tanggal
                .Cells(r, 1).Value = d
                .Cells(r, 1).Font.Bold = True
                .Cells(r, 1).HorizontalAlignment = xlCenter
                If d Mod 2 = 0 Then
                    .Cells(r, 1).Interior.Color = RGB(240, 248, 255)
                Else
                    .Cells(r, 1).Interior.Color = RGB(255, 255, 255)
                End If

                ' Kolom B..L = input nilai (kosong, diisi user)
                For k = 0 To UBound(kategori)
                    c = k + 2
                    .Cells(r, c).Value = ""
                    .Cells(r, c).HorizontalAlignment = xlCenter
                    .Cells(r, c).NumberFormat = "0"
                    If d Mod 2 = 0 Then
                        .Cells(r, c).Interior.Color = RGB(240, 248, 255)
                    Else
                        .Cells(r, c).Interior.Color = RGB(255, 255, 255)
                    End If
                Next k

                ' ── Kolom Total+ : jumlahkan nilai*bobot jika > 0 ──
                Dim colStart As Integer : colStart = 2
                Dim colEnd   As Integer : colEnd   = UBound(kategori) + 2
                c = UBound(kategori) + 3

                ' Total positif: SUMPRODUCT dari setiap nilai*bobot jika nilai>0
                Dim fPos As String, fNeg As String
                fPos = "SUMPRODUCT(("
                fNeg = "SUMPRODUCT(("
                For k = 0 To UBound(kategori)
                    Dim colLetter As String
                    colLetter = ColLetter(k + 2)
                    If k > 0 Then
                        fPos = fPos & "+"
                        fNeg = fNeg & "+"
                    End If
                    fPos = fPos & "IF(" & colLetter & r & ">0," & colLetter & r & "*" & kategori(k)(1) & ",0)"
                    fNeg = fNeg & "IF(" & colLetter & r & "<1,IF(" & colLetter & r & "<>""""",-1,0),0)"
                Next k
                fPos = fPos & "))"
                fNeg = fNeg & "))"

                .Cells(r, c).Formula     = "=" & fPos
                .Cells(r, c + 1).Formula = "=" & fNeg
                .Cells(r, c + 2).Formula = "=" & ColLetter(c) & r & "+" & ColLetter(c + 1) & r

                ' Warna kolom total
                .Cells(r, c).Interior.Color     = RGB(220, 255, 220)
                .Cells(r, c + 1).Interior.Color = RGB(255, 220, 220)
                .Cells(r, c + 2).Interior.Color = RGB(230, 230, 255)
                .Cells(r, c).Font.Bold     = True
                .Cells(r, c + 1).Font.Bold = True
                .Cells(r, c + 2).Font.Bold = True
            Next d

            ' ── Baris TOTAL di bawah (baris 37) ─────────────────────
            r = 38
            .Cells(r, 1).Value = "TOTAL"
            .Cells(r, 1).Font.Bold = True
            .Cells(r, 1).Interior.Color = RGB(0, 70, 127)
            .Cells(r, 1).Font.Color = RGB(255, 255, 255)
            .Cells(r, 1).HorizontalAlignment = xlCenter

            For k = 0 To UBound(kategori)
                c = k + 2
                .Cells(r, c).Formula = "=SUM(" & ColLetter(c) & "6:" & ColLetter(c) & "36)"
                .Cells(r, c).Font.Bold = True
                .Cells(r, c).Interior.Color = RGB(0, 70, 127)
                .Cells(r, c).Font.Color = RGB(255, 255, 255)
                .Cells(r, c).HorizontalAlignment = xlCenter
            Next k

            c = UBound(kategori) + 3
            .Cells(r, c).Formula     = "=SUM(" & ColLetter(c) & "6:" & ColLetter(c) & "36)"
            .Cells(r, c + 1).Formula = "=SUM(" & ColLetter(c + 1) & "6:" & ColLetter(c + 1) & "36)"
            .Cells(r, c + 2).Formula = "=SUM(" & ColLetter(c + 2) & "6:" & ColLetter(c + 2) & "36)"

            For Each cl In .Range(.Cells(r, c), .Cells(r, c + 2))
                cl.Font.Bold = True
                cl.Interior.Color = RGB(0, 70, 127)
                cl.Font.Color = RGB(255, 255, 255)
                cl.HorizontalAlignment = xlCenter
            Next cl

            ' ── Format kolom ────────────────────────────────────────
            .Columns("A").ColumnWidth = 5
            For k = 0 To UBound(kategori)
                .Columns(k + 2).ColumnWidth = 9
            Next k
            .Columns(c).ColumnWidth     = 10
            .Columns(c + 1).ColumnWidth = 10
            .Columns(c + 2).ColumnWidth = 10

            ' Row height header
            .Rows(4).RowHeight = 45
            .Rows(5).RowHeight = 16

            ' Freeze panes
            .Cells(6, 2).Select
            ActiveWindow.FreezePanes = True

            ' Border grid
            With .Range(.Cells(4, 1), .Cells(38, c + 2)).Borders
                .LineStyle = xlContinuous
                .Color = RGB(200, 200, 200)
                .Weight = xlThin
            End With
        End With
    Next i
End Sub

'─────────────────────────────────────────────────────────────────────────────
' 2.  SHEET DASHBOARD  (ringkasan semua cabang per hari/bulan)
'─────────────────────────────────────────────────────────────────────────────
Sub BuatSheetDashboard()
    Dim cabang   As Variant : cabang   = DaftarCabang()
    Dim kategori As Variant : kategori = DaftarKategori()
    Dim ws       As Worksheet
    Dim i As Integer, d As Integer, r As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(SHEET_DASHBOARD).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = Sheets.Add(Before:=Sheets(1))
    ws.Name = SHEET_DASHBOARD
    ws.Tab.Color = RGB(0, 150, 100)

    With ws
        ' ── Judul ────────────────────────────────────────────────────
        .Cells(1, 1).Value = "DASHBOARD POIN CABANG — WILAYAH KEPRI"
        With .Range("A1:AF1")
            .Merge
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 100, 70)
            .HorizontalAlignment = xlCenter
            .RowHeight = 30
        End With

        .Cells(2, 1).Value = "Saldo = Total Positif + Total Penalti Nihil  |  Positif: nilai > 0 dikali bobot  |  Nihil: setiap kategori kosong/0 = -1"
        With .Range("A2:AF2")
            .Merge
            .Font.Italic = True
            .Font.Size = 9
            .Font.Color = RGB(180, 230, 210)
            .Interior.Color = RGB(0, 130, 90)
            .HorizontalAlignment = xlCenter
        End With

        ' ── Header kolom: TGL | C1 | C2 ... C27 | AVG ───────────────
        .Cells(4, 1).Value = "TANGGAL"
        .Cells(4, 1).Interior.Color = RGB(0, 100, 70)
        .Cells(4, 1).Font.Color = RGB(255, 255, 255)
        .Cells(4, 1).Font.Bold = True
        .Cells(4, 1).HorizontalAlignment = xlCenter

        For i = 0 To UBound(cabang)
            Dim shName As String
            shName = Left(cabang(i), 5)
            .Cells(4, i + 2).Value = shName
            .Cells(5, i + 2).Value = Mid(cabang(i), 9)   ' nama tanpa kode
            .Cells(4, i + 2).Interior.Color = RGB(0, 100, 70)
            .Cells(4, i + 2).Font.Color = RGB(255, 255, 255)
            .Cells(4, i + 2).Font.Bold = True
            .Cells(4, i + 2).HorizontalAlignment = xlCenter
            .Cells(5, i + 2).Interior.Color = RGB(0, 130, 90)
            .Cells(5, i + 2).Font.Color = RGB(200, 240, 220)
            .Cells(5, i + 2).Font.Size = 8
            .Cells(5, i + 2).WrapText = True
            .Columns(i + 2).ColumnWidth = 11
        Next i

        ' Kolom Rata-rata
        Dim colAvg As Integer : colAvg = UBound(cabang) + 3
        .Cells(4, colAvg).Value = "RATA-RATA"
        .Cells(5, colAvg).Value = "Tim"
        .Cells(4, colAvg).Interior.Color = RGB(0, 70, 50)
        .Cells(4, colAvg).Font.Color = RGB(255, 255, 255)
        .Cells(4, colAvg).Font.Bold = True
        .Cells(5, colAvg).Interior.Color = RGB(0, 70, 50)
        .Cells(5, colAvg).Font.Color = RGB(200, 240, 220)
        .Cells(5, colAvg).HorizontalAlignment = xlCenter
        .Columns(colAvg).ColumnWidth = 12

        .Cells(4, 1).RowHeight = 30
        .Rows(5).RowHeight = 40

        ' ── 31 baris tanggal ─────────────────────────────────────────
        For d = 1 To 31
            r = d + 5

            .Cells(r, 1).Value = d
            .Cells(r, 1).Font.Bold = True
            .Cells(r, 1).HorizontalAlignment = xlCenter
            If d Mod 2 = 0 Then
                .Cells(r, 1).Interior.Color = RGB(235, 250, 242)
            Else
                .Cells(r, 1).Interior.Color = RGB(255, 255, 255)
            End If

            For i = 0 To UBound(cabang)
                shName = Left(cabang(i), 5)
                Dim colIdx As Integer : colIdx = i + 2
                ' Referensi ke sheet cabang kolom "SALDO" (kolom N = 14)
                ' Kolom SALDO di sheet input = kolom 14 (N), baris d+5
                .Cells(r, colIdx).Formula = "=IF(ISBLANK('" & shName & "'!" & ColLetter(14) & (d + 5) & _
                                             "),0,'" & shName & "'!" & ColLetter(14) & (d + 5) & ")"
                .Cells(r, colIdx).NumberFormat = "+0;-0;0"
                .Cells(r, colIdx).HorizontalAlignment = xlCenter
                If d Mod 2 = 0 Then
                    .Cells(r, colIdx).Interior.Color = RGB(235, 250, 242)
                Else
                    .Cells(r, colIdx).Interior.Color = RGB(255, 255, 255)
                End If
            Next i

            ' Rata-rata tim
            .Cells(r, colAvg).Formula = "=IFERROR(AVERAGE(" & ColLetter(2) & r & ":" & ColLetter(UBound(cabang) + 2) & r & "),0)"
            .Cells(r, colAvg).NumberFormat = "+0.0;-0.0;0"
            .Cells(r, colAvg).Font.Bold = True
            .Cells(r, colAvg).HorizontalAlignment = xlCenter
            If d Mod 2 = 0 Then
                .Cells(r, colAvg).Interior.Color = RGB(200, 240, 220)
            Else
                .Cells(r, colAvg).Interior.Color = RGB(220, 250, 235)
            End If
        Next d

        ' ── Baris TOTAL BULAN (baris 37) ─────────────────────────────
        r = 38
        .Cells(r, 1).Value = "TOTAL BULAN"
        .Cells(r, 1).Font.Bold = True
        .Cells(r, 1).Interior.Color = RGB(0, 60, 40)
        .Cells(r, 1).Font.Color = RGB(255, 255, 255)
        .Cells(r, 1).HorizontalAlignment = xlCenter

        For i = 0 To UBound(cabang)
            Dim colI As Integer : colI = i + 2
            .Cells(r, colI).Formula = "=SUM(" & ColLetter(colI) & "6:" & ColLetter(colI) & "36)"
            .Cells(r, colI).Font.Bold = True
            .Cells(r, colI).NumberFormat = "+0;-0;0"
            .Cells(r, colI).Interior.Color = RGB(0, 60, 40)
            .Cells(r, colI).Font.Color = RGB(255, 255, 255)
            .Cells(r, colI).HorizontalAlignment = xlCenter
        Next i

        .Cells(r, colAvg).Formula = "=AVERAGE(" & ColLetter(2) & r & ":" & ColLetter(UBound(cabang) + 2) & r & ")"
        .Cells(r, colAvg).Font.Bold = True
        .Cells(r, colAvg).NumberFormat = "+0.0;-0.0;0"
        .Cells(r, colAvg).Interior.Color = RGB(0, 60, 40)
        .Cells(r, colAvg).Font.Color = RGB(255, 255, 255)
        .Cells(r, colAvg).HorizontalAlignment = xlCenter

        ' ── RANKING — baris 40 ───────────────────────────────────────
        .Cells(40, 1).Value = "RANKING CABANG (Bulan Ini)"
        With .Range(.Cells(40, 1), .Cells(40, colAvg))
            .Merge
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(0, 100, 70)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        .Cells(41, 1).Value = "RANK"
        .Cells(41, 2).Value = "KODE"
        .Cells(41, 3).Value = "NAMA CABANG"
        .Cells(41, 4).Value = "TOTAL POIN"
        .Cells(41, 5).Value = "POSISI"

        For Each cl In .Range("A41:E41")
            cl.Font.Bold = True
            cl.Interior.Color = RGB(0, 130, 90)
            cl.Font.Color = RGB(255, 255, 255)
            cl.HorizontalAlignment = xlCenter
        Next cl

        ' Ranking menggunakan formula LARGE + INDEX/MATCH
        For i = 1 To 27
            r = 41 + i
            ' RANK
            .Cells(r, 1).Value = i
            .Cells(r, 1).HorizontalAlignment = xlCenter
            .Cells(r, 1).Font.Bold = True

            ' Nama cabang - LARGE dari total baris 38
            ' Gunakan formula untuk ambil nilai terbesar ke-i dari range total cabang
            Dim rankRef As String
            rankRef = "$B$38:$" & ColLetter(UBound(cabang) + 2) & "$38"
            .Cells(r, 2).Formula = "=IFERROR(INDEX($B$4:$" & ColLetter(UBound(cabang) + 2) & "$4," & _
                                   "MATCH(LARGE(" & rankRef & "," & i & ")," & rankRef & ",0)),"""")"
            .Cells(r, 3).Formula = "=IFERROR(INDEX($B$5:$" & ColLetter(UBound(cabang) + 2) & "$5," & _
                                   "MATCH(LARGE(" & rankRef & "," & i & ")," & rankRef & ",0)),"""")"
            .Cells(r, 4).Formula = "=IFERROR(LARGE(" & rankRef & "," & i & "),0)"
            .Cells(r, 4).NumberFormat = "+0;-0;0"
            .Cells(r, 4).Font.Bold = True
            .Cells(r, 4).HorizontalAlignment = xlCenter

            ' Warna medali
            Select Case i
                Case 1
                    .Cells(r, 5).Value = "Terbaik"
                    .Cells(r, 5).Font.Color = RGB(184, 134, 11)
                    .Cells(r, 1).Interior.Color = RGB(255, 215, 0)
                Case 2
                    .Cells(r, 5).Value = "2"
                    .Cells(r, 1).Interior.Color = RGB(220, 220, 220)
                Case 3
                    .Cells(r, 5).Value = "3"
                    .Cells(r, 1).Interior.Color = RGB(205, 127, 50)
                    .Cells(r, 1).Font.Color = RGB(255, 255, 255)
                Case Is >= 25
                    .Cells(r, 5).Value = "Perlu Perhatian"
                    .Cells(r, 5).Font.Color = RGB(200, 0, 0)
                Case Else
                    .Cells(r, 5).Value = "-"
            End Select

            .Cells(r, 1).HorizontalAlignment = xlCenter
            .Cells(r, 3).ColumnWidth = 30

            If i Mod 2 = 0 Then
                .Range(.Cells(r, 2), .Cells(r, 5)).Interior.Color = RGB(240, 252, 246)
            End If
        Next i

        .Columns(1).ColumnWidth = 12
        .Columns(3).ColumnWidth = 32

        ' Conditional formatting saldo hari — merah jika negatif
        Dim cfRange As Range
        Set cfRange = .Range("B6:" & ColLetter(UBound(cabang) + 2) & "36")
        cfRange.FormatConditions.Delete

        ' Negatif = merah muda
        With cfRange.FormatConditions.Add(xlCellValue, xlLess, 0)
            .Interior.Color = RGB(255, 200, 200)
            .Font.Color = RGB(180, 0, 0)
            .Font.Bold = True
        End With

        ' Positif besar = hijau
        With cfRange.FormatConditions.Add(xlCellValue, xlGreater, 0)
            .Interior.Color = RGB(200, 255, 200)
            .Font.Color = RGB(0, 130, 0)
            .Font.Bold = True
        End With

        ' Freeze
        .Cells(6, 2).Select
        ActiveWindow.FreezePanes = True

        ' Border
        With .Range(.Cells(4, 1), .Cells(38, colAvg)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(180, 220, 200)
            .Weight = xlThin
        End With
    End With
End Sub

'─────────────────────────────────────────────────────────────────────────────
' 3.  SHEET REKAP BULANAN  (detail skor per cabang: positif, penalti, saldo)
'─────────────────────────────────────────────────────────────────────────────
Sub BuatSheetRekap()
    Dim cabang   As Variant : cabang   = DaftarCabang()
    Dim kategori As Variant : kategori = DaftarKategori()
    Dim ws       As Worksheet
    Dim i As Integer, k As Integer, r As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(SHEET_REKAP).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = Sheets.Add(After:=Sheets(SHEET_DASHBOARD))
    ws.Name = SHEET_REKAP
    ws.Tab.Color = RGB(150, 50, 0)

    With ws
        ' Judul
        .Cells(1, 1).Value = "REKAP BULANAN — POIN PER KATEGORI PER CABANG"
        With .Range("A1:Q1")
            .Merge
            .Font.Bold = True
            .Font.Size = 13
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(150, 50, 0)
            .HorizontalAlignment = xlCenter
            .RowHeight = 28
        End With

        ' Header
        .Cells(3, 1).Value = "NO"
        .Cells(3, 2).Value = "KODE"
        .Cells(3, 3).Value = "NAMA CABANG"
        Dim colOff As Integer : colOff = 4

        For k = 0 To UBound(kategori)
            .Cells(3, colOff + k).Value = kategori(k)(0)
            .Cells(4, colOff + k).Value = "Bobot " & kategori(k)(1)
        Next k

        .Cells(3, colOff + UBound(kategori) + 1).Value = "TOTAL +"
        .Cells(3, colOff + UBound(kategori) + 2).Value = "TOTAL -"
        .Cells(3, colOff + UBound(kategori) + 3).Value = "SALDO BULAN"
        .Cells(4, colOff + UBound(kategori) + 1).Value = "Positif"
        .Cells(4, colOff + UBound(kategori) + 2).Value = "Penalti"
        .Cells(4, colOff + UBound(kategori) + 3).Value = "Bersih"

        For Each cl In .Range(.Cells(3, 1), .Cells(4, colOff + UBound(kategori) + 3))
            cl.Font.Bold = True
            cl.Interior.Color = RGB(150, 50, 0)
            cl.Font.Color = RGB(255, 255, 255)
            cl.HorizontalAlignment = xlCenter
            cl.WrapText = True
        Next cl
        .Rows(3).RowHeight = 45
        .Rows(4).RowHeight = 16

        ' Data per cabang
        For i = 0 To UBound(cabang)
            r = i + 5
            Dim shN As String : shN = Left(cabang(i), 5)

            .Cells(r, 1).Value = i + 1
            .Cells(r, 2).Value = Left(cabang(i), 5)
            .Cells(r, 3).Value = Mid(cabang(i), 9)

            For k = 0 To UBound(kategori)
                ' Ambil total dari baris 38 (TOTAL) sheet cabang, kolom k+2
                .Cells(r, colOff + k).Formula = "='" & shN & "'!" & ColLetter(k + 2) & "38"
                .Cells(r, colOff + k).HorizontalAlignment = xlCenter
                .Cells(r, colOff + k).NumberFormat = "0"
            Next k

            ' Total positif bulan (kolom 12 di sheet cabang = kolom L)
            .Cells(r, colOff + UBound(kategori) + 1).Formula = "='" & shN & "'!" & ColLetter(12) & "38"
            ' Total penalti bulan (kolom 13 = M)
            .Cells(r, colOff + UBound(kategori) + 2).Formula = "='" & shN & "'!" & ColLetter(13) & "38"
            ' Saldo bulan (kolom 14 = N)
            .Cells(r, colOff + UBound(kategori) + 3).Formula = "='" & shN & "'!" & ColLetter(14) & "38"

            ' Format warna total kolom
            .Cells(r, colOff + UBound(kategori) + 1).Interior.Color = RGB(220, 255, 220)
            .Cells(r, colOff + UBound(kategori) + 2).Interior.Color = RGB(255, 220, 220)
            .Cells(r, colOff + UBound(kategori) + 3).Interior.Color = RGB(230, 220, 255)
            .Cells(r, colOff + UBound(kategori) + 1).Font.Bold = True
            .Cells(r, colOff + UBound(kategori) + 2).Font.Bold = True
            .Cells(r, colOff + UBound(kategori) + 3).Font.Bold = True
            .Cells(r, colOff + UBound(kategori) + 3).NumberFormat = "+0;-0;0"

            If i Mod 2 = 0 Then
                .Range(.Cells(r, 1), .Cells(r, 3)).Interior.Color = RGB(255, 245, 240)
            End If
        Next i

        ' Baris grand total
        r = UBound(cabang) + 6
        .Cells(r, 1).Value = "TOTAL"
        With .Range(.Cells(r, 1), .Cells(r, 3))
            .Merge
            .Font.Bold = True
            .Interior.Color = RGB(100, 30, 0)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        For k = 0 To UBound(kategori) + 3
            Dim tcol As Integer : tcol = colOff + k
            .Cells(r, tcol).Formula = "=SUM(" & ColLetter(tcol) & "5:" & ColLetter(tcol) & (r - 1) & ")"
            .Cells(r, tcol).Font.Bold = True
            .Cells(r, tcol).Interior.Color = RGB(100, 30, 0)
            .Cells(r, tcol).Font.Color = RGB(255, 255, 255)
            .Cells(r, tcol).HorizontalAlignment = xlCenter
        Next k

        ' Conditional formatting saldo bersih
        Dim cfR As Range
        Set cfR = .Range(ColLetter(colOff + UBound(kategori) + 3) & "5:" & _
                         ColLetter(colOff + UBound(kategori) + 3) & (r - 1))
        cfR.FormatConditions.Delete
        With cfR.FormatConditions.Add(xlCellValue, xlLess, 0)
            .Interior.Color = RGB(255, 200, 200)
            .Font.Color = RGB(180, 0, 0)
        End With
        With cfR.FormatConditions.Add(xlCellValue, xlGreater, 0)
            .Interior.Color = RGB(200, 255, 200)
            .Font.Color = RGB(0, 130, 0)
        End With

        ' Lebar kolom
        .Columns(1).ColumnWidth = 5
        .Columns(2).ColumnWidth = 8
        .Columns(3).ColumnWidth = 32
        For k = 0 To UBound(kategori) + 3
            .Columns(colOff + k).ColumnWidth = 10
        Next k

        ' Freeze & border
        .Cells(5, 4).Select
        ActiveWindow.FreezePanes = True
        With .Range(.Cells(3, 1), .Cells(r, colOff + UBound(kategori) + 3)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 180, 170)
            .Weight = xlThin
        End With
    End With
End Sub

'─────────────────────────────────────────────────────────────────────────────
' UTILITAS — Konversi nomor kolom ke huruf (1=A, 26=Z, 27=AA, dst.)
'─────────────────────────────────────────────────────────────────────────────
Private Function ColLetter(ByVal colNum As Integer) As String
    Dim result As String : result = ""
    Do While colNum > 0
        Dim remainder As Integer
        remainder = (colNum - 1) Mod 26
        result    = Chr(65 + remainder) & result
        colNum    = (colNum - 1) \ 26
    Loop
    ColLetter = result
End Function

'─────────────────────────────────────────────────────────────────────────────
' MACRO TAMBAHAN: Refresh semua formula di semua sheet sekaligus
'─────────────────────────────────────────────────────────────────────────────
Sub RefreshSemua()
    Application.CalculateFull
    MsgBox "Semua formula telah diperbarui.", vbInformation, "Refresh"
End Sub

'─────────────────────────────────────────────────────────────────────────────
' MACRO TAMBAHAN: Bersihkan input harian cabang tertentu
'─────────────────────────────────────────────────────────────────────────────
Sub BersihkanInputCabang()
    Dim shName As String
    shName = InputBox("Masukkan kode cabang (contoh: 10900):", "Bersihkan Input")
    If shName = "" Then Exit Sub

    On Error GoTo tdk_ada
    Dim ws As Worksheet
    Set ws = Sheets(shName)

    If MsgBox("Yakin hapus semua input di sheet " & shName & "?", _
              vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
        ws.Range("B6:L36").ClearContents
        MsgBox "Input cabang " & shName & " telah dibersihkan.", vbInformation
    End If
    Exit Sub

tdk_ada:
    MsgBox "Sheet dengan kode '" & shName & "' tidak ditemukan.", vbExclamation
End Sub
