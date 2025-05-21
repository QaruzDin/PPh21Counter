Attribute VB_Name = "PPhTERFunc"
Option Explicit

Dim colGaji As String
Dim fileDial As Boolean


Public Sub formatting_PPh21TER()
    Dim ws As Worksheet
    Dim rng As Range
    Dim confirm As VbMsgBoxResult
    
    ' Memastikan sheet aktif adalah sheet yang diinginkan
    confirm = MsgBox("Apakah anda ingin menjalankan modul?" & vbCrLf & _
                    "Pastikan kolom PTKP berada tepat di sisi kiri kolom penerimaan bruto!", vbYesNo)
    If confirm = vbNo Then Exit Sub
    
    Import_DataTER
    
    ' penanganan pembatalan import oleh user
    If Not fileDial Then
        confirm = MsgBox("Apakah anda ingin membatalkan modul?", vbYesNo)
        If confirm = vbYes Then
            Exit Sub
        Else
            formatting_PPh21TER
            Exit Sub
        End If
    End If

    colGaji = inputColGaji()
    
    ' Penanganan pembatalan module oleh user
    If colGaji = "" Then MsgBox "Tidak ada kolom gaji yang dimasukkan. Module dibatalkan.", vbExclamation: Exit Sub
    
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range(colGaji & "1")

    With rng
        .Offset(0, 1) = "TER"
        .Offset(0, 2) = "Tarif"
        .Offset(0, 3) = "PPh 21"
        .Offset(0, -1).Copy
        .Offset(0, 1).Resize(1, 3).PasteSpecial Paste:=xlPasteFormats
        .Offset(0, 1).Resize(1, 3).HorizontalAlignment = xlCenter
    End With
    
    iterratingCell
    
    sumPPH21
    
End Sub

Function cariTER(PTKP As String) As Variant
    Dim lookup_table As Variant
    Dim i As Integer

    ' Tabel Kategori TER
    lookup_table = Array( _
        Array("TK/0", "A"), _
        Array("TK/1", "A"), _
        Array("TK/2", "B"), _
        Array("TK/3", "B"), _
        Array("K/0", "A"), _
        Array("K/1", "B"), _
        Array("K/2", "B"), _
        Array("K/3", "C") _
    )

    ' Ulangi tabel pencarian untuk menemukan nilai yang cocok
    For i = LBound(lookup_table) To UBound(lookup_table)
        If lookup_table(i)(0) = PTKP Then
            ' Memperoleh nilai kategori TER sesuai kolom pertama (PTKP)
             cariTER = lookup_table(i)(1)
             Exit Function
        End If
    Next i

    ' Jika tidak ditemukan kecocokan, mengembalikan nilai "Invalid"
    cariTER = "Invalid"
End Function

Public Sub Import_DataTER()
    Dim fileDialog As fileDialog
    Dim selectedFile As String
    Dim wbSc As Workbook ' Workbook Souce
    Dim wsSc As Worksheet '  Worksheet Source
    Dim wbTg As Workbook ' Workbook Target
    Dim wsTg As Worksheet ' Worksheet Target
    
    ' Set Workbook penerima (Target)
    Set wbTg = ThisWorkbook
    
    ' Dialog file untuk memilih file DATA TER
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Pilih file DATA TER"
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        .AllowMultiSelect = False
        fileDial = .Show
        If Not fileDial Then Exit Sub ' pembatalan oleh user
        selectedFile = .SelectedItems(1)
    End With
    
    Set wbSc = Workbooks.Open(selectedFile)
    
    For Each wsSc In wbSc.Sheets
        wsSc.Copy after:=wbTg.Sheets(wbTg.Sheets.Count)
        Set wsTg = wbTg.Sheets(wbTg.Sheets.Count)
        
        ' Penanganan bila DATA TER telah tersedia
        If Not wsTg.Name = wsSc.Name Then
            MsgBox "Data TER yang lama akan ditimpa dengan berkas terbaru!", vbExclamation, "Duplikasi Data TER Terdeksi"
            Application.DisplayAlerts = False
            wsTg.Delete ' Hapus worksheet yang baru dicopy
            wbTg.Sheets(wsSc.Name).Delete ' Hapus worksheet lama dengan nama yang sama
            Application.DisplayAlerts = True
            wsSc.Copy after:=wbTg.Sheets(wbTg.Sheets.Count) ' Copy lagi setelah dihapus
            wbTg.Sheets(wbTg.Sheets.Count).Name = wsSc.Name
        End If
        
    Next wsSc
    
    wbSc.Close False
    
    wbTg.Sheets(ActiveSheet.Index - 1).Activate
    MsgBox "Sheet berhasil disalin ke workbook ini :)", vbInformation

End Sub


Function tarifTER(TER As String, gajiBruto As Double) As Double
    Dim kolomTER As Range
    Dim batasBawah As Range
    Dim lo As ListObject
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("DATA TER") ' pastikan nama sheet telah sesuai
    On Error GoTo 0
    
    ' memeriksa ketersediaan DATA TER
    If ws Is Nothing Then
        MsgBox "Sheet 'DATA TER' tidak ditemukan. Mohon periksa kembali apakah sheet sudah ter-upload."
        Exit Function
    End If
    
    ' case conditional untuk penentuan kategori TER
    Select Case TER
        Case "A"
            Set lo = ws.ListObjects("tabelA")
            Set batasBawah = lo.ListColumns("Batas Bawah").DataBodyRange
            Set kolomTER = lo.ListColumns("TER").DataBodyRange
        Case "B"
            Set lo = ws.ListObjects("tabelB")
            Set batasBawah = lo.ListColumns("Batas Bawah").DataBodyRange
            Set kolomTER = lo.ListColumns("TER").DataBodyRange
        Case "C"
            Set lo = ws.ListObjects("tabelC")
            Set batasBawah = lo.ListColumns("Batas Bawah").DataBodyRange
            Set kolomTER = lo.ListColumns("TER").DataBodyRange
        Case Else
            Exit Function
    End Select
            
    ' pencarian tarif TER
    tarifTER = Application.WorksheetFunction.Index(kolomTER, Application.WorksheetFunction.Match(gajiBruto, batasBawah, 1))
    
End Function


Function PPH21TER(trf As Double, gajiBrt As Double) As Double
  
  PPH21TER = WorksheetFunction.Round(trf * gajiBrt, 0)
End Function


Public Sub iterratingCell()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastcellVal As Variant
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim startTime As Single
    Dim timeout As Single
    
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range(colGaji & "1").Offset(0, 1)
    
    ' mendeteksi kolom aktif terakhir
    lastRow = ws.Cells(ws.Rows.Count, colGaji).End(xlUp).row - 1
    
    ' pengaturan batas waktu untuk pencegahan infinite loop (batas waktu dalam detik)
    startTime = Timer
    timeout = 60
    
    'iterrating one cell at the time
    For i = 1 To lastRow
    
    lastcellVal = ws.Cells(i + 1, 1).Value
    ' mengabaikan kolom jumlah/total yang dibuat user
    If lastcellVal = "Total" Or lastcellVal = "Jumlah" Then
        Exit For
    End If
        For j = 1 To 3
            Set cell = rng.Offset(i, j - 1) ' will be moving one cell above
        
            cell.Select
            Select Case j
                Case 1
                    cell.Formula = "=cariTER(" & cell.Offset(0, -2).Address(False, True) & ")"
                    cell.HorizontalAlignment = xlCenter ' formatting : centered
                Case 2
                    cell.Formula = "=tarifTER(" & cell.Offset(0, -1).Address(False, True) & ", " & _
                                    cell.Offset(0, -2).Address(False, True) & ")"
                    cell.NumberFormat = "0.00%" ' formatting : percentage
                Case 3
                    cell.Formula = "=PPH21TER(" & cell.Offset(0, -1).Address(False, True) & ", " & _
                                    cell.Offset(0, -3).Address(False, True) & ")"
                    cell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)" ' formatting : IDR
            End Select
            
        Next j

        ' Periksa apakah durasi batas waktu telah terlampaui
        If Timer - startTime > timeout Then
            MsgBox "Time limit is reached out"
            Exit Sub
        End If
    Next i
    
    MsgBox "Module perhitungan berhasil dijalankan! Huray!"
End Sub

Public Function inputColGaji() As String
    Dim colnumbTER As Integer
    Dim abortMsg As VbMsgBoxResult
    
    Do
        inputColGaji = InputBox("Mohon input letak kolom gaji bruto anda :" & vbCrLf & _
                    "(Pastikan Kolom PTKP tepat berada disisi kiri kolom gaji bruto.)", "Input Kolom Gaji")
        If inputColGaji = vbNullString Then
            abortMsg = MsgBox("Apakah anda yakin ingin mengakhiri modul?", vbExclamation + vbYesNo)
            If abortMsg = vbYes Then
                Exit Function
            End If
        End If
        
        If Not IsNumeric(inputColGaji) Then
            On Error Resume Next
            colnumbTER = Columns(inputColGaji).Column
            On Error GoTo 0
            
            If colnumbTER > 1 Then
                inputColGaji = Split(Cells(1, colnumbTER).Address(False, False), "1")(0)
                Exit Do
            Else
                MsgBox "Input Kolom yang diberikan harus berada tepat di sebelah sisi kanan kolom PTKP.", vbExclamation
            End If
        Else
            MsgBox "Kolom yang anda masukkan berupa huruf(contoh : 'C' [tanpa tanda petik])", vbExclamation
        End If
    Loop
    
End Function

Public Sub sumPPH21()
    Dim sumcells As Range
    Dim lastResult As String
    
    Set sumcells = Range("A1").End(xlToRight).End(xlDown).Offset(1, 0)
    sumcells.Select
    
    With sumcells
        ' Menjumlahkan seluruh nilai pada kolom TER
        .Formula = "=SUM(" & Range(sumcells.Offset(-1, 0).End(xlUp).Offset(1, 0), sumcells.Offset(-1, 0)).Address & ")"
        ' Formatting
        .Offset(-1, 0).Copy
        .PasteSpecial Paste:=xlPasteFormats
        .Offset(0, -5).Value = "Total"
    End With
    
    Application.CutCopyMode = False
    
    ' Menyiapkan hasil sesuai format
    lastResult = Format(sumcells, "#,##0")
    
    ' autofit kolom total
    Columns(sumcells.Column).AutoFit
    sumcells.Select
    
    ' Menampilkan perolehan PPh 21 TER
    MsgBox "Total PPh 21 TER yang harus dibayar adalah Rp " & lastResult, vbInformation
End Sub
