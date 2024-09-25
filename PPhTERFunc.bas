Attribute VB_Name = "PPhTERFunc"
Option Explicit

Dim colGaji As String


Public Sub formatting_PPh21TER()
    Dim ws As Worksheet
    Dim rng As Range

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
    
    Import_DataTER
    
    iterratingCell
    
    sumPPH21
    
End Sub

Function cariTER(PTKP As String) As Variant
    Dim lookup_table As Variant
    Dim i As Integer

    ' Define the TER's categories
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

    ' Loop through the lookup table to find the matching value
    For i = LBound(lookup_table) To UBound(lookup_table)
        If lookup_table(i)(0) = PTKP Then
            ' Return the categories according the first column (PTKP)
             cariTER = lookup_table(i)(1)
             Exit Function
        End If
    Next i

    ' If no match is found, return "Invalid"
    cariTER = "Invalid"
End Function

Public Sub Import_DataTER()
    Dim fileDialog As fileDialog
    Dim selectedFile As String
    Dim wbSc As Workbook ' Workbook Souce
    Dim wsSc As Worksheet '  Worksheet Source
    Dim wsTg As Worksheet ' Worksheet Target
    
    ' Dialog file untuk memilih file DATA TER
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Pilih file DATA TER"
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub ' pembatalan oleh user
        selectedFile = .SelectedItems(1)
    End With
    
    ' Mencopy isi file
    Set wbSc = Workbooks.Open(selectedFile)
    
    For Each wsSc In wbSc.Sheets
        wsSc.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsTg = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
        ' penanganan error bila DATA TER telah ter
        On Error Resume Next
        wsTg.Name = wsSc.Name
        If Err.Number <> 0 Then
            MsgBox "Data TER terdeksi telah tersedia! Data yang akan diunggah akan dihapus secara otomatis.", vbExclamation
            Application.DisplayAlerts = False
            wsTg.Delete
            Application.DisplayAlerts = True
            Err.Clear
        End If
        On Error GoTo 0
    Next wsSc
    
    wbSc.Close False
    
    ThisWorkbook.Sheets(ActiveSheet.Index - 1).Activate
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
            Set lo = ws.ListObjects("tabelA") ' Ganti dengan nama tabel yang sesuai
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
            MsgBox "Invalid data TER", vbExclamation
            Exit Function
    End Select
            
    ' pencarian tarif TER
    tarifTER = Application.WorksheetFunction.Index(kolomTER, Application.WorksheetFunction.Match(gajiBruto, batasBawah, 1))
    
End Function


Function PPH21TER(trf As Double, gajiBrt As Double) As Double
  
  PPH21TER = WorksheetFunction.RoundDown(trf * gajiBrt, 0)
End Function


Public Sub iterratingCell()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim startTime As Single
    Dim timeout As Single
    
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range(colGaji & "1").Offset(0, 1)
    
    ' mendeteksi kolom aktif terakhir
    lastRow = ws.Cells(ws.Rows.Count, colGaji).End(xlUp).row - 1
    
    ' timeout setting for infinte loop prevention (timeout is in second)
    startTime = Timer
    timeout = 60
    
    'iterrating one cell at the time
    For i = 1 To lastRow
        For j = 1 To 3
            Set cell = rng.Offset(i, j - 1) ' will be moving one cell above
        
            cell.Select
            Select Case j
                Case 1
                    cell.Value = cariTER(cell.Offset(0, -2))
                    cell.HorizontalAlignment = xlCenter ' formatting : centered
                Case 2
                    cell.Value = tarifTER(cell.Offset(0, -1), cell.Offset(0, -2))
                    cell.NumberFormat = "0.00%" ' formatting : percentage
                Case 3
                    cell.Value = PPH21TER(cell.Offset(0, -1), cell.Offset(0, -3))
                    cell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)" ' formatting : IDR
            End Select
            
        Next j
        
            
        ' Here! You can add more conditions as needed in this block of FOR loop

        ' Check if the timeout duration has been exceeded
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
                    "(Pastikan Kolom PTKP berada disisi kiri kolom.", "Input Kolom Gaji")
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
        .Offset(0, -1).Value = "Total"
    End With
    
    ' Menyiapkan hasil sesuai format
    lastResult = Format(sumcells, "#,##0")
    
    ' autofit kolom total
    Columns(sumcells.Column).AutoFit
    
    ' Menampilkan perolehan PPh 21 TER
    MsgBox "Total PPh 21 TER yang harus dibayar adalah Rp " & lastResult, vbOKOnly
End Sub
