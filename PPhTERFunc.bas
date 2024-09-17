Attribute VB_Name = "PPhTERFunc"
Option Explicit

Public Sub formatting_PPh21TER()
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range("D1")
    
    With rng
        .Offset(0, 1) = "TER"
        .Offset(0, 2) = "Tarif"
        .Offset(0, 3) = "PPh 21"
        .Offset(0, 1).Resize(1, 3).HorizontalAlignment = xlCenter
    End With
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
        wsTg.Name = wsSc.Name
    Next wsSc
    
    wbSc.Close False
    
    MsgBox "Sheet berhasil disalin ke workbook ini :)", vbInformation
End Sub


Public Function tarifTER(TER As String, gajiBruto As Double) As Double
    Dim kolomTER As Range
    Dim batasBawah As Range
    Dim lo As ListObject
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DATA TER") ' pastikan nama sheet telah sesuai
    
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

Public Function PPH21(trf As Double, gajiBrt As Double) As Double
    ' Validasi input untuk menghindari kesalahan
    If trf < 0 Or trf > 1 Then
        MsgBox "Tarif harus antara 0,00 s/d 1,00.", vbExclamation
        Exit Function
    End If
    
    If gajiBrt < 0 Then
        MsgBox "Gaji bruto tidak boleh negatif.", vbExclamation
        Exit Function
    End If

    ' Perhitungan PPh 21
    PPH21 = Application.WorksheetFunction.RoundDown(trf * gajiBrt, 0)
End Function




