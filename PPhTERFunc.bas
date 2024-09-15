Attribute VB_Name = "PPhTERFunc"
Option Explicit

Function cariTER(lookup_value As String) As Variant
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
        If lookup_table(i)(0) = lookup_value Then
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
