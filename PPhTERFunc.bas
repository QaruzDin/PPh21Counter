Attribute VB_Name = "PPhTERFunc"
Option Explicit

Public Function PPh21TER(ptkp As String, gaji As Currency, ter As String, tarif As Single) As Currency
    
    
    PPh21TER = WorksheetFunction.RoundDown(gaji * tarif, 0)
    
End Function

Public Sub calcPPh21TER()
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


Function LoadCSVToArray(filePath As String) As Variant
    Dim fileContent As String
    Dim fileNumber As Integer
    Dim lines() As String
    Dim dataArray() As Variant
    Dim i As Long, numColumns As Long
    Dim tempArray() As String
    
    ' Mendapatkan nomor file yang tidak digunakan
    fileNumber = FreeFile
    
    ' Membuka file CSV
    On Error GoTo ErrorHandler
    Open filePath For Input As #fileNumber
    
    ' Membaca seluruh konten file
    fileContent = Input$(LOF(fileNumber), #fileNumber)
    Close #fileNumber
    
    ' Memisahkan konten berdasarkan baris
    lines = Split(fileContent, vbCrLf)
    
    ' Menghitung jumlah kolom dari baris pertama
    If UBound(lines) >= 0 Then
        tempArray = Split(lines(0), ";")
        numColumns = UBound(tempArray) + 1
    Else
        numColumns = 0
    End If
    
    ' Mengonversi data ke array dua dimensi
    ReDim dataArray(UBound(lines))
    
    For i = LBound(lines) To UBound(lines)
        ' Memisahkan baris berdasarkan delimiter titik koma
        tempArray = Split(lines(i), ";")
        
        ' Memastikan jumlah kolom sesuai
        If UBound(tempArray) < numColumns - 1 Then
            ReDim Preserve tempArray(numColumns - 1)
        End If
        
        dataArray(i) = tempArray
    Next i
    
    ' Mengembalikan array
    LoadCSVToArray = dataArray
    Exit Function

ErrorHandler:
    ' Menangani kesalahan jika file tidak ditemukan atau tidak dapat dibaca
    MsgBox "Error reading the file: " & Err.Description
    LoadCSVToArray = Array()
End Function

Public Sub TestLoadCSV()
    Dim csvArray As Variant
    Dim filePath As String
    Dim i As Long, j As Long
    
    ' Membuka file CSV
    filePath = Application.GetOpenFilename(FileFilter:="CSV Files (*.csv), *.csv", Title:="Mencari lokasi file TER.csv")
        
    ' Ganti "C:\path\to\file.csv" dengan jalur ke file CSV Anda
    If filePath <> "False" Then
        
        csvArray = LoadCSVToArray(filePath)
        
        ' Menampilkan data dalam Immediate Window
        For i = LBound(csvArray) To UBound(csvArray)
            For j = LBound(csvArray(i)) To UBound(csvArray(i))
                Debug.Print csvArray(i)(j); " ";
            Next j
            Debug.Print
        Next i
    Else
        MsgBox "Tidak ada file yang dipilih", vbInformation
    End If
End Sub

Public Function cariTarifTER()
    Dim csvArray As Variant
    Dim filePath As String
    Dim i As Long, j As Long
    
    ' Membuka file CSV
    filePath = Application.GetOpenFilename(FileFilter:="CSV Files (*.csv), *.csv", Title:="Mencari lokasi file TER.csv")
        
    ' Ganti "C:\path\to\file.csv" dengan jalur ke file CSV Anda
    If filePath <> "False" Then
        
        csvArray = LoadCSVToArray(filePath)
        
        
    Else
        MsgBox "Tidak ada file yang dipilih", vbInformation
    End If
        
End Function
