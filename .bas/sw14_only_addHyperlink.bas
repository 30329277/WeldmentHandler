Attribute VB_Name = "sw14_only_addHyperlink"
' 改一下，保留全路径和文件名，超链接也取消掉

Sub only_path_left(ByVal sht As Worksheet)
On Error Resume Next
Dim i, j, k, o As Integer

For Each rng In sht.Range(sht.Range("a2"), sht.Range("a1048576").End(xlUp))
    
    i = InStrRev(rng, "\", -1, 0)
    j = InStrRev(rng, ".", -1, 0)
    rng.Offset(0, 1) = Left(Right(rng, Len(rng) - i), j - i - 1)
    k = InStrRev(rng.Offset(0, 1), "_", -1, 0)
    o = InStrRev(Left(rng.Offset(0, 1), k - 1), "_", -1, 0)
    rng.Offset(0, 7) = Left(rng.Offset(0, 1), o - 1)
    rng.Offset(0, 8) = Left(rng.Offset(0, 1), o - 1)
    rng.Offset(0, 9) = Left(Right(rng.Offset(0, 1), Len(rng.Offset(0, 1)) - o), k - o - 1)
    rng.Offset(0, 10) = Right(rng.Offset(0, 1), Len(rng.Offset(0, 1)) - k)
    
    
    '把截取路径和文件夹注释掉，改成截取文件名不含扩展名
    'i = InStrRev(sht.Cells(Rng.Row, 1), "\", -1, 0)
    'j = InStrRev(Left(Rng, i - 1), "\", -1, 0)
    'sht.Cells(Rng.Row, 2) = Left(Left(Rng, i - 1), j)
    'sht.Cells(Rng.Row, 3) = Right(Left(Rng, i - 1), Len(Left(Rng, i - 1)) - j) & "\" & Right(Rng, Len(Rng) - i)
    '把截取路径和文件夹注释掉，改成截取文件名不含扩展名
    '把超链接注释掉
    'sht.Hyperlinks.Add sht.Cells(Rng.Row, 1), sht.Cells(Rng.Row, 1)
    'sht.Hyperlinks.Add sht.Cells(Rng.Row, 2), sht.Cells(Rng.Row, 2)
    'sht.Hyperlinks.Add sht.Cells(Rng.Row, 3), Left(Rng, i)
    '把超链接注释掉
    
Next rng

'With sht.Range("A1:C1")
'.ColumnWidth = 50
'.RowHeight = 20
'.AutoFilter
'.HorizontalAlignment = xlCenter
'.Interior.Color = 65535
'End With

'sht.Range("A1") = "full name"
'sht.Range("B1") = "file name"
'sht.Range("C1") = "lower folder"

End Sub


'Sub only_path_left(ByVal sht As Worksheet)
'
'Dim i As Integer
'Dim j As Integer
'
'For Each Rng In sht.Range(sht.Range("a2"), sht.Range("a1048576").End(xlUp))
'i = InStrRev(sht.Cells(Rng.Row, 1), "\", -1, 0)
'j = InStrRev(Left(Rng, i - 1), "\", -1, 0)
'sht.Cells(Rng.Row, 2) = Left(Left(Rng, i - 1), j)
'sht.Cells(Rng.Row, 3) = Right(Left(Rng, i - 1), Len(Left(Rng, i - 1)) - j) & "\" & Right(Rng, Len(Rng) - i)
'
'sht.Hyperlinks.Add sht.Cells(Rng.Row, 1), sht.Cells(Rng.Row, 1)
'sht.Hyperlinks.Add sht.Cells(Rng.Row, 2), sht.Cells(Rng.Row, 2)
'sht.Hyperlinks.Add sht.Cells(Rng.Row, 3), Left(Rng, i)
'
'Next Rng
'
'With sht.Range("A1:C1")
'.ColumnWidth = 50
'.RowHeight = 20
'.AutoFilter
'.HorizontalAlignment = xlCenter
'.Interior.Color = 65535
'End With
'
'sht.Range("A1") = "full name"
'sht.Range("B1") = "upper folder"
'sht.Range("C1") = "lower folder"
'
'End Sub





