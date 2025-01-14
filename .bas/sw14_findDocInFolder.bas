Attribute VB_Name = "sw14_findDocInFolder"

Public Function DirFolder(lj As String, searchkey As String) As Worksheet


Dim MyName, Dic, Did, i, t, F, TT, MyFileName, ii

    On Error Resume Next ' error happens
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = Nothing
'    Set objFolder = objShell.BrowseForFolder(0, "选择文件夹", 0, 0)

    Set objShell = Nothing
  
    t = Time
    Set Dic = CreateObject("Scripting.Dictionary")    '创建一个字典对象
    Set Did = CreateObject("Scripting.Dictionary")
    Dic.Add (lj), ""
    i = 0
    Do While i < Dic.Count
        Ke = Dic.keys   '开始遍历字典
        MyName = Dir(Ke(i), vbDirectory)    '查找目录
        Do While MyName <> ""
            If MyName <> "." And MyName <> ".." Then
                If (GetAttr(Ke(i) & MyName) And vbDirectory) = vbDirectory Then    '如果是次级目录
                    Dic.Add (Ke(i) & MyName & "\"), ""  '就往字典中添加这个次级目录名作为一个条目
                End If
            End If
            MyName = Dir    '继续遍历寻找
        Loop
        i = i + 1
    Loop
    Did.Add ("文件清单"), ""    '以查找D盘下所有EXCEL文件为例
    For Each Ke In Dic.keys
        MyFileName = Dir(Ke & searchkey)
        Do While MyFileName <> ""
            Did.Add (Ke & MyFileName), ""
            MyFileName = Dir
        Loop
    Next
    Dim shtName1 As String
    Dim shtName2 As String
    shtName1 = Replace(searchkey, Chr(34), "")
    shtName2 = Replace(shtName1, Chr(42), "+")
    
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name = shtName2 Then
'            Sheets(shtName2).Cells.Delete
            sh.Range("a1").End(xlDown).Offset(1, 0).Resize(500, 13).ClearContents
            sh.Columns("a:b").ClearContents '这里修改一下,修改的材质和粗糙度可以保留
            sh.Columns("h:m").ClearContents
            
            F = True
            Exit For
        Else
            F = False
        End If
    Next
    If Not F Then
        Sheets.Add.Name = shtName2
    End If
    Sheets(shtName2).[A1].Resize(Did.Count, 1) = WorksheetFunction.Transpose(Did.keys)
    TT = Time - t
'    MsgBox Minute(TT) & "分" & Second(TT) & "秒"

'**********增加超链接 add hyperlink

ii = 2
For Each rng In Sheets(shtName2).Range(Sheets(shtName2).Cells(1, "A"), Sheets(shtName2).Cells(1048576, "A").End(xlUp))
            Sheets(shtName2).Hyperlinks.Add Sheets(shtName2).Cells(ii, "M"), Sheets(shtName2).Cells(rng.Row + 1, 1)
            ii = ii + 1
Next
'**********增加超链接 add hyperlink


Set DirFolder = Sheets(shtName2) ' function retval



End Function


