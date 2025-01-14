Attribute VB_Name = "sw14_0main_start_create"
'***************程序说明
'1.读取文件夹内的所有sldprt文件名信息,分割字符
'2.利用solidworks接口批量生成工程图
'***************程序输入
'1.打开模块sw14_0main_start_create，把Set retVal = DirFolder( _“填入新的文件夹路径”），
    '这个路径一般是 save bodies (可以是rm plate cutting ， 也可以是 rm tube and shapes) 。
'2.其余模块一般不需要更改,运行程序后,按照提示操作即可,第一次提示是否更改材质、粗糙度,选择是会完成sheet内容后终止程序,选择否会继续执行
'3.工程图生成后,会第二次提示是否更改,这里100%需要更改,此时忽视按钮,直接在solidworks界面更改,完成后点按钮生成pdf文件。
'***************程序输出
'1.sheet中有一个类似BOM
'2.批量得到工程和slddrw和pdf文件


Sub MainCreateEngineeringDrawings()

' retVal is a worksheet

'Set retVal = DirFolder _
'("C:\Machines\SFSRFH205015 90000079813 FILLER FMA COMBI PREDIS 205015\04342570605 INTERM.BASE MACHINING C.1080 FDA\shim\", _
'"**.sldprt*")

'20240220 这里改成读取文本框
'Set retVal = DirFolder(Sheets("++.sldprt+").OLEObjects("TextBox1").Object.Value, "**.sldprt*")  ' retVal is a worksheet

'20240926 增加了反斜杠检查
Dim folderPath As String

' 获取工作表 "++.sldprt+" 中的 TextBox1 的值
folderPath = Sheets("++.sldprt+").OLEObjects("TextBox1").Object.Value

' 检查路径是否以反斜杠 "\" 结尾，如果没有则添加
If Right(folderPath, 1) <> "\" Then
    folderPath = folderPath & "\"
End If

' 使用修改后的路径调用 DirFolder
Set retVal = DirFolder(folderPath, "**.sldprt*")



'Add table header 法语描述与模板对应
retVal.Range("a1:k1") = Array("Full Name", "File Name", _
"Matière", "Mat_fiche", "Etat_surface", "Dessiné_par", "Date_de_création", "Désignation_ang", "Description", "SAPmaterial", "SAPVersion")

'用于获得模型的文件名,无扩展名,截取数字部分,截取描述部分
Call sw14_only_addHyperlink.only_path_left(retVal)
'停顿2s
Application.Wait Now() + VBA.TimeValue("00:00:02")


'20240129 把下划线去掉
For Each rng In retVal.Range("h2:i500")
    rng.Value = Replace(rng, "_", " ")
Next rng


'####msgbox
If MsgBox("是否还要修改材质和粗糙度?", vbYesNo) = vbYes Then End
'####msgbox

'20240219 把下划线去掉
For Each rng In retVal.Range("h2:i500")
    rng.Value = Replace(rng, "_", " ")
Next rng

'打开模型,写入属性,并保存模型
For Each rng In retVal.Range("a2:a500")
    If rng <> "" And rng.Interior.Color <> 16777215 Then '20230213增加了颜色判断
        Call mdl_update_standard_views.AdjustPartToMaxAreaPlane(rng.Value)
        Call sw14_custprops1.CustomProperty(rng.Value, rng.Offset(0, 2), rng.Offset(0, 3), rng.Offset(0, 4), rng.Offset(0, 7), rng.Offset(0, 8), rng.Offset(0, 9), rng.Offset(0, 10))
    End If
Next rng

'创建工程图,写入模型中的属性到标题栏,包括材质、粗糙度、作者、描述等等
'{'读取单选框内容开始
'Dim Process As String
'If Sheet1.OLEObjects("OptionButton1").Object.value = True Then
'    Process = "Machining"
'ElseIf Sheet1.OLEObjects("OptionButton2").Object.value = True Then
'    Process = "Fabrication"
'ElseIf Sheet1.OLEObjects("OptionButton3").Object.value = True Then
'    Process = "Cutting"
'End If
'读取单选框内容结束}
'20230210 新增加了 颜色判断
For Each rng In retVal.Range("a2:a500")
    If rng <> "" And rng.Interior.Color <> 16777215 And retVal.OLEObjects("OptionButton1").Object.Value = True Then
        Call sw14_func_create_A2.Create_Drawings(rng.Value)
    ElseIf rng <> "" And rng.Interior.Color <> 16777215 And retVal.OLEObjects("OptionButton2").Object.Value = True Then
        Call sw14_func_create_A1.Create_Drawings(rng.Value)
    ElseIf rng <> "" And rng.Interior.Color <> 16777215 And retVal.OLEObjects("OptionButton3").Object.Value = True Then
        Call sw14_func_create_A3.Create_Drawings(rng.Value)
    End If
Next rng
    
'20240129 把下划线去掉
For Each rng In retVal.Range("h2:i500")
    rng.Value = Replace(rng, "_", " ")
Next rng
    
'autofit the column width
    
retVal.Columns("B:K").EntireColumn.AutoFit
retVal.Columns("B:K").HorizontalAlignment = xlLeft
    
'clear contents
Set retVal = Nothing
MsgBox "finished"
    

End Sub



