Attribute VB_Name = "sw14_0main_start_create"
'***************����˵��
'1.��ȡ�ļ����ڵ�����sldprt�ļ�����Ϣ,�ָ��ַ�
'2.����solidworks�ӿ��������ɹ���ͼ
'***************��������
'1.��ģ��sw14_0main_start_create����Set retVal = DirFolder( _�������µ��ļ���·��������
    '���·��һ���� save bodies (������rm plate cutting �� Ҳ������ rm tube and shapes) ��
'2.����ģ��һ�㲻��Ҫ����,���г����,������ʾ��������,��һ����ʾ�Ƿ���Ĳ��ʡ��ֲڶ�,ѡ���ǻ����sheet���ݺ���ֹ����,ѡ�������ִ��
'3.����ͼ���ɺ�,��ڶ�����ʾ�Ƿ����,����100%��Ҫ����,��ʱ���Ӱ�ť,ֱ����solidworks�������,��ɺ�㰴ť����pdf�ļ���
'***************�������
'1.sheet����һ������BOM
'2.�����õ����̺�slddrw��pdf�ļ�


Sub MainCreateEngineeringDrawings()

' retVal is a worksheet

'Set retVal = DirFolder _
'("C:\Machines\SFSRFH205015 90000079813 FILLER FMA COMBI PREDIS 205015\04342570605 INTERM.BASE MACHINING C.1080 FDA\shim\", _
'"**.sldprt*")

'20240220 ����ĳɶ�ȡ�ı���
'Set retVal = DirFolder(Sheets("++.sldprt+").OLEObjects("TextBox1").Object.Value, "**.sldprt*")  ' retVal is a worksheet

'20240926 �����˷�б�ܼ��
Dim folderPath As String

' ��ȡ������ "++.sldprt+" �е� TextBox1 ��ֵ
folderPath = Sheets("++.sldprt+").OLEObjects("TextBox1").Object.Value

' ���·���Ƿ��Է�б�� "\" ��β�����û�������
If Right(folderPath, 1) <> "\" Then
    folderPath = folderPath & "\"
End If

' ʹ���޸ĺ��·������ DirFolder
Set retVal = DirFolder(folderPath, "**.sldprt*")



'Add table header ����������ģ���Ӧ
retVal.Range("a1:k1") = Array("Full Name", "File Name", _
"Mati��re", "Mat_fiche", "Etat_surface", "Dessin��_par", "Date_de_cr��ation", "D��signation_ang", "Description", "SAPmaterial", "SAPVersion")

'���ڻ��ģ�͵��ļ���,����չ��,��ȡ���ֲ���,��ȡ��������
Call sw14_only_addHyperlink.only_path_left(retVal)
'ͣ��2s
Application.Wait Now() + VBA.TimeValue("00:00:02")


'20240129 ���»���ȥ��
For Each rng In retVal.Range("h2:i500")
    rng.Value = Replace(rng, "_", " ")
Next rng


'####msgbox
If MsgBox("�Ƿ�Ҫ�޸Ĳ��ʺʹֲڶ�?", vbYesNo) = vbYes Then End
'####msgbox

'20240219 ���»���ȥ��
For Each rng In retVal.Range("h2:i500")
    rng.Value = Replace(rng, "_", " ")
Next rng

'��ģ��,д������,������ģ��
For Each rng In retVal.Range("a2:a500")
    If rng <> "" And rng.Interior.Color <> 16777215 Then '20230213��������ɫ�ж�
        Call mdl_update_standard_views.AdjustPartToMaxAreaPlane(rng.Value)
        Call sw14_custprops1.CustomProperty(rng.Value, rng.Offset(0, 2), rng.Offset(0, 3), rng.Offset(0, 4), rng.Offset(0, 7), rng.Offset(0, 8), rng.Offset(0, 9), rng.Offset(0, 10))
    End If
Next rng

'��������ͼ,д��ģ���е����Ե�������,�������ʡ��ֲڶȡ����ߡ������ȵ�
'{'��ȡ��ѡ�����ݿ�ʼ
'Dim Process As String
'If Sheet1.OLEObjects("OptionButton1").Object.value = True Then
'    Process = "Machining"
'ElseIf Sheet1.OLEObjects("OptionButton2").Object.value = True Then
'    Process = "Fabrication"
'ElseIf Sheet1.OLEObjects("OptionButton3").Object.value = True Then
'    Process = "Cutting"
'End If
'��ȡ��ѡ�����ݽ���}
'20230210 �������� ��ɫ�ж�
For Each rng In retVal.Range("a2:a500")
    If rng <> "" And rng.Interior.Color <> 16777215 And retVal.OLEObjects("OptionButton1").Object.Value = True Then
        Call sw14_func_create_A2.Create_Drawings(rng.Value)
    ElseIf rng <> "" And rng.Interior.Color <> 16777215 And retVal.OLEObjects("OptionButton2").Object.Value = True Then
        Call sw14_func_create_A1.Create_Drawings(rng.Value)
    ElseIf rng <> "" And rng.Interior.Color <> 16777215 And retVal.OLEObjects("OptionButton3").Object.Value = True Then
        Call sw14_func_create_A3.Create_Drawings(rng.Value)
    End If
Next rng
    
'20240129 ���»���ȥ��
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



