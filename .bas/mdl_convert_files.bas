Attribute VB_Name = "mdl_convert_files"

Sub SaveSLDDRWAsDWGAndPDF()

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim boolstatus As Boolean
Dim errors As Long
Dim warnings As Long
Dim swExportPDFData As SldWorks.ExportPdfData
Dim opt As Long

    ' ��ȡ�ļ���·��
    folderPath = Sheets("++.sldprt+").OLEObjects("TextBox1").Object.Value

    ' �����ļ�ϵͳ����
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' ��ʼ��SOLIDWORKSӦ�ó���
    Set swApp = CreateObject("SldWorks.Application")

    ' �����ļ����е�����.slddrw�ļ�
    For Each file In folder.Files
'        If Right(file.Name, 7) = ".slddrw" Or Right(file.Name, 7) = ".SLDDRW" Then
        If UCase(Right(file.Name, 7)) = ".SLDDRW" Then
            ' ��.slddrw�ļ�
'            Set swModel = swApp.OpenDoc6(strFilePath, 1, 0, 0, 0, 0, 0)
            Set swModel = swApp.OpenDoc6(file, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
            Set swDrawing = swModel

            ' ��ȡ�ļ�·��
            strModelPath = file.Path
            
            ' ����ΪDWG�ļ�
            strFilePath = Replace(UCase(strModelPath), "SLDDRW", "DWG", , , vbTextCompare)
            swModel.Extension.SaveAs strFilePath, 0, 0, Nothing, Empty, Empty

            ' ����ΪPDF�ļ�
            strFilePath = Replace(UCase(strModelPath), "SLDDRW", "PDF", , , vbTextCompare)
            swModel.Extension.SaveAs strFilePath, 0, 0, Nothing, Empty, Empty

            ' �رյ�ǰ�򿪵�.slddrw�ļ�
            swApp.CloseDoc swModel.GetTitle
        End If
    Next file

    ' ����
    Set swModel = Nothing
    Set swDrawing = Nothing
    Set swApp = Nothing
    Set fso = Nothing
    Set folder = Nothing
    Set file = Nothing
    
    MsgBox "Convert finished"
    
End Sub

