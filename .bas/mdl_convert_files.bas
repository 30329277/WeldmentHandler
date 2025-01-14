Attribute VB_Name = "mdl_convert_files"

Sub SaveSLDDRWAsDWGAndPDF()

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim boolstatus As Boolean
Dim errors As Long
Dim warnings As Long
Dim swExportPDFData As SldWorks.ExportPdfData
Dim opt As Long

    ' 获取文件夹路径
    folderPath = Sheets("++.sldprt+").OLEObjects("TextBox1").Object.Value

    ' 创建文件系统对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' 初始化SOLIDWORKS应用程序
    Set swApp = CreateObject("SldWorks.Application")

    ' 遍历文件夹中的所有.slddrw文件
    For Each file In folder.Files
'        If Right(file.Name, 7) = ".slddrw" Or Right(file.Name, 7) = ".SLDDRW" Then
        If UCase(Right(file.Name, 7)) = ".SLDDRW" Then
            ' 打开.slddrw文件
'            Set swModel = swApp.OpenDoc6(strFilePath, 1, 0, 0, 0, 0, 0)
            Set swModel = swApp.OpenDoc6(file, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
            Set swDrawing = swModel

            ' 获取文件路径
            strModelPath = file.Path
            
            ' 保存为DWG文件
            strFilePath = Replace(UCase(strModelPath), "SLDDRW", "DWG", , , vbTextCompare)
            swModel.Extension.SaveAs strFilePath, 0, 0, Nothing, Empty, Empty

            ' 保存为PDF文件
            strFilePath = Replace(UCase(strModelPath), "SLDDRW", "PDF", , , vbTextCompare)
            swModel.Extension.SaveAs strFilePath, 0, 0, Nothing, Empty, Empty

            ' 关闭当前打开的.slddrw文件
            swApp.CloseDoc swModel.GetTitle
        End If
    Next file

    ' 清理
    Set swModel = Nothing
    Set swDrawing = Nothing
    Set swApp = Nothing
    Set fso = Nothing
    Set folder = Nothing
    Set file = Nothing
    
    MsgBox "Convert finished"
    
End Sub

