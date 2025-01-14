'最核心的子程序,获取尺寸对象的属性
Sub ExportViewDimensions(view As SldWorks.view, draw As SldWorks.DrawingDoc, fileNmb As Integer, viewNum As Integer)
    
    Dim swDispDim As SldWorks.DisplayDimension
    Dim swDim As SldWorks.Dimension
    Set swDispDim = view.GetFirstDisplayDimension5 'Gets the first display dimension in this drawing view.
    
    
    Dim swSheet As SldWorks.Sheet
    Set swSheet = view.Sheet 'Get the Sheet on which this View exists
    
    If swSheet Is Nothing Then
        Set swSheet = draw.Sheet(view.Name)
    End If
    
    '循环显示的尺寸,如果是空,返回ExportDrawingDimensions函数, 运行下一个尺寸
    Dim dimNum As Integer
    dimNum = 1
    While Not swDispDim Is Nothing
        
        Dim swAnn As SldWorks.Annotation
        Set swAnn = swDispDim.GetAnnotation
        
        Dim vPos As Variant
        vPos = swAnn.GetPosition()
        '这段代码的位置,移动到了重写和导出之间,为了保持重复运行程序的一致性, _
        有时候带 [尺1-1], 但打开文件首次运行程序不显示
                
        Set swDim = swDispDim.GetDimension2(0) '(*^_^*) 用.getdimension2()获得的对象就是具体的那个尺寸
        Dim swTextFormat As SldWorks.TextFormat
                
        Dim drwZone As String
        drwZone = swSheet.GetDrawingZone(vPos(0), vPos(1))
        vPos = GetPositionInDrawingUnits(vPos, draw)
        
        Dim tolType As String
        Dim minVal As Double
        Dim maxVal As Double
        'GetDimensionTolerance是获取尺寸的公差,
        GetDimensionTolerance draw, swDim, tolType, minVal, maxVal
        'OutputDimensionData 是拼接+写入, 把尺寸名称,类型,位置,值,公差等拼接后写入csv文件 , 注意这里做了修改 _
        因为drwZone也就是grid位置不准,所以用尺寸全名替代
        
        '这里是为了跳过基础尺寸 basic
        If tolType <> "Basic" Then
            swDispDim.SetText 2, "  [" & "DIM-" & viewNum & "-" & dimNum & "]"   '参数选 1234  对应着"前后上下", 不选带definition的
            swAnn.color = 255  '20221015 增加颜色说明
            '20230413把字体改小一些
            For ii = 0 To swAnn.GetTextFormatCount - 1
                Set swTextFormat = swAnn.GetTextFormat(ii)
                'Change text to be 2mm, bold, italic, and Comic Sans MS font
                swTextFormat.CharHeight = 0.004
                swTextFormat.Bold = True
                swTextFormat.Italic = False
                swTextFormat.WidthFactor = 0.4
                bRet = swAnn.SetTextFormat(ii, False, swTextFormat)
            Next
            
'            draw.EditRebuild 'rebuild一下
            '{{{抓取尺寸旁边的文本,因为这个swDim的成员主要是数值,类型等(不包括它旁边的文本)
            Dim swDisplayData As SldWorks.DisplayData
            Set swDisplayData = swAnn.GetDisplayData
            Dim strFullDimension As String
            Dim i As Long
            strFullDimension = "" '拼接的文本循环前重置一下
            For i = swDisplayData.GetTextCount To 0 Step -1 '采用倒序
            strFullDimension = swDisplayData.GetTextAtIndex(i) & " " & strFullDimension
            Next i
            '}}}抓取尺寸旁边的文本
            '如果是倒角尺寸,则需要调用GetSystemChamferValues, 其他类型尺寸调用swDim.GetValue3
            Dim length As Double
            Dim angle As Double
            If GetDimensionType(swDispDim) = "Chamfer" Then
                swDim.GetSystemChamferValues length, angle
                WriteToSheet "  [" & "DIM-" & viewNum & "-" & dimNum & "]", GetDimensionType(swDispDim), _
                 ConvertToUserUnits(draw, CDbl(length), swLengthUnit), _
                strFullDimension, tolType, minVal, maxVal, swDim.FullName, view.Name, CDbl(vPos(0)), CDbl(vPos(1))
                dimNum = dimNum + 1
            Else
                WriteToSheet "  [" & "DIM-" & viewNum & "-" & dimNum & "]", GetDimensionType(swDispDim), _
                 CDbl(swDim.GetValue3(swInConfigurationOpts_e.swThisConfiguration, Empty)(0)), _
                strFullDimension, tolType, minVal, maxVal, swDim.FullName, view.Name, CDbl(vPos(0)), CDbl(vPos(1))
                dimNum = dimNum + 1
            End If
        End If
        '这里是为了跳过基础尺寸 basic
        
        Set swDispDim = swDispDim.GetNext5

    Wend
    
    '{20221203
    Set swDispDim = Nothing
    Set swSheet = Nothing
    Set swSheet = Nothing
    Set swAnn = Nothing
    Set swDim = Nothing
    Set swDisplayData = Nothing
    Set swDispDim = Nothing
    '20221203}
    
End Sub '最核心的子程序,获取尺寸对象的属性