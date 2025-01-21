Attribute VB_Name = "sw14_func_create_A2"
    
'����ͼģ������·��
Const strFormatePath As String = _
"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2019\lang\english\sheetformat\A2.drwdot"
'BOM��ģ��
Const strBOMTemplatePath As String = _
"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\bom-standard.sldbomtbt"


Public Function Create_Drawings(strModelPath As String)
    On Error Resume Next
'    Dim swApp As SldWorks.SldWorks
    Set swApp = CreateObject("SldWorks.Application")
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    
    'Create new drawing
    '{
    Dim strTemplatePath As String
    strTemplatePath = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)
    Set swModel = swApp.NewDocument(strFormatePath, swDwgPapersUserDefined, 0.297 * 2, 0.21 * 2)
    '}
    
    Set swDraw = swModel
    
    'Setup first sheet
    swDraw.SetupSheet5 "Sheet1", 12, 12, 1, 5, True, strTemplatePath, 0.297 * 2, 0.21 * 2, "Default", True
    
    'Add a new sheet
    swDraw.NewSheet3 "Sheet1", 12, 12, 1, 5, True, strFormatePath, 0.297 * 2, 0.21 * 2, "Default"
    
    'Change name of sheets
    Dim swSheet As SldWorks.Sheet
    Dim vSheetNames As Variant
    Dim i As Integer
    vSheetNames = swDraw.GetSheetNames
    For i = 0 To UBound(vSheetNames)
        Set swSheet = swDraw.Sheet(vSheetNames(i))
        If i = 0 Then swSheet.SetName "Sheet1"
        If i = 1 Then swSheet.SetName "Sheet2"
    Next i

    'Activate first sheet
    vSheetNames = swDraw.GetSheetNames
    swDraw.ActivateSheet vSheetNames(0)
    
    'Add iso view of assy on first sheet######### ע�͵�
    Dim swView As SldWorks.View
    Dim swIsoView As SldWorks.View
    Set swView = swDraw.CreateDrawViewFromModelView3(strModelPath, "*Isometric", 0.5, 0.3, 0)
    Application.Wait Now() + VBA.TimeValue("00:00:02")
    
    'Add standard three-view
    swDraw.Create3rdAngleViews2 strModelPath
    
    'Traverse views and turn off tangent edges
    Dim vViews As Variant
    vViews = swSheet.GetViews
    For i = 0 To UBound(vViews)
        Set swView = vViews(i)
        swView.SetDisplayTangentEdges2 swTangentEdgesHidden
    Next i
    
    'Rebuild
    swModel.ForceRebuild3 True
    
    'Add dimensions
    swDraw.InsertModelAnnotations3 swImportModelItemsFromEntireModel, _
        swInsertDimensionsMarkedForDrawing, True, True, False, True
    
    'Select all display dimensions
    Dim swDispDim As SldWorks.DisplayDimension
    Dim swDispData As SldWorks.DisplayData
    swModel.ClearSelection2 True
    For i = 0 To UBound(vViews)
        Set swView = vViews(i)
        Set swDispDim = swView.GetFirstDisplayDimension5
        While Not swDispDim Is Nothing
            swModel.Extension.SelectByID2 swDispDim.GetNameForSelection, _
                "DIMENSION", 0, 0, 0, True, 0, Nothing, 0
            Set swDispDim = swDispDim.GetNext5
        Wend
    Next i
    
    'Reposition dimensions
    swModel.Extension.AlignDimensions swAlignDimensionType_AutoArrange, 0.06
    
    'Activate first sheet
    swDraw.ActivateSheet vSheetNames(0)
    
    'Autoballoon######### ע�͵�
    Dim swFeat As SldWorks.Feature
    Dim vNotes As Variant
    
    Dim swNote As SldWorks.Note
    Dim swAnn As SldWorks.Annotation
    Dim vPos As Variant

    'Determine which view is top view
    Dim swMathUtil As SldWorks.MathUtility
    Dim swTransform As SldWorks.MathTransform
    Dim swMathVect As SldWorks.MathVector
    Dim dblPts(2) As Double
    Dim vPts As Variant
    Set swMathUtil = swApp.GetMathUtility
    
    dblPts(0) = 0
    dblPts(1) = 1
    dblPts(2) = 0
    Set swMathVect = swMathUtil.CreateVector(dblPts)
    
    For i = 1 To UBound(vViews)
        Set swView = vViews(i)
        Set swTransform = swView.ModelToViewTransform
'        Set swTransform = swTransform.Inverse
        Set swMathVect = swMathVect.MultiplyTransform(swTransform)
        Set swMathVect = swMathVect.Normalise
        vPts = swMathVect.ArrayData
        If Round(vPts(0), 1) = 0 And Round(vPts(1), 1) = 0 And _
            Round(vPts(2), 1) = 1 Then Exit For
    Next i
    
    'Get center of top drawing view
    vPts = swView.GetXform
    
    'Get sheet scale
    Dim vSheetsProps As Variant
    vSheetsProps = swSheet.GetProperties
    
    'ͣ��2s
    Application.Wait Now() + VBA.TimeValue("00:00:02")
    
    'Create line+/-100mm in the X from view center
    Dim swSketchMgr As SldWorks.SketchManager
    Dim swSketchSeg As SldWorks.SketchSegment
    Set swSketchMgr = swModel.SketchManager
    Dim dblX1 As Double, dblY1 As Double, dblX2 As Double, dblY2 As Double
    dblX1 = Round(vSheetsProps(3) * (vPts(0) - 0.1), 3)
    dblY1 = Round(vSheetsProps(3) * vPts(1), 3)
    dblX2 = Round(vSheetsProps(3) * (vPts(0) + 0.1), 3)
    dblY2 = Round(vSheetsProps(3) * vPts(1), 3)
    Set swSketchSeg = swSketchMgr.CreateLine(dblX1, dblY1, 0, dblX2, dblY2, 0)
    Debug.Print vSheetsProps(3) & Chr(32) & vPts(0) & Chr(32); vPts(1) & Chr(10)
    Debug.Print dblX1 & Chr(32) & dblY1 & Chr(32) & dblX2 & Chr(32) & dblY2 & Chr(10)
    'Select the line
    swSketchSeg.Select4 False, Nothing
    
    'Create section View
    Set swView = swDraw.CreateSectionViewAt5(0, 0, 0, "A", 0, Emprty, 0)
    
    'Select and flip section view, then rebuild
    swModel.Extension.SelectByID2 "Selction Line1", "SECTIONLINE", 0, 0, 0, False, 0, Nothing, 0
    swDraw.FlipSectionLine
    swModel.ForceRebuild3 False
    
    'Break alignment
    swView.AlignWithView 0, Nothing
    
    'Relocate view
    vPts = swView.GetXform
    vPts(0) = 0.34
    vPts(1) = 0.3
    swView.SetXform vPts
    
    '########DIMENSION THIN PIECE ########
    'Get and ulderlying model edge as the launching point for yor face locator code
    Dim vEdges As Variant
    Dim swEdge As SldWorks.Edge
    vEdges = swView.GetPolylines7(1, Empty)
    For i = 0 To UBound(vEdges)
        Set swEdge = vEdges(i)
        If TypeOf swEdge Is SldWorks.Edge Then Exit For
    Next i

    'Locate the faces for dimensioning
    Dim swBody As SldWorks.Body2
    Dim swFace As SldWorks.Face2
    Dim swSurf As SldWorks.Surface
    Dim vBodies As Variant
    Dim vFaces As Variant
    Dim swFinalFace(1) As SldWorks.Face2
    Dim swEnt As SldWorks.Entity

    Set swBody = swEdge.GetBody
    vFaces = swBody.GetFaces
    For i = 0 To UBound(vFaces)
        Set swFace = vFaces(i)
        Set swSurf = swFace.GetSurface
        If swSurf.IsPlane Then
            If swFinalFace(0) Is Nothing Then
                Set swFinalFace(0) = swFace
                Set swFinalFace(1) = swFace
            End If
            If swFace.GetArea > swFinalFace(0).GetArea Then
                Set swFinalFace(1) = swFinalFace(0)
                Set swFinalFace(0) = swFace
            ElseIf swFace.GetArea > swFinalFace(1).GetArea Then
                Set swFinalFace(1) = swFace
            End If
        End If
    Next i

    'Get and edge from the face and select in drawing view
    For i = 0 To 1
        vEdges = swFinalFace(i).GetEdges
        Set swEdge = vEdges(0)
        Set swEnt = swEdge
        If i = 0 Then swView.SelectEntity swEnt, False
        If i = 1 Then swView.SelectEntity swEnt, True
    Next i

    'Garbage collection
    Set swFinalFace(0) = Nothing
    Set swFinalFace(1) = Nothing

    'Add dimension ����һ��clear selection
    swModel.AddDimension2 0.2, 0.165, 0
    swModel.ClearSelection2 True
    
    'Save drawing to the same location as model
    Dim strFilePath As String
'    strFilePath = swIsoView.GetReferencedModelName
    strFilePath = Replace(strModelPath, "SLDPRT", "SLDDRW", , , vbTextCompare)
    swModel.Extension.SaveAs strFilePath, 0, 0, Nothing, Empty, Empty
    strFilePath = Replace(strModelPath, "SLDPRT", "PDF", , , vbTextCompare)
    swModel.Extension.SaveAs strFilePath, 0, 0, Nothing, Empty, Empty
    
    swApp.CloseDoc swModel.GetTitle
  
End Function








