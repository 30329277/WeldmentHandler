Attribute VB_Name = "mdl_update_standard_views"
Dim lErrors As Long
Dim lWarnings As Long
'>>>>>>>>>>>>
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swDocSpec As SldWorks.DocumentSpecification
Dim swPart As SldWorks.PartDoc
Dim swBody As SldWorks.Body2
Dim swFace As SldWorks.Face2
Dim swFinalFace As SldWorks.Face2
Dim swEnt As SldWorks.Entity
Dim swSurf As SldWorks.Surface
Dim swModelDocExt As SldWorks.ModelDocExtension     ' add for measure perimeter
Dim swMeasure As SldWorks.Measure    ' add for measure perimeter
Dim status As Boolean       ' add for measure perimeter
Dim vBodies As Variant
Dim i As Integer
Dim dblArea As Double
Dim dblPeri As Double
Dim val As Double ' add for measure volume
' new added
Dim errors As Long
Dim warnings As Long
Dim boolstatus As Boolean
Dim swFeat As SldWorks.Feature
Dim swSubFeat As SldWorks.Feature
' new added

Public Function AdjustPartToMaxAreaPlane(strFilePath As String) As String

    'todo如果文件名字中有非法字符比如"Φ", 如何处理?
'    Set swApp = Application.SldWorks
    Set swApp = CreateObject("SldWorks.Application")
'    Set swModel = swApp.ActiveDoc
    Set swDocSpec = swApp.GetOpenDocSpec(strFilePath)
    Set swModel = swApp.OpenDoc7(swDocSpec)
    Set swPart = swModel
    
    '##############Get flat-pattern feature start
    'traversal start
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName = "FlatPattern" Then Exit Do
        Set swFeat = swFeat.GetNextFeature
    Loop
    On Error Resume Next ' 如果不是折弯件会报错
    swFeat.Select2 False, 0
    'traversal end
    swModel.EditUnsuppress2
    '##############Get flat-pattern feature end
    
    vBodies = swPart.GetBodies2(swSolidBody, False)
    Set swBody = vBodies(0)
    
    dblArea = 0
    
    vFaces = swBody.GetFaces
    For i = 0 To UBound(vFaces)
        Set swFace = vFaces(i)
        Set swSurf = swFace.GetSurface
'        If swSurf.IsCylinder Then
            If swFace.GetArea > dblArea Then
                dblArea = swFace.GetArea
                Set swFinalFace = swFace
            End If
'        End If
    Next i

    Debug.Print "LargestArea: " & dblArea
    
    Set swEnt = swFinalFace
    swEnt.Select4 False, Nothing
    
'..............................................................
    
    ' 更新标准视图 20240418
    
    swModel.ShowNamedView2 "*Normal To", -1
    
    ModelViewManager = swModel.ModelViewManager
    ModelViewManager.ResetStandardViews
        Set swModelDocExt = swModel.Extension
        swModelDocExt.UpdateStandardViews "", 5

'..............................................................

    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    swApp.CloseDoc swModel.GetTitle
        
    swApp.CloseDoc swModel.GetTitle
    
End Function








'>>>>>>>>>>>>


'    ' Create a plane perpendicular to the max area face
'    Set swPlane = swFeatMgr.CreatePlane(maxAreaFace, swPlaneNormalToFace, False, False)
'
'    ' Rotate the part to align with the plane
'    swModelDocExt.RotateAboutCenter swPlane, 90
'
'    ' Update standard views
'    Set swView = swModel.ActiveView
'    swView.ViewOrientation = swStandardViews_e.swIsometricView
'
'    ' Save the part
'    Set swSaveAsOptions = swApp.GetSaveAsOptions(swPart)
'    swSaveAsOptions.SaveAsVersion = swSaveAsCurrentVersion
'    swPart.SaveAs3 filePath, 0, 0, swSaveAsOptions, 0, 0
'    swModel.Close True
'
'    AdjustPartToMaxAreaPlane = "Part adjusted and saved successfully!"
'End Function

