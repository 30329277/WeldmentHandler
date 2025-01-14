Attribute VB_Name = "sw14_custprops1"
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
'把sw06_custprops1 的dim 复制过来
'Dim swApp As SldWorks.SldWorks
'Dim swModel As ModelDoc2
'Dim swModelDocExt As ModelDocExtension
Dim swCustProp As CustomPropertyManager
Dim vConfigNames As Variant
Dim strValue As String
Dim strResolved As String
Dim lErrors As Long
Dim lWarnings As Long
'把sw06_custprops1 的dim 复制过来


Public Function CustomProperty(strFilePath As String, Material As String, Mat_fiche As String, Etat_surface As String, Description_EN As String, Description As String, SAPmaterial As String, SAPVersion As String) As String
'    Set swApp = Application.SldWorks
    Set swApp = CreateObject("SldWorks.Application")
'    Set swModel = swApp.ActiveDoc
    Set swDocSpec = swApp.GetOpenDocSpec(strFilePath)
    Set swModel = swApp.OpenDoc7(swDocSpec)
    Set swModelDocExt = swModel.Extension
    
        '增加一行代码断开所有链接
        swModel.BreakAllExternalReferences
        '增加一行代码断开所有链接
        
        ' Get the custom property data
        Set swCustProp = swModelDocExt.CustomPropertyManager("")
        If Material = "" Then Material = "X2 Cr Ni 18-9"
        swCustProp.Add3 "Matière", swCustomInfoText, Material, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        If Mat_fiche = "" Then Mat_fiche = "1-3-01"
        swCustProp.Add3 "Mat_fiche", swCustomInfoText, Mat_fiche, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        If Etat_surface = "" Then Etat_surface = "12.5"
        swCustProp.Add3 "Etat_surface", swCustomInfoText, Etat_surface, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "Dessiné_par", swCustomInfoText, "WANGNING", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "Date_de_création", swCustomInfoText, Date, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "Désignation_ang", swCustomInfoText, Description_EN, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "Description", swCustomInfoText, Description, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "SAPmaterial", swCustomInfoText, SAPmaterial, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "SAPVersion", swCustomInfoText, SAPVersion, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        
        swCustProp.Add3 "Traitement_1_fr", swCustomInfoText, "100130890", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "Traitement_2_fr", swCustomInfoText, "10719757", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        swCustProp.Add3 "Traitement_3_fr", swCustomInfoText, "100017110", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        
        'access cust props at config level
        vConfigNames = swModel.GetConfigurationNames
        Dim i As Integer
        For i = 0 To UBound(vConfigNames)
            Set swCustProp = swModel.Extension.CustomPropertyManager(vConfigNames(i))
            swCustProp.Add2 "PartNo", swCustomInfoText, strValue & "-" & vConfigNames(i)
            swCustProp.Add2 "Material", swCustomInfoText, Chr(34) & "SW-Material@" & swModel.GetTitle & Chr(34)
        Next i
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    swApp.CloseDoc swModel.GetTitle
    
End Function









