Option Explicit
Public Const LINE_TYPE As Integer = 3001
Public Const CIRCLE_TYPE As Integer = 3002
Public Const ELLIPSE_TYPE As Integer = 3003
Public Const INTERSECTION_TYPE As Integer = 3004
Public Const BCURVE_TYPE As Integer = 3005
Public Const SPCURVE_TYPE As Integer = 3006
Public Const CONSTPARAM_TYPE As Integer = 3008
Public Const TRIMMED_TYPE As Integer = 3009
' Define two types
Type DoubleRec
    dValue As Double
End Type
Type Long2Rec
    iLower As Long
    iUpper As Long
End Type
' Extract two integer values out of a single double value,
' by assigning a DoubleRec to the double value and
' copying the value over an Long2Rec and
' extracting the integer values.
Function ExtractFields _
( _
    ByVal dValue As Double, _
    iLower As Long, _
    iUpper As Long _
)
    Dim dr As DoubleRec
    Dim i2r As Long2Rec
    ' Set the double value
    dr.dValue = dValue
    ' Copy the values
    LSet i2r = dr
    ' Extract the values
    iLower = i2r.iLower
    iUpper = i2r.iUpper
End Function
Sub main()
    Const sPathToTemplate As String = "c:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\templates\drawing.drtdot"
    Const nTolerance As Double = 0.00000001
    Const nXoffset As Double = 0.01
    Const nYoffset As Double = 0.01
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swEdge As SldWorks.Edge
    Dim swEnt As SldWorks.entity
    Dim swCurve As SldWorks.Curve
    Dim vCurveParam As Variant
    Dim nDummy As Long
    Dim nIdentity As Long
    Dim nTag As Long
    Dim nSense As Long
    Dim swMathUtil As SldWorks.MathUtility
    Dim nPtData(2) As Double
    Dim vPtData As Variant
    Dim swModelStartPt As SldWorks.MathPoint
    Dim swModelEndPt As SldWorks.MathPoint
    Dim swViewStartPt As SldWorks.MathPoint
    Dim swViewEndPt As SldWorks.MathPoint
    Dim i2r                         As Long2Rec

    ' Set the double value

    dr.dValue = dValue

    ' Copy the values

    LSet i2r = dr

    ' Extract the values

    iLower = i2r.iLower

    iUpper = i2r.iUpper

End Function

 

Sub main()

 

    Const sPathToTemplate           As String = "c:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\templates\drawing.drtdot"

    

    Const nTolerance                As Double = 0.00000001

    

    Const nXoffset                  As Double = 0.01

    Const nYoffset                  As Double = 0.01

    

    Dim swApp                       As SldWorks.SldWorks

    Dim swModel                     As SldWorks.ModelDoc2

    Dim swSelMgr                    As SldWorks.SelectionMgr

    Dim swEdge                      As SldWorks.Edge

    Dim swEnt                       As SldWorks.entity

    Dim swCurve                     As SldWorks.Curve

    Dim vCurveParam                 As Variant

    Dim nDummy                      As Long

    Dim nIdentity                   As Long

    Dim nTag                        As Long

    Dim nSense                      As Long

    

    Dim swMathUtil                  As SldWorks.MathUtility

    Dim nPtData(2)                  As Double

    Dim vPtData                     As Variant

    Dim swModelStartPt              As SldWorks.MathPoint

    Dim swModelEndPt                As SldWorks.MathPoint

    Dim swViewStartPt               As SldWorks.MathPoint

    Dim swViewEndPt                 As SldWorks.MathPoint

    

    Dim swDraw                      As SldWorks.DrawingDoc

    Dim swDrawModel                 As SldWorks.ModelDoc2

    Dim swView                      As SldWorks.view

    Dim swViewXform                 As SldWorks.MathTransform

    Dim vOutline                    As Variant

    Dim swDispDim                   As SldWorks.DisplayDimension

    

    Dim nXpos                       As Double

    Dim nYpos                       As Double

    Dim bRet                        As Boolean

    Set swApp = CreateObject("SldWorks.Application")

    Set swModel = swApp.ActiveDoc

    Set swSelMgr = swModel.SelectionManager

    Set swEdge = swSelMgr.GetSelectedObject5(1)

    Set swCurve = swEdge.GetCurve

    Set swEnt = swEdge

    

    vCurveParam = swEdge.GetCurveParams2

    ExtractFields vCurveParam(8), nDummy, nIdentity

    ExtractFields vCurveParam(9), nDummy, nTag

    ExtractFields vCurveParam(10), nDummy, nSense


    ' Derived quantity

    Debug.Print "Length     = " & swCurve.GetLength2(vCurveParam(6), vCurveParam(7)) * 1000# & " mm "

    Debug.Print ""

    

    ' Only makes sense for straight edges

    If LINE_TYPE <> nIdentity Then Exit Sub

    

    Set swMathUtil = swApp.GetMathUtility

    

    nPtData(0) = vCurveParam(0)

    nPtData(1) = vCurveParam(1)

    nPtData(2) = vCurveParam(2)

    vPtData = nPtData

    Set swModelStartPt = swMathUtil.CreatePoint(vPtData)

    

    nPtData(0) = vCurveParam(3)

    nPtData(1) = vCurveParam(4)

    nPtData(2) = vCurveParam(5)

    vPtData = nPtData

    Set swModelEndPt = swMathUtil.CreatePoint(vPtData)

    

    

    ' Start creating drawing of the model

    Set swDraw = swApp.NewDocument("C:\Program Files\SOLIDWORKS\data\templates\drawing.drwdot", swDwgPaperAsize, 0, 0)

    Set swDrawModel = swDraw

    

    bRet = swDraw.Create3rdAngleViews2(swModel.GetPathName)

    Debug.Assert bRet

    

    Set swView = swDraw.GetFirstView

    Set swView = swView.GetNextView

    Do While Not swView Is Nothing

        ' Select regardless

        bRet = swView.SelectEntity(swEnt, False)

        Debug.Assert bRet

        

        vOutline = swView.GetOutline

        

        Set swViewXform = swView.ModelToViewTransform

        Set swViewStartPt = swModelStartPt.MultiplyTransform(swViewXform)

        Set swViewEndPt = swModelEndPt.MultiplyTransform(swViewXform)

        

        Debug.Print "View       = " & swView.Name

        Debug.Print "Start      = (" & swViewStartPt.ArrayData(0) * 1000# & ", " & swViewStartPt.ArrayData(1) * 1000# & ", " & swViewStartPt.ArrayData(2) * 1000# & ") mm "

        Debug.Print "End        = (" & swViewEndPt.ArrayData(0) * 1000# & ", " & swViewEndPt.ArrayData(1) * 1000# & ", " & swViewEndPt.ArrayData(2) * 1000# & ") mm "

        Debug.Print ""

        

        If Abs(swViewStartPt.ArrayData(0) - swViewEndPt.ArrayData(0)) < nTolerance Then

            ' Must be vertical

            ' Place dimension midway up edge and to the right of view

            nXpos = vOutline(0) - nXoffset

            nYpos = Abs((swViewStartPt.ArrayData(1) + swViewEndPt.ArrayData(1)) / 2#)

            

            ' NULL if cannot convert edge in this view

            Set swDispDim = swDrawModel.AddVerticalDimension2(nXpos, nYpos, 0#)

        ElseIf Abs(swViewStartPt.ArrayData(1) - swViewEndPt.ArrayData(1)) < nTolerance Then

            ' Must be horizontal

            ' Place dimension midway across edge and above view

            nXpos = Abs((swViewStartPt.ArrayData(0) + swViewEndPt.ArrayData(0)) / 2#)

            nYpos = vOutline(3) + nYoffset

            

            ' NULL if cannot convert edge in this view

            Set swDispDim = swDrawModel.AddHorizontalDimension2(nXpos, nYpos, 0#)

        Else

            ' Neither horizontal or vertical

            ' Place dimension near middle of edge

            nXpos = Abs((swViewStartPt.ArrayData(0) + swViewEndPt.ArrayData(0)) / 2#) + nXoffset

            nYpos = Abs((swViewStartPt.ArrayData(1) + swViewEndPt.ArrayData(1)) / 2#) + nYoffset

    

            ' Depends on the orientation of the entity in the drawing view,

            ' thus, could be NULL

            '

            ' Create the dimension even if the entity is not

            ' visible in the drawing view

            Set swDispDim = swDrawModel.AddDimension2(nXpos, nYpos, 0#)

        End If

        

        Set swView = swView.GetNextView

    Loop

End Sub
