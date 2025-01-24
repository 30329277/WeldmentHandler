using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
namespace Dimensioning.csproj
{
    partial class SolidWorksMacro
    {
        ModelDoc2 swModel;
        DrawingDoc swDrawDoc;
        ModelDocExtension swModelDocExt = default(ModelDocExtension);
        View swView;
        CenterOfMass swCenterOfMass;
        Annotation swAnnotation;
        int sheetCount;
        int viewCount;
        object[] ss;
        object[] vv;
        double[] coord;
        Edge swEdge;
        bool status = false;

        public void Main()
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDrawDoc = (DrawingDoc)swModel;
            swModelDocExt = (ModelDocExtension)swModel.Extension;

            viewCount = swDrawDoc.GetViewCount();
            ss = (object[])swDrawDoc.GetViews();

            // 添加循环设置所有视图的切线为隐藏
            for (int i = ss.GetLowerBound(0); i <= ss.GetUpperBound(0); i++)
            {
                object[] views = (object[])ss[i];
                for (int j = 1; j <= views.GetUpperBound(0); j++)
                {
                    View view = (View)views[j];
                    view.SetDisplayTangentEdges2((int)swDisplayTangentEdges_e.swTangentEdgesHidden);
                }
            }

            for (sheetCount = ss.GetLowerBound(0); sheetCount <= ss.GetUpperBound(0); sheetCount++)
            {
                vv = (object[])ss[sheetCount];
                for (viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                {
                    Debug.Print(((View)vv[viewCount]).GetName2());
                    swView = (View)vv[viewCount];

                    MathTransform swViewXform1 = (MathTransform)swView.ModelToViewTransform;
                    double[] swViewXform2 = (double[])swView.GetViewXform();

                    object vEdgesOut;
                    object[] vEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);

                    int nCountShaded = 0;
                    int polyLineCount = swView.GetPolyLineCount5(1, out nCountShaded);
                    Debug.Print("Polyline count: " + polyLineCount);

                    double[] viewOutline = (double[])swView.GetOutline();
                    double viewWidth = Math.Abs(viewOutline[2] - viewOutline[0]);
                    double viewHeight = Math.Abs(viewOutline[3] - viewOutline[1]);
                    Debug.Print("View Width: " + Math.Round(viewWidth * 1000, 2) + " mm");
                    Debug.Print("View Height: " + Math.Round(viewHeight * 1000, 2) + " mm");

                    double[] viewPosition = (double[])swView.Position;
                    Debug.Print("View Position: X = " + Math.Round(viewPosition[0] * 1000, 2) + " mm, Y = " + Math.Round(viewPosition[1] * 1000, 2) + " mm");

                    if (polyLineCount == 16 ||
                        polyLineCount == 17 && vEdges.Count(e => e is Edge && ((Edge)e).GetCurveParams3().CurveType == 3002) == 8)
                    {
                        // 这个view是方管的右左视图，执行以下的for循环逻辑
                        DimensioningTubeSection(vEdges);
                        RemoveDuplicate(swView, swDrawDoc, 0, viewCount);
                        RelocateDimension(swView);
                    }
                    else if (Math.Round(viewWidth / viewHeight, 2) == 1 || Math.Round(viewWidth / viewHeight, 2) == 0.95 ||
                             Math.Round(viewWidth / viewHeight, 2) == Math.Round(100.0 / 50.0, 2) || Math.Round(viewWidth / viewHeight, 2) == Math.Round(50.0 / 100.0, 2) ||
                             Math.Round(viewWidth / viewHeight, 2) == Math.Round(100.0 / 60.0, 2) || Math.Round(viewWidth / viewHeight, 2) == Math.Round(60.0 / 100.0, 2) && polyLineCount >= 16)
                    {
                        CreateBrokenOutView(swView, swDrawDoc);
                        int maxAttempts = 20;
                        int attempts = 0;

                        while (attempts < maxAttempts)
                        {
                            CreateBrokenOutView(swView, swDrawDoc);
                            polyLineCount = swView.GetPolyLineCount5(1, out nCountShaded);
                            if (polyLineCount == 16 ||
                                polyLineCount == 17 && vEdges.Count(e => e is Edge && ((Edge)e).GetCurveParams3().CurveType == 3002) == 8)
                            {
                                object[] updatedEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);
                                DimensioningTubeSection(updatedEdges);
                                RemoveDuplicate(swView, swDrawDoc, 0, viewCount);
                                RelocateDimension(swView);
                                break;
                            }
                            attempts++;
                        }

                        // Add additional logic to execute after maxAttempts
                        if (attempts >= maxAttempts)
                        {
                            object[] updatedEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);
                            DimensioningTubeSection(updatedEdges);
                            RemoveDuplicate(swView, swDrawDoc, 0, viewCount);
                            RelocateDimension(swView);
                        }

                    }
                    else if (vEdges.Any(edge => edge is Edge && ((Edge)edge).GetCurveParams3().CurveType == 3002))
                    {
                        // 满足条件的代码逻辑
                        DimensioningHoles2(swView);
                        DimensioningTubeSide(vEdges);
                        Remove90And180DegreeDimensions(swView, swDrawDoc);
                    }

                    else
                    {
                        DimensioningTubeSide(vEdges);
                        Remove90And180DegreeDimensions(swView, swDrawDoc);
                    }

                }
            }

            // Save the drawing
            /*    int errors = 0;
                int warnings = 0;
                bool saveStatus = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                if (!saveStatus)
                {
                    Debug.Print("Failed to save the drawing. Errors: " + errors + ", Warnings: " + warnings);
                }
                else
                {
                    Debug.Print("Drawing saved successfully.");
                }*/
        }
        private void DimensioningTubeSection(object[] vEdges)
        {
            for (int i = 0; i <= vEdges.GetUpperBound(0); i++)
            {
                if (vEdges[i] is Edge swEdge)
                {
                    if (swEdge == null) continue; // Add this check

                    Curve swCurve = (Curve)swEdge.GetCurve();
                    CurveParamData swCurveParaData = (CurveParamData)swEdge.GetCurveParams3();

                    double[] curveParams = (double[])swEdge.GetCurveParams2();
                    object[] vCurveParam = new object[curveParams.Length];
                    for (int j = 0; j < curveParams.Length; j++)
                    {
                        vCurveParam[j] = curveParams[j];
                    }

                    double[] vLineParam = swCurve.LineParams as double[];
                    if (vLineParam != null)
                    {
                    }
                    else
                    {
                    }

                    if (swCurveParaData.CurveType == 3001)
                    {
                        swModel.ClearSelection2(true);
                        swView.SelectEntity(swEdge, true);

                        for (int j = 0; j <= vEdges.GetUpperBound(0); j++)
                        {
                            if (i == j) continue;

                            if (vEdges[j] is Edge edge2)
                            {
                                if (edge2 == null) continue; // Add this check
                                double[] curveParams2 = (double[])edge2.GetCurveParams2();
                                CurveParamData curveParams3 = (CurveParamData)edge2.GetCurveParams3();

                                Curve curve2 = (Curve)edge2.GetCurve();
                                double[] lineParams2 = curve2.LineParams as double[];

                                if (lineParams2 != null)
                                {
                                    bool areParallel =
                                    Math.Round(Math.Abs(vLineParam[3]), 2) == Math.Round(Math.Abs(lineParams2[3]), 2) &&
                                    Math.Round(Math.Abs(vLineParam[4]), 2) == Math.Round(Math.Abs(lineParams2[4]), 2) &&
                                    Math.Round(Math.Abs(vLineParam[5]), 2) == Math.Round(Math.Abs(lineParams2[5]), 2);

                                    double length1 = swCurve.GetLength2(curveParams[6], (double)curveParams[7]) * 1000;
                                    double length2 = curve2.GetLength2((double)curveParams2[6], (double)curveParams2[7]) * 1000;
                                    bool areEqualLength = Math.Round(length1, 2) == Math.Round(length2, 2);

                                    if (areParallel && areEqualLength)
                                    {
                                        swView.SelectEntity(edge2, true);

                                        // 计算放置尺寸的位置 - 视图中间并稍微向上
                                        double nXoffset = ((double[])curveParams3.EndPoint)[0] - ((double[])curveParams3.StartPoint)[0];
                                        double nYoffset = ((double[])curveParams3.EndPoint)[1] - ((double[])curveParams3.StartPoint)[1];
                                        double nZoffset = ((double[])curveParams3.EndPoint)[2] - ((double[])curveParams3.StartPoint)[2];

                                        //先把nOffset设置为0
                                        double nOffset = Math.Sqrt(Math.Pow(nXoffset, 2) + Math.Pow(nYoffset, 2) + Math.Pow(nZoffset, 2)) * 0;

                                        double[] vOutline = (double[])swView.GetOutline();
                                        double nXpos = (vOutline[0] + vOutline[2]) / 2.0 + nOffset;
                                        double nYpos = vOutline[3] + nOffset;

                                        // 创建尺寸，即使实体在图纸视图中不可见
                                        swModel.AddDimension2(nXpos, nYpos, 0.0);

                                        // 重新选中 edge1 继续遍历
                                        swView.SelectEntity(swEdge, false);

                                    }
                                    else
                                    {
                                        // Debug.Print("不平行或者不相等");
                                    }
                                }
                                else
                                {
                                    // Debug.Print("不是直线");
                                }
                            }
                        }
                    }
                }
            }
        }
        private void RemoveDuplicate(View view, DrawingDoc draw, int fileNmb, int viewNum)
        {
            DisplayDimension swDispDim = view.GetFirstDisplayDimension5();
            Sheet swSheet = view.Sheet;

            if (swSheet == null)
            {
                swSheet = draw.Sheet[view.Name];
            }

            List<DisplayDimension> dimensions = new List<DisplayDimension>();
            while (swDispDim != null)
            {
                dimensions.Add(swDispDim);
                swDispDim = swDispDim.GetNext5();
            }

            // Sort dimensions by their value in descending order
            dimensions.Sort((dim1, dim2) =>
            {
                double val1 = Math.Round(dim1.GetDimension2(0).GetValue2(""), 3);
                double val2 = Math.Round(dim2.GetDimension2(0).GetValue2(""), 3);
                return val2.CompareTo(val1);
            });

            // Count occurrences of each value
            Dictionary<double, int> valueCounts = new Dictionary<double, int>();
            List<double> uniqueValues = new List<double>();
            foreach (var dim in dimensions)
            {
                double value = Math.Round(dim.GetDimension2(0).GetValue2(""), 3);
                if (!valueCounts.ContainsKey(value))
                {
                    valueCounts[value] = 1;
                    uniqueValues.Add(value);
                }
                else
                {
                    valueCounts[value]++;
                }
            }

            // 获取所有独特的值并按升序排序
            uniqueValues = valueCounts.Keys.ToList();
            uniqueValues.Sort();

            double maxValue = uniqueValues[uniqueValues.Count - 1];
            double minValue = uniqueValues[0];

            // 计算真正的中间值：找到最接近 (maxValue + minValue) / 2 的值，并且优先选择 10 的倍数
            double targetMidValue = (maxValue + minValue) / 2;
            double midValue = uniqueValues
                .Where(v => v % 10 == 0)
                .OrderBy(v => Math.Abs(v - targetMidValue))
                .DefaultIfEmpty(uniqueValues.OrderBy(v => Math.Abs(v - targetMidValue)).First())
                .First();

            // Process dimensions based on case
            bool isKeeping;
            int maxValueKept = 0;
            int midValueKept = 0;
            int minValueKept = 0;

            foreach (var dim in dimensions)
            {
                double value = Math.Round(dim.GetDimension2(0).GetValue2(""), 3);
                isKeeping = false;

                if (valueCounts[maxValue] >= 4)  // Case A
                {
                    if (value == maxValue && maxValueKept < 3)  // 改为保留3个最大值
                    {
                        isKeeping = true;
                        maxValueKept++;
                    }
                    else if (value == minValue && minValueKept < 1)
                    {
                        isKeeping = true;
                        minValueKept++;
                    }
                }
                else  // Case B
                {
                    if (value == maxValue && maxValueKept < 1)
                    {
                        isKeeping = true;
                        maxValueKept++;
                    }
                    else if (value == midValue && midValueKept < 1)
                    {
                        isKeeping = true;
                        midValueKept++;
                    }
                    else if (value == minValue && minValueKept < 1)
                    {
                        isKeeping = true;
                        minValueKept++;
                    }
                }

                if (!isKeeping)
                {
                    Annotation annotation = (Annotation)dim.GetAnnotation();
                    // Debug.Print("Deleting dimension: " + annotation.GetName());
                    annotation.Select2(false, 0);
                    swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                }
            }
        }
        private void DimensioningTubeSide(object[] vEdges)
        {
            List<Edge> largestEdges = new List<Edge>();
            double maxLength = 0;
            double[] largestLineParams = null;

            // First pass: find the largest edges
            for (int i = 0; i <= vEdges.GetUpperBound(0); i++)
            {
                if (vEdges[i] is Edge swEdge)
                {
                    // Proceed with the rest of the logic
                }
                else
                {
                    // Handle the case where the object is not an Edge
                    continue;
                }
                swEdge = (Edge)vEdges[i];
                Curve swCurve = (Curve)swEdge.GetCurve();
                CurveParamData swCurveParaData = (CurveParamData)swEdge.GetCurveParams3();

                if (swCurveParaData.CurveType != 3001) // Skip non-line edges
                {
                    continue;
                }

                double[] curveParams = (double[])swEdge.GetCurveParams2();
                double length = swCurve.GetLength2(curveParams[6], curveParams[7]) * 1000;

                if (length > maxLength)
                {
                    maxLength = length;
                    largestEdges.Clear();
                    largestEdges.Add(swEdge);
                    largestLineParams = (double[])swCurve.LineParams;
                }
                else if (length == maxLength)
                {
                    largestEdges.Add(swEdge);
                }
            }

            if (largestEdges.Count > 0)
            {
                // Select one of the largest edges
                swModel.ClearSelection2(true);
                swView.SelectEntity(largestEdges[0], true);

                double[] vOutline = (double[])swView.GetOutline();
                double nXpos = (vOutline[0] + vOutline[2]) / 2.0;
                double nYpos = vOutline[3];

                // Add dimension for the largest edge
                swModel.AddDimension2(nXpos, nYpos, 0.0);

                int angleDimensionsAdded = 0;

                // Second pass: find edges that form an angle with the largest edge, excluding 90° and 180° angles
                for (int i = 0; i <= vEdges.GetUpperBound(0); i++)
                {
                    if (vEdges[i] is Edge swEdge)
                    {
                        // Proceed with the rest of the logic
                    }
                    else
                    {
                        // Handle the case where the object is not an Edge
                        continue;
                    }
                    swEdge = (Edge)vEdges[i];
                    if (largestEdges.Contains(swEdge))
                    {
                        continue;
                    }

                    Curve swCurve = (Curve)swEdge.GetCurve();
                    CurveParamData swCurveParaData = (CurveParamData)swEdge.GetCurveParams3();

                    if (swCurveParaData.CurveType != 3001) // Skip non-line edges
                    {
                        continue;
                    }

                    double[] lineParams = (double[])swCurve.LineParams;

                    if (lineParams != null && largestLineParams != null)
                    {
                        double dotProduct = largestLineParams[3] * lineParams[3] +
                                            largestLineParams[4] * lineParams[4] +
                                            largestLineParams[5] * lineParams[5];

                        double magnitude1 = Math.Sqrt(Math.Pow(largestLineParams[3], 2) +
                                                      Math.Pow(largestLineParams[4], 2) +
                                                      Math.Pow(largestLineParams[5], 2));

                        double magnitude2 = Math.Sqrt(Math.Pow(lineParams[3], 2) +
                                                      Math.Pow(lineParams[4], 2) +
                                                      Math.Pow(lineParams[5], 2));

                        double angle = Math.Acos(dotProduct / (magnitude1 * magnitude2)) * (180.0 / Math.PI);

                        // Debug.Print("Angle between edges: " + angle);

                        if ((angle > 0 && angle < 90) || (angle > 90 && angle < 180)) // Exclude 90° and 180° angles
                        {
                            // Clear selection and select the first edge again
                            swModel.ClearSelection2(true);
                            swView.SelectEntity(largestEdges[0], true);
                            // Select the second edge
                            swView.SelectEntity(swEdge, true);

                            // Adjust nXpos to place the dimension closer to the left or right edge
                            if (angleDimensionsAdded == 0)
                            {
                                nXpos = vOutline[0] + (vOutline[2] - vOutline[0]) * 0.1; // 10% from the left edge
                            }
                            else
                            {
                                nXpos = vOutline[2] - (vOutline[2] - vOutline[0]) * 0.1; // 10% from the right edge
                            }

                            // Add dimension between the two selected edges
                            swModel.AddDimension2(nXpos, nYpos, 0.0);
                            angleDimensionsAdded++;
                            if (angleDimensionsAdded == 2)
                            {
                                break; // Only add two angle dimensions
                            }
                        }
                    }
                }
            }

            swModel.ClearSelection2(true);
        }
        private void Remove90And180DegreeDimensions(View view, DrawingDoc draw)
        {
            DisplayDimension swDispDim = view.GetFirstDisplayDimension5();
            Sheet swSheet = view.Sheet;

            if (swSheet == null)
            {
                swSheet = draw.Sheet[view.Name];
            }

            List<DisplayDimension> dimensions = new List<DisplayDimension>();
            while (swDispDim != null)
            {
                dimensions.Add(swDispDim);
                swDispDim = swDispDim.GetNext5();
            }

            foreach (var dim in dimensions)
            {
                double value = dim.GetDimension2(0).GetValue2("");
                if (Math.Abs(value - 90.0) < 0.001 || Math.Abs(value - 180.0) < 0.001) // Check if the dimension value is 90 or 180 degrees
                {
                    Annotation annotation = (Annotation)dim.GetAnnotation();
                    // Debug.Print("Deleting 90 or 180 degree dimension: " + annotation.GetName());
                    annotation.Select2(false, 0);
                    swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                }
            }
            swModel.ClearSelection2(true);
        }
        private void CreateBrokenOutView(View swView, DrawingDoc swDrawDoc)
        {
            object vEdgesOut;
            object[] vEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);

            double[] viewEndData = null;

            foreach (object edgeObj in vEdges)
            {
                if (edgeObj is Edge edge)
                {
                    Curve curve = (Curve)edge.GetCurve();
                    CurveParamData curveData = edge.GetCurveParams3();

                    if (curveData.CurveType == 3002) // Circle
                    {
                        double[] circleParams = curve.CircleParams as double[];
                        if (circleParams != null)
                        {
                            // Get circle center point
                            double[] centerPoint = new double[] { circleParams[0], circleParams[1], circleParams[2] };

                            // Create MathPoint for center point
                            MathUtility swMathUtil = (MathUtility)swApp.GetMathUtility();
                            MathPoint swModelCenterPt = (MathPoint)swMathUtil.CreatePoint(centerPoint);

                            // Transform center point to view space
                            MathTransform swViewXform = (MathTransform)swView.ModelToViewTransform;
                            MathPoint swViewCenterPt = (MathPoint)swModelCenterPt.MultiplyTransform(swViewXform);

                            // Get transformed coordinates for center point 
                            viewEndData = (double[])swViewCenterPt.ArrayData;

                            Debug.Print($"Circle center point in view space: ({viewEndData[0]}, {viewEndData[1]}, {viewEndData[2]})");
                            Debug.Print($"Circle radius: {circleParams[6]}");

                            break; // Exit after finding first circle
                        }
                    }
                }
            }

            try
            {
                // Check if there is an active view and deactivate it
                View activeView = (View)swDrawDoc.ActiveDrawingView;
                swDrawDoc.ActivateView(swView.Name);

                if (activeView != null && activeView.Name != swView.Name)
                {
                    Debug.Print("Active View Name: " + activeView.GetName2());
                    Debug.Print("swView Name: " + swView.GetName2());
                    swDrawDoc.ActivateView(swView.Name);
                }
                else
                {
                    Debug.Print("No active view.");
                }

                SketchManager swSketchMgr = swModel.SketchManager;
                SketchSegment swSketchSeg;

                // Get sheet properties for scaling 
                Sheet swSheet = swView.Sheet;
                double[] sheetProps = (double[])swSheet.GetProperties();
                double sheetScale = sheetProps[3];

                // Get view center coordinates
                double[] viewOutline = (double[])swView.GetOutline();
                double viewWidth = Math.Abs(viewOutline[2] - viewOutline[0]);
                double viewHeight = Math.Abs(viewOutline[3] - viewOutline[1]);

                //swSketchMgr.InsertSketch(true);
                swSketchSeg = swSketchMgr.CreateCircleByRadius(0, 0, 0, Math.Round((viewWidth + viewHeight), 2) * 3);
                //swSketchMgr.InsertSketch(false);

                // Select the line
                if (swSketchSeg != null)
                {
                    swSketchSeg.Select4(false, null);
                }

                double[] swViewXform2 = (double[])swView.GetViewXform();

                if (viewEndData != null && swViewXform2 != null && swViewXform2.Length > 12)
                {
                    // bool status = swDrawDoc.CreateBreakOutSection(viewEndData[2] / swViewXform2[12] * 0.5);
                    int attempts = 0; // Define the attempts variable
                    attempts++; // Increment the attempts variable

                    if (viewEndData != null && viewEndData.Length > 2 && swViewXform2 != null && swViewXform2.Length > 12 && swViewXform2[12] != 0)
                    {
                        double cutDepth = viewEndData[2] / swViewXform2[12] * (0.5 + attempts * 0.5);
                        Debug.Print($"Cut Depth: {cutDepth}");
                        bool status = swDrawDoc.CreateBreakOutSection(cutDepth);
                    }
                    else
                    {
                        Debug.Print("Invalid data for calculating cut depth.");
                    }
                
                }
                else
                {
                    Debug.Print("viewEndData or swViewXform2 is not properly initialized.");
                }

            }
            catch (Exception ex)
            {
                Debug.Print($"Error creating broken-out view: {ex.Message}");
            }
            swModel.ClearSelection2(true);
        }
        private void RelocateDimension(View view)
        {
            DisplayDimension currentDim = view.GetFirstDisplayDimension5();
            List<DisplayDimension> dimensions = new List<DisplayDimension>();

            // Collect all dimensions
            while (currentDim != null)
            {
                dimensions.Add(currentDim);
                currentDim = currentDim.GetNext5();
            }

            // Sort dimensions by vertical position
            dimensions.Sort((a, b) =>
            {
                double[] posA = (double[])((Annotation)a.GetAnnotation()).GetPosition();
                double[] posB = (double[])((Annotation)b.GetAnnotation()).GetPosition();
                return posA[1].CompareTo(posB[1]);
            });

            const double minVerticalGap = 0.015; // Minimum vertical gap between dimensions

            double[] vOutline = (double[])view.GetOutline();
            double viewCenterX = (vOutline[0] + vOutline[2]) / 2.0;
            double viewWidth = Math.Abs(vOutline[2] - vOutline[0]);

            // Auto-arrange all dimensions in the view
            swModel.Extension.AlignDimensions((int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.06);

            // Adjust positions to distribute dimensions around the view
            for (int i = 0; i < dimensions.Count; i++)
            {
                // Check if this is an angle dimension
                bool isAngleDimension = dimensions[i].GetDimension2(0).GetType() == (int)swDimensionType_e.swAngularDimension;

                // Skip angle dimensions
                if (isAngleDimension)
                {
                    continue;
                }

                Annotation currentAnnotation = (Annotation)dimensions[i].GetAnnotation();
                double[] currentPos = (double[])currentAnnotation.GetPosition();

                // Calculate new position based on index
                if (i % 2 == 0)
                {
                    // Even indices - place on left side
                    currentPos[0] = vOutline[0] - viewWidth * 0.05;
                }
                else
                {
                    // Odd indices - place on right side
                    currentPos[0] = vOutline[2] + viewWidth * 0.05;
                }

                // Adjust vertical position with consistent spacing
                currentPos[1] = vOutline[3] - (i * minVerticalGap);

                // Apply new position
                currentAnnotation.SetPosition2(currentPos[0], currentPos[1], currentPos[2]);
            }


        }
        private void AddOrdinateDimension(View swView)
        {
            try
            {
                ModelDocExtension swModelDocExt = swModel.Extension;
                SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
                DisplayDimension swDisplayDimension;
                bool status = false;
                int errors = 0;

                // Get all edges
                object vEdgesOut;
                object[] vEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);

                // Process circles
                foreach (object edgeObj in vEdges)
                {
                    Edge edge = (Edge)edgeObj;
                    Curve curve = (Curve)edge.GetCurve();
                    CurveParamData curveData = edge.GetCurveParams3();

                    if (curveData.CurveType == 3002) // Circle
                    {
                        double[] center = curve.CircleParams as double[];
                        if (center != null)
                        {
                            // Add horizontal ordinate dimension
                            swModel.ClearSelection2(true);
                            status = swModelDocExt.SelectByID2("", "VERTEX", center[0], center[1], 0, false, 0, null, 0);
                            errors = swModelDocExt.AddOrdinateDimension(
                                (int)swAddOrdinateDims_e.swHorizontalOrdinate,
                                center[0],
                                center[1] + 0.02,
                                0
                            );

                            // Add vertical ordinate dimension
                            swModel.ClearSelection2(true);
                            status = swModelDocExt.SelectByID2("", "VERTEX", center[0], center[1], 0, false, 0, null, 0);
                            errors = swModelDocExt.AddOrdinateDimension(
                                (int)swAddOrdinateDims_e.swVerticalOrdinate,
                                center[0] - 0.02,  // Offset to the left
                                center[1],
                                0
                            );

                            // Set arrow size if needed
                            if (swModel.Extension.SelectByID2("D1@" + swView.GetName2(), "DIMENSION", center[0], center[1], 0, false, 0, null, 0))
                            {
                                swDisplayDimension = (DisplayDimension)swSelMgr.GetSelectedObject6(1, -1);
                                swDisplayDimension.SetOrdinateDimensionArrowSize(false, 0.003);
                            }
                        }
                    }
                }

                swModel.ClearSelection2(true);
                Debug.Print("Ordinate dimensioning completed successfully");
            }
            catch (Exception ex)
            {
                Debug.Print($"Error adding ordinate dimensions: {ex.Message}");
                swModel.ClearSelection2(true);
            }
        }
        private void DimensioningHoles(View swView)
        {
            object vEdgesOut;
            object[] vEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            Edge originEdge = null;

            // Find the edge closest to bottom-left corner
            foreach (object edgeObj in vEdges)
            {
                Edge edge = (Edge)edgeObj;
                Curve curve = (Curve)edge.GetCurve();
                double[] lineParams = curve.LineParams as double[];

                if (lineParams != null)
                {
                    if (lineParams[0] <= minX && lineParams[1] <= minY)
                    {
                        minX = lineParams[0];
                        minY = lineParams[1];
                        originEdge = edge;
                    }
                }
            }

            // Create a list to store already dimensioned locations
            HashSet<string> dimensionedLocations = new HashSet<string>();

            // Process each edge
            foreach (object edgeObj in vEdges)
            {
                Edge edge = (Edge)edgeObj;
                Curve curve = (Curve)edge.GetCurve();
                CurveParamData curveData = edge.GetCurveParams3();

                // Check for circular edges (holes)
                if (curveData.CurveType == 3002) // Arc/Circle type
                {
                    double[] center = curve.CircleParams as double[];
                    if (center != null)
                    {
                        // Get view position
                        double[] viewPos = (double[])swView.Position;
                        // Print view position coordinates
                        if (viewPos != null)
                        {
                            Debug.Print($"View Position: ({viewPos[0]}, {viewPos[1]}");
                        }
                        // Get view outline bounds
                        double[] outlineBounds = (double[])swView.GetOutline();
                        if (outlineBounds != null)
                        {
                            Debug.Print($"View Outline: Left={outlineBounds[0]}, Top={outlineBounds[1]}, Right={outlineBounds[2]}, Bottom={outlineBounds[3]}");
                        }

                        // Print circle center coordinates 
                        Debug.Print($"Circle Center: ({center[0]}, {center[1]}, {center[2]})");

                        string locationKey = $"{Math.Round(center[0], 6)},{Math.Round(center[1], 6)}";

                        if (!dimensionedLocations.Contains(locationKey))
                        {
                            // Add diameter dimension closer to the circle
                            swView.SelectEntity(edge, true);
                            //swModel.AddDimension2(center[0] + 0.01, center[1] + 0.01, 0);
                            swModel.AddDimension2(outlineBounds[0], outlineBounds[3], 0);
                            swModel.ClearSelection2(true);

                            if (originEdge != null)
                            {
                                // Add vertical dimension closer to circle
                                swModel.ClearSelection2(true);
                                swView.SelectEntity(originEdge, false);
                                swView.SelectEntity(edge, true);

                                CurveParamData curveParams3 = edge.GetCurveParams3();
                                double[] curveParams2 = (double[])edge.GetCurveParams2();

                                Debug.Print("GetCurveParams3: " + curveParams3.CurveType);
                                Debug.Print("GetCurveParams2: " + string.Join(", ", curveParams2));
                                // swModel.AddVerticalDimension2(center[0] - 0.01, center[1], 1);
                                //swModel.AddVerticalDimension2(outlineBounds[0], center[3], 1);
                                swModel.AddVerticalDimension2(outlineBounds[0], outlineBounds[3], 0);


                                // Select the leftmost point of originEdge
                                Curve originCurve = (Curve)originEdge.GetCurve();
                                double[] originEdgeParams = originCurve.LineParams as double[];
                                if (originEdgeParams != null)
                                {
                                    Vertex startVertex = (Vertex)originEdge.GetStartVertex();
                                    Vertex endVertex = (Vertex)originEdge.GetEndVertex();

                                    double[] startPoint = (double[])startVertex.GetPoint();
                                    double[] endPoint = (double[])endVertex.GetPoint();

                                    swModel.ClearSelection2(true);
                                    if (swView.GetOrientationName().ToLower().Contains("front"))
                                    {
                                        swView.SelectEntity(startVertex, true);
                                    }
                                    else
                                    {
                                        swView.SelectEntity(endVertex, true);
                                    }

                                }
                                swView.SelectEntity(edge, true);
                                // swModel.AddHorizontalDimension2(center[0], center[1] - 0.01, 1);
                                //swModel.AddHorizontalDimension2(center[1], outlineBounds[3], 1);
                                swModel.AddHorizontalDimension2(outlineBounds[1], outlineBounds[3], 0);

                            }

                            dimensionedLocations.Add(locationKey);
                        }
                    }
                }
            }
            swModel.ClearSelection2(true);
            swModel.SetPickMode();
        }
        private void DimensioningHoles2(View swView)
        {
            object vEdgesOut;
            object[] vEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            Edge originEdge = null;

            // Find the edge closest to bottom-left corner
            foreach (object edgeObj in vEdges)
            {
                if (edgeObj is Edge edge)
                {
                    Curve curve = (Curve)edge.GetCurve();
                    double[] lineParams = curve.LineParams as double[];

                    if (lineParams != null)
                    {
                        if (lineParams[0] <= minX && lineParams[1] <= minY)
                        {
                            minX = lineParams[0];
                            minY = lineParams[1];
                            originEdge = edge;
                        }
                    }
                }
            }

            // Create a list to store already dimensioned locations
            HashSet<string> dimensionedLocations = new HashSet<string>();

            // Get view position
            double[] viewPos = (double[])swView.Position;
            // Print view position coordinates
            // if (viewPos != null)
            // {
            //     Debug.Print($"View Position: ({viewPos[0]}, {viewPos[1]}");
            // }
            // Get view outline bounds
            double[] outlineBounds = (double[])swView.GetOutline();
            // if (outlineBounds != null)
            // {
            //     Debug.Print($"View Outline: Left={outlineBounds[0]}, Top={outlineBounds[1]}, Right={outlineBounds[2]}, Bottom={outlineBounds[3]}");
            // }

            // Process each edge
            foreach (object edgeObj in vEdges)
            {
                if (edgeObj is Edge edge)
                {
                    Curve curve = (Curve)edge.GetCurve();
                    CurveParamData curveData = edge.GetCurveParams3();

                    // Check for circular edges (holes)
                    if (curveData.CurveType == 3002)
                    {
                        double[] circleParams = curve.CircleParams as double[];
                        if (circleParams != null &&
                            Math.Round(circleParams[3], 6) == 0 &&
                            Math.Round(circleParams[4], 6) == 1 &&
                            Math.Round(circleParams[5], 6) == 0)
                        {
                            // double[] circleParams = curve.CircleParams as double[];
                            if (circleParams != null)
                            {
                                Debug.Print($"Circle Params: Center=({circleParams[0]}, {circleParams[1]}, {circleParams[2]}), Axis=({circleParams[3]}, {circleParams[4]}, {circleParams[5]}), Radius={circleParams[6]}");
                            }

                            double[] center = curve.CircleParams as double[];
                            if (center != null)
                            {
                                // Print circle center coordinates 
                                //Debug.Print($"Circle Center: ({center[0]}, {center[1]}, {center[2]})");

                                // Create MathUtility object
                                MathUtility swMathUtil = (MathUtility)swApp.GetMathUtility();

                                // Create MathPoint objects for the center point
                                double[] centerData = new double[] { center[0], center[1], center[2] };
                                MathPoint swModelCenterPt = (MathPoint)swMathUtil.CreatePoint(centerData);

                                // Get the transformation matrix for the view
                                MathTransform swViewXform = (MathTransform)swView.ModelToViewTransform;

                                // Transform the center point to view space
                                MathPoint swViewCenterPt = (MathPoint)swModelCenterPt.MultiplyTransform(swViewXform);

                                // Extract the transformed coordinates
                                double[] viewCenterData = (double[])swViewCenterPt.ArrayData;
                                double viewCenterX = viewCenterData[0];
                                double viewCenterY = viewCenterData[1];
                                double viewCenterZ = viewCenterData[2];

                                // Calculate the paper center coordinates
                                //double paperCenterX = viewCenterX + viewPos[0] - outlineBounds[0];
                                double paperCenterX = viewCenterX + (viewPos[0] - outlineBounds[0]) * 0.2;
                                //double paperCenterY = viewCenterY + viewPos[1] - outlineBounds[1];
                                double paperCenterY = viewCenterY + (viewPos[1] - outlineBounds[1]) * 0.6;
                                double paperCenterZ = viewCenterZ;

                                // Print the paper center coordinates
                                Debug.Print($"Paper Center: ({paperCenterX}, {paperCenterY}, {paperCenterZ})");

                                string locationKey = $"{Math.Round(center[0], 6)},{Math.Round(center[1], 6)}";

                                if (!dimensionedLocations.Contains(locationKey))
                                {
                                    // Add diameter dimension closer to the circle
                                    swView.SelectEntity(edge, true);
                                    swModel.AddDimension2(paperCenterX, paperCenterY, 0);
                                    swModel.ClearSelection2(true);

                                    if (originEdge != null)
                                    {
                                        // Add vertical dimension closer to circle
                                        swModel.ClearSelection2(true);
                                        swView.SelectEntity(originEdge, false);
                                        swView.SelectEntity(edge, true);

                                        // Select the leftmost point of originEdge
                                        Curve originCurve = (Curve)originEdge.GetCurve();
                                        double[] originEdgeParams = originCurve.LineParams as double[];
                                        if (originEdgeParams != null)
                                        {
                                            Vertex startVertex = (Vertex)originEdge.GetStartVertex();
                                            Vertex endVertex = (Vertex)originEdge.GetEndVertex();

                                            double[] startPoint = (double[])startVertex.GetPoint();
                                            double[] endPoint = (double[])endVertex.GetPoint();

                                            swModel.ClearSelection2(true);
                                            if (swView.GetOrientationName().ToLower().Contains("front") ||
                                                swView.GetOrientationName().ToLower().Contains("top") ||
                                                swView.GetOrientationName().ToLower().Contains("left"))
                                            {
                                                swView.SelectEntity(startVertex, true);
                                            }
                                            else
                                            {
                                                swView.SelectEntity(endVertex, true);
                                            }
                                        }
                                        swView.SelectEntity(edge, true);
                                        // DisplayDimension swDispDim = (DisplayDimension)swModel.AddHorizontalDimension2(outlineBounds[3], paperCenterY,  0);
                                        //DisplayDimension swDispDim = (DisplayDimension)swModel.AddHorizontalDimension2(outlineBounds[3], 0.3 + paperCenterX / 5, 0);
                                        DisplayDimension horizontalDim = (DisplayDimension)swModel.AddHorizontalDimension2(viewCenterX, viewCenterY, 0);

                                        if (horizontalDim != null)
                                        {
                                            horizontalDim.OffsetText = true;
                                            /*horizontalDim.CenterText = true;
                                            horizontalDim.OffsetText = false;*/
                                        }
                                        else
                                        {
                                            Debug.Print("Failed to create DisplayDimension.");
                                        }
                                    }

                                    if (originEdge != null)
                                    {
                                        // Add vertical dimension closer to circle
                                        swModel.ClearSelection2(true);
                                        swView.SelectEntity(originEdge, false);
                                        swView.SelectEntity(edge, true);

                                        // Select the leftmost point of originEdge
                                        Curve originCurve = (Curve)originEdge.GetCurve();
                                        double[] originEdgeParams = originCurve.LineParams as double[];
                                        if (originEdgeParams != null)
                                        {
                                            Vertex startVertex = (Vertex)originEdge.GetStartVertex();
                                            Vertex endVertex = (Vertex)originEdge.GetEndVertex();

                                            double[] startPoint = (double[])startVertex.GetPoint();
                                            double[] endPoint = (double[])endVertex.GetPoint();

                                            swModel.ClearSelection2(true);
                                            if (swView.GetOrientationName().ToLower().Contains("front") ||
                                             swView.GetOrientationName().ToLower().Contains("top") ||
                                             swView.GetOrientationName().ToLower().Contains("left"))
                                            {
                                                swView.SelectEntity(startVertex, true);
                                            }
                                            else
                                            {
                                                swView.SelectEntity(endVertex, true);
                                            }
                                        }
                                        swView.SelectEntity(edge, true);
                                        DisplayDimension verticalDim = (DisplayDimension)swModel.AddVerticalDimension2(viewCenterX, paperCenterY, 0);
                                        if (verticalDim != null)
                                        {
                                            verticalDim.OffsetText = true;
                                            /*verticalDim.CenterText = true;
                                            verticalDim.OffsetText = false;*/
                                        }
                                        else
                                        {
                                            Debug.Print("Failed to create DisplayDimension.");
                                        }
                                    }
                                    dimensionedLocations.Add(locationKey);
                                }
                            }
                        }
                    }

                    swModel.ClearSelection2(true);
                    swModel.SetPickMode();
                }
                // Select all dimensions in current view
                DisplayDimension swDispDim = swView.GetFirstDisplayDimension5();
                while (swDispDim != null)
                {
                    Annotation annotation = (Annotation)swDispDim.GetAnnotation();
                    annotation.Select2(true, 0);
                    swDispDim = swDispDim.GetNext5();
                }
                swModel.Extension.AlignDimensions((int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.06);
            }
        }

        public SldWorks swApp;
    }
}
