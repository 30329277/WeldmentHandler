using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
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

            for (sheetCount = ss.GetLowerBound(0); sheetCount <= ss.GetUpperBound(0); sheetCount++)
            {
                vv = (object[])ss[sheetCount];
                for (viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                {
                    Debug.Print(((View)vv[viewCount]).GetName2());
                    swView = (View)vv[viewCount];

                    object vEdgesOut;
                    object[] vEdges = (object[])swView.GetPolylines7(1, out vEdgesOut);

                    int nCountShaded = 0;
                    int polyLineCount = swView.GetPolyLineCount5(1, out nCountShaded);

                    if (polyLineCount == 16)
                    {
                        // 这个view是方管的右左视图，执行以下的for循环逻辑
                        DimensioningTubeSection(vEdges);
                        RemoveDuplicate(swView, swDrawDoc, 0, viewCount);
                    }
                    else
                    {
                        DimensioningTubeSide(vEdges);
                    }
                }
            }

            // Save the drawing
            int errors = 0;
            int warnings = 0;
            bool saveStatus = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
            if (!saveStatus)
            {
                Debug.Print("Failed to save the drawing. Errors: " + errors + ", Warnings: " + warnings);
            }
            else
            {
                Debug.Print("Drawing saved successfully.");
            }
        }

        private void DimensioningTubeSection(object[] vEdges)
        {
            for (int i = 0; i <= vEdges.GetUpperBound(0); i++)
            {
                swEdge = (Edge)vEdges[i];

                Curve swCurve = (Curve)swEdge.GetCurve();
                CurveParamData swCurveParaData = (CurveParamData)swEdge.GetCurveParams3();

                double[] curveParams = (double[])swEdge.GetCurveParams2();
                object[] vCurveParam = new object[curveParams.Length];
                for (int j = 0; j < curveParams.Length; j++)
                {
                    vCurveParam[j] = curveParams[j];
                }

                // Debug.Print("The curve tag is: " + swCurveParaData.CurveTag);
                // Debug.Print("The curve type as defined in swCurveType_e is: " + swCurveParaData.CurveType);

                // Debug.Print("CurveType   = " + swCurve.Identity());
                // Debug.Print("CurveLength = " + (swCurve.GetLength2((double)vCurveParam[6], (double)vCurveParam[7]) * 1000) + " mm");

                // Debug.Print("The curve type as defined in swCurveType_e is: " + swCurveParaData.CurveType);
                // Debug.Print("CurveLength = " + (swCurve.GetLength3((double)vCurveParam[6], (double)vCurveParam[7]) * 1000) + " mm");

                double[] vLineParam = swCurve.LineParams as double[];
                if (vLineParam != null)
                {
                    // Debug.Print("Root point = (" + (vLineParam[0] * 1000) + ", " + (vLineParam[1] * 1000) + ", " + (vLineParam[2] * 1000) + ") mm");
                    // Debug.Print("Direction = (" + vLineParam[3] + ", " + vLineParam[4] + ", " + vLineParam[5] + ")");
                    // Debug.Print("+++++++++++++++++++++++++++");
                }
                else
                {
                    // Debug.Print("Line parameters are not available.");
                }

                if (swCurveParaData.CurveType == 3001)
                {
                    // Debug.Print("The curve tag is: " + swCurveParaData.CurveTag);
                    swModel.ClearSelection2(true);
                    swView.SelectEntity(swEdge, true);

                    for (int j = 0; j <= vEdges.GetUpperBound(0); j++)
                    {
                        if (i == j) continue;

                        Edge edge2 = (Edge)vEdges[j];
                        double[] curveParams2 = (double[])edge2.GetCurveParams2();
                        CurveParamData curveParams3 = (CurveParamData)edge2.GetCurveParams3();

                        // Debug.Print("The curve tag is: " + curveParams3.CurveTag);
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
                                //swModel.AddDimension2(0.2, 0.165, 0);

                                // 计算放置尺寸的位置 - 视图中间并稍微向上
                                double nXoffset = ((double[])curveParams3.EndPoint)[0] - ((double[])curveParams3.StartPoint)[0];
                                double nYoffset = ((double[])curveParams3.EndPoint)[1] - ((double[])curveParams3.StartPoint)[1];
                                double nZoffset = ((double[])curveParams3.EndPoint)[2] - ((double[])curveParams3.StartPoint)[2];

                                //先把nOffset设置为0
                                double nOffset = Math.Sqrt(Math.Pow(nXoffset, 2) + Math.Pow(nYoffset, 2) + Math.Pow(nZoffset, 2))*0;
                                // Debug.Print("nOffset (暂时设置为0)= " + nOffset);

                                double[] vOutline = (double[])swView.GetOutline();
                                double nXpos = (vOutline[0] + vOutline[2]) / 2.0 + nOffset;
                                double nYpos = vOutline[3] + nOffset;

                                // Debug.Print("nXpos = " + nXpos);
                                // Debug.Print("nYpos = " + nYpos);
                                // 创建尺寸，即使实体在图纸视图中不可见
                                swModel.AddDimension2(nXpos, nYpos, 0.0);

                                /*DisplayDimension myDisplayDim;
                                myDisplayDim = (DisplayDimension)swModel.Extension.AddDimension(nXpos, nYpos, 0.0, (int)swSmartDimensionDirection_e.swSmartDimensionDirection_Right);
                                myDisplayDim.ExplementaryAngle();
                                myDisplayDim.VerticallyOppositeAngle();*/

                                // Debug.Print("Selected two parallel and equal length edges and added dimension.");
                                //swModel.ClearSelection2(true);

                                // 重新选中 edge1 继续遍历
                                swView.SelectEntity(swEdge, false);
                                // Debug.Print("The curve tag is: " + swCurveParaData.CurveTag);
                                // Debug.Print("满足平行且相等");

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
                double val1 = dim1.GetDimension2(0).GetValue2("");
                double val2 = dim2.GetDimension2(0).GetValue2("");
                return val2.CompareTo(val1);
            });

            // Count occurrences of each value
            Dictionary<double, int> valueCounts = new Dictionary<double, int>();
            List<double> uniqueValues = new List<double>();
            foreach (var dim in dimensions)
            {
                double value = dim.GetDimension2(0).GetValue2("");
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
            
            // 计算真正的中间值：找到最接近 (maxValue + minValue) / 2 的值
            double targetMidValue = (maxValue + minValue) / 2;
            double midValue = uniqueValues
                .OrderBy(v => Math.Abs(v - targetMidValue))
                .First();

            // Process dimensions based on case
            bool isKeeping;
            int maxValueKept = 0;
            int midValueKept = 0;
            int minValueKept = 0;

            foreach (var dim in dimensions)
            {
                double value = dim.GetDimension2(0).GetValue2("");
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
                    Debug.Print("Deleting dimension: " + annotation.GetName());
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

                        if (angle > 0 && angle < 90 || angle > 90 && angle < 180) // Exclude 90° and 180° angles
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
        }

        public SldWorks swApp;
    }
}
