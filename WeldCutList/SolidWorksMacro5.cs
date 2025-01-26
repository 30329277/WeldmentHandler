using GalaSoft.MvvmLight;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
//SOLIDWORKS API Help
//Get Centers of Mass in Drawing Views Example (C#)
//This example shows how to get the centers of mass in drawing views.

//----------------------------------------------------------------------
// Preconditions: Open a drawing that has one or more centers of mass.
//
// Postconditions: Inspect the Immediate window for the coordinates of
// all of the centers of mass in all of the views of the
// drawing.
//----------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using WeldCutList;
using WeldCutList.ViewModel;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace CenterOfMass_CSharp.csproj
{
    /// <summary>
    /// 和命名空间 CenterOfMass_CSharp无关, 用于遍历工程图上所有的 view 并获取新的集合,气球, 摆正等功能
    /// </summary>
    public partial class SolidWorksMacro : ViewModelBase
    {

        View swView;
        Sheet swSheet;
        private string swViewName;
        private string swSheetName;

        DrawingDoc swDrawDoc;
        ModelDoc2 swModel;
        CenterOfMass swCenterOfMass;
        Annotation swAnnotation;
        int sheetCount;
        int viewCount;
        object[] ss;
        object[] vv;
        double[] coord;
        //<<<<<<<<<<<<<<
        DrawingDoc swDraw;
        DocumentSpecification swDocSpecification;
        DrawingComponent swDrawingComponent;
        Component2 swComponent;
        Entity swEntity;
        object[] vEdges;
        bool bRet;
        int i;

        object[] Bodies = new object[1];
        DispatchWrapper[] arrBodiesIn = new DispatchWrapper[1];

        /// <summary>
        /// 获取新的集合,气球, 摆正等功能
        /// </summary>
        /// <param name="drawingViewModel">DrawingViewModel类型的drawingViewModel</param>
        public void Main(DrawingViewModel drawingViewModel)
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            //try catch 只是判断当前打开的是否是工程图
            try
            {
                swDrawDoc = (DrawingDoc)swModel;
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("solidworks当前窗口必须是一个工程图!");
                return;
            }
            ss = (object[])swDrawDoc.GetViews();
            var tempvv = (object[])ss[0];
            var tempView = (View)tempvv[1];
            var swModelTemp01 = tempView.ReferencedDocument;
            var swPart = (PartDoc)swModelTemp01;
            //获取当前工程图的reference part 的所有body 的集合
            var arrBody = (object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);

            //使用了 select 函数的重载方法 获取集合中元素 + index, 投影为新的集合 queryArrBodyWithIndex
            //一般情况使用的 select c=>c , 这里的重载用了两个参数 select((item,index)=>new{})
            var queryArrBodyWithIndex = arrBody.Select((item, index) => new { index, Body = (Body2)item });

            //用这个方法拿到工程图中的 weld cut list bom , var list (集合用于为 body 和 localDB 进行排序)
            AnnotationCounts_CSharp.csproj.SolidWorksMacro macro = new AnnotationCounts_CSharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
            var list = macro.Main();

            #region autoballoons的参数

            AutoBalloonOptions autoballoonParams;
            autoballoonParams = swDrawDoc.CreateAutoBalloonOptions();

            #endregion

            // Replace Excel COM reading code with EPPlus
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cutlist.xlsx");
            if (!File.Exists(excelPath))
            {
                System.Windows.Forms.MessageBox.Show($"Excel file not found: {excelPath}");
                return;
            }

            var cutListData = new List<CutList>();
            
            // Remove license context code since we're using EPPlus 4.5.3.3
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var workbook = package.Workbook;
                if (workbook.Worksheets.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("No worksheets found in the Excel file.");
                    return;
                }

                var worksheet = workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.Rows;

                // Skip header row, start from row 2
                for (int row = 2; row <= rowCount; row++)
                {
                    cutListData.Add(new CutList
                    {
                        Folder_Name = worksheet.Cells[row, 1].Value?.ToString(),
                        Body_Name = worksheet.Cells[row, 2].Value?.ToString(),
                        MaterialProperty = worksheet.Cells[row, 3].Value?.ToString()
                    });
                }
            }

            // Project Excel data to queryDB
            var queryDB = cutListData.Select(product => new { product.Folder_Name, product.Body_Name, product.MaterialProperty });

            //零件设计树的body集合和 queryDB（来自与localDB）join 后投影为新的集合 resultquery
            //生成的新集合包括：三列 index bodyname foldername , 集合元素的数量 和 localDB的元素数量一致
            var resultquery =
            from c in queryArrBodyWithIndex
            join p in queryDB on c.Body.Name equals p.Body_Name
            select new { Index = c.index, BodyName = c.Body.Name, FolderName = p.Folder_Name };

            var arrayFromResultQuery1 = resultquery.ToArray();

            //注意这个 拉姆达表达式用于把 arrayFromResultQuery1 按照工程图中的 weld cut list bom 进行排序  (a.foldername , 是排序列)
            var arrayFromResultQuery = arrayFromResultQuery1.OrderBy(a => Array.IndexOf(list.ToArray(), a.FolderName)).ToArray();

            //这里是打印 queryDB,arrBody,queryArrBodyWithIndex 集合中元素的数量
            Console.WriteLine(queryDB.Count() + " " + arrBody.Count() + " " + queryArrBodyWithIndex.Count());
            Console.WriteLine($"当前工程图的reference part 的所有body 的集合数量是:{arrBody.Count()}");
            Console.WriteLine($"{queryDB.FirstOrDefault().ToString()},queryDB集合的数量: {queryDB.Count()}");
            Console.WriteLine($"{queryArrBodyWithIndex.FirstOrDefault().index},{queryArrBodyWithIndex.FirstOrDefault().Body},queryArrBodyWithIndex的数量是:  {queryArrBodyWithIndex.Count()}");
            Console.WriteLine($"list(来自BOM)的数量是:{list.Count}");
            Console.WriteLine($"resultquery的数量是：{resultquery.Count()}");
            Console.WriteLine($"arrayFromResultQuery的数量是:{arrayFromResultQuery.Count()}");

            // Constants for A0 sheet dimensions and view spacing
            const double sheetWidth = 1.189; // A0 width in meters
            const double sheetHeight = 0.841; // A0 height in meters
            const int viewsPerRow = 5;
            const int maxRows = 6;
            const double viewSpacingX = sheetWidth / (viewsPerRow + 1);
            const double viewSpacingY = sheetHeight / (maxRows + 1);

            for (sheetCount = ss.GetLowerBound(0); sheetCount <= ss.GetUpperBound(0); sheetCount++)
            {
                vv = (object[])ss[sheetCount];
                //这个循环实际测试没有效果
                for (viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                {
                    swView = (View)vv[viewCount];
                    swView.AlignWithView(0, swView);
                }

                //这个循环有效果
                for (viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                {
                    //Debug.Print(((View)vv[viewCount]).GetName2());
                    swView = (View)vv[viewCount];
                    swSheetName = swView.Sheet.GetName();
                    SwViewName = (string)swView.Name;
                    //Debug.Print(swSheetName);
                    //Debug.Print(SwViewName);

                    //试试outline在view 重置 body 以后是否发生了变化
                    double[] outline = (double[])swView.GetOutline();
                    double[] pos = (double[])swView.Position;
                /*    Debug.Print("  X and Y positions = (" + pos[0] * 1000.0 + ", " + pos[1] * 1000.0 + ") mm");
                    Debug.Print("  X and Y bounding box minimums = (" + outline[0] * 1000.0 + ", " + outline[1] * 1000.0 + ") mm");
                    Debug.Print("  X and Y bounding box maximums = (" + outline[2] * 1000.0 + ", " + outline[3] * 1000.0 + ") mm");
                    Debug.Print("  bounding box size = (" + (outline[2] - outline[0]) * 1000.0 + ", " + (outline[3] - outline[1]) * 1000.0 + ") mm");*/

                    try
                    {
                        Bodies[0] = arrBody[arrayFromResultQuery[sheetCount * 30 + viewCount - 1].Index];
                    }
                    catch (Exception)
                    {
                        System.Windows.Forms.MessageBox.Show(((30 * (ss.GetUpperBound(0) + 1)) - arrayFromResultQuery.Count()).ToString() + "个view无效");
                        return;
                    }
                    //这两行代码真正的替换 body 
                    arrBodiesIn[0] = new DispatchWrapper(Bodies[0]);
                    swView.Bodies = (arrBodiesIn);

                    #region 添加气球 并重置位置

                    //增加气球功能
                    swModel.Extension.SelectByID2(((View)vv[viewCount]).GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swDrawDoc.AutoBalloon5(autoballoonParams);

                    // 检查视图是否已经有气球
                    bool hasBalloons = false;
                    var annotations = (object[])swView.GetAnnotations();
                    if (annotations != null)
                    {
                        foreach (var annotation in annotations)
                        {
                            var swAnnotation = (Annotation)annotation;
                            if (swAnnotation.GetType() == (int)swAnnotationType_e.swNote) // Changed from swBalloonNote to swNote

                            {
                                hasBalloons = true;
                                break;
                            }
                        }
                    }

                    // 如果没有气球，则添加气球
                    if (!hasBalloons)
                    {
                        // Configure balloon options
                        BalloonOptions bomBalloonParams = swModel.Extension.CreateBalloonOptions();
                        bomBalloonParams.Style = (int)swBalloonStyle_e.swBS_Circular;
                        bomBalloonParams.Size = (int)swBalloonFit_e.swBF_2Chars;
                        bomBalloonParams.UpperTextContent = (int)swBalloonTextContent_e.swBalloonTextItemNumber;
                        bomBalloonParams.ShowQuantity = true;
                        bomBalloonParams.QuantityPlacement = (int)swBalloonQuantityPlacement_e.swBalloonQuantityPlacement_Right;
                        bomBalloonParams.QuantityDenotationText = "PLACES";
                        bomBalloonParams.ItemNumberStart = 1;
                        bomBalloonParams.ItemNumberIncrement = 1;
                        bomBalloonParams.ItemOrder = (int)swBalloonItemNumbersOrder_e.swBalloonItemNumbers_DoNotChangeItemNumbers;

                        // Select any edge in the view
                        var edges = (object[])swView.GetVisibleEntities2(swComponent, (int)swViewEntityType_e.swViewEntityType_Edge);
                        if (edges != null && edges.Length > 0)
                        {
                            var swEdge = (Entity)edges[0];
                            swEdge.Select4(true, null);

                            // Insert balloon with configured options
                            var swNote = (Note)swModel.Extension.InsertBOMBalloon2(bomBalloonParams);

                            swModel.ClearSelection2(true);
                        }
                    }
                    #endregion

                    //视图breakalignment 和 position 功能
                    swModel.ClearSelection2(true);

                    //已经把 break alignment 提前到了前边, 所以这一行注释掉
                    //swView.AlignWithView(0, swView);

                    //根据长边调整一下方向
                    AlignViewWithTheLongestEdge(swModel, swView.Name);
                    //把 editrebuild 注释掉 非常慢
                    //swModel.EditRebuild3();
                    //注释掉 forcerebuild 非常慢
                    //swModel.ForceRebuild3(true);

                    // Calculate the position for the view
                    int row = (viewCount - 1) / viewsPerRow;
                    int col = (viewCount - 1) % viewsPerRow;
                    double[] vPos = { (col + 1) * viewSpacingX, sheetHeight - (row + 1) * viewSpacingY };
                    swView.Position = vPos;

                    // After setting the position, get the view's outline and center it vertically
                    //double[] finalOutline = (double[])swView.GetOutline();
                    //double viewHeight = finalOutline[3] - finalOutline[1];
                    //// Adjust Y position to center the view based on its height
                    //vPos[1] += viewHeight / 2;
                    //swView.Position = vPos;

                    #region 新增逻辑,用于纠正视图的位置

                    // Get the annotations in the view
                    var viewAnnotations = (object[])swView.GetAnnotations();
                    if (viewAnnotations != null && viewAnnotations.Length > 0)
                    {
                        for (int i = 0; i < viewAnnotations.Length; i++)
                        {
                            swAnnotation = (Annotation)viewAnnotations[i];
                            if (swAnnotation.GetType() == (int)swAnnotationType_e.swNote)
                            {
                                double[] annotationPos = (double[])swAnnotation.GetPosition();
                                if (annotationPos != null && annotationPos.Length > 0)
                                {
                                    double ax = annotationPos[0];
                                    double ay = annotationPos[1];
                                    double[] viewPos = (double[])swView.Position;
                                    double viewX = viewPos[0];
                                    double viewY = viewPos[1];
                                    double offsetX = ax - viewX;
                                    double offsetY = ay - viewY;
                                    // Check if offsets are greater than 0.01
                                    if (Math.Abs(offsetX) >= 0.01 || Math.Abs(offsetY) >= 0.01)
                                    {
                                        double[] newPosition = { viewX - offsetX, viewY - offsetY, swView.ScaleDecimal };
                                        swView.SetXform(newPosition);
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    // Optionally, force update
                    //swModel.EditRebuild3();
                    Console.WriteLine($"第{viewCount}个视图的坐标:x={vPos[0]};y={vPos[1]}");


                    //试试outline在view 重置 body 以后是否发生了变化
                    var outline2 = (double[])swView.GetOutline();
                    var pos2 = (double[])swView.Position;
                   /* Debug.Print("  X and Y positions = (" + pos2[0] * 1000.0 + ", " + pos2[1] * 1000.0 + ") mm");
                    Debug.Print("  X and Y bounding box minimums = (" + outline2[0] * 1000.0 + ", " + outline2[1] * 1000.0 + ") mm");
                    Debug.Print("  X and Y bounding box maximums = (" + outline2[2] * 1000.0 + ", " + outline2[3] * 1000.0 + ") mm");
                    Debug.Print("  bounding box size = (" + (outline2[2] - outline2[0]) * 1000.0 + ", " + (outline2[3] - outline2[1]) * 1000.0 + ") mm");*/

                    //把 editrebuild 注释掉 非常慢
                    //swModel.EditRebuild3();

                    #region 想反偏置 效果不好

                    //判断一下新的view和旧的view比,向哪个象限发生了offset
                    //double xBoundingBox = outline2[2] - outline[2];
                    //double yBoundingBox = outline2[3] - outline[3];

                    //if (xBoundingBox > 0 && yBoundingBox > 0)
                    //{
                    //    vPos[0] -= ((outline[2] - outline[0]) - (outline2[2] - outline2[0])) / 2;
                    //    vPos[1] -= ((outline[3] - outline[1]) - (outline2[3] - outline2[1])) / 2;
                    //}
                    //else if (xBoundingBox > 0 && yBoundingBox < 0)
                    //{
                    //    vPos[0] -= ((outline[2] - outline[0]) - (outline2[2] - outline2[0])) / 2;
                    //    vPos[1] += ((outline[3] - outline[1]) - (outline2[3] - outline2[1])) / 2;
                    //}
                    //else if (xBoundingBox < 0 && yBoundingBox > 0)
                    //{
                    //    vPos[0] += ((outline[2] - outline[0]) - (outline2[2] - outline2[0])) / 2;
                    //    vPos[1] -= ((outline[3] - outline[1]) - (outline2[3] - outline2[1])) / 2;
                    //}
                    //else if (xBoundingBox < 0 && yBoundingBox < 0)
                    //{
                    //    vPos[0] += ((outline[2] - outline[0]) - (outline2[2] - outline2[0])) / 2;
                    //    vPos[1] += ((outline[3] - outline[1]) - (outline2[3] - outline2[1])) / 2;
                    //}

                    //swView.Position = vPos;
                    //swModel.EditRebuild3();

                    #endregion

                    //注释掉, 比较费时
                    //AlignViewWithTheLongestEdge(swModel, swView.Name);

                    drawingViewModel.SheetName = SwSheetName;
                    drawingViewModel.ViewName = SwViewName;
                }
            }
        }

        /// <summary>
        /// 这个方法用于调整 view 的方向, 找到最长的边, 然后align
        /// </summary>
        /// <param name="swModel"></param>
        /// <param name="viewName"></param>
        public void AlignViewWithTheLongestEdge(ModelDoc2 swModel, string viewName)
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDraw = (DrawingDoc)swModel;

            // Get the current sheet
            //swSheet = (Sheet)swDraw.GetCurrentSheet();

            //想激活当前的sheet, 失败了
            //activate the current sheet 
            //swModel.SetPickMode();
            //var boolActivated= swDraw.ActivateSheet(swSheet.GetName());
            //swDraw.GetCurrentSheet();
            //Console.WriteLine(boolActivated);

            //SwSheetName = (string)swSheet.GetName();
            //Debug.Print(swSheet.GetName());
            //Application.Current.Dispatcher.Invoke(new Action(() => {}));

            // Select Drawing View1
            bRet = swModel.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0.0, 0.0, 0.0, true, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swView = (View)((SelectionMgr)swModel.SelectionManager).GetSelectedObject6(1, -1);

            // Print the drawing view name and get the component in the drawing view
            //SwViewName = (string)swView.Name;
            //Debug.Print(swView.GetName2());
            swDrawingComponent = swView.RootDrawingComponent;
            swComponent = swDrawingComponent.Component;

            // Get the component's visible entities in the drawing view
            int eCount = 0;
            eCount = swView.GetVisibleEntityCount2(swComponent, (int)swViewEntityType_e.swViewEntityType_Edge);
            vEdges = (object[])swView.GetVisibleEntities2(swComponent, (int)swViewEntityType_e.swViewEntityType_Edge);
            //Debug.Print("Number of edges found: " + eCount);

            Dictionary<int, double> e = new Dictionary<int, double>();


            // Hide all of the visible edges in the drawing view
            for (i = 0; i <= eCount - 1; i++)
            {
                swEntity = (Entity)vEdges[i];
                swEntity.Select4(true, null);
                //swDraw.HideEdge();

                ModelDocExtension swModelDocExt = swModel.Extension;
                Measure swMeasure = swModelDocExt.CreateMeasure();

                SelectionMgr swSelMgr = swModel.ISelectionManager;

                swMeasure.ArcOption = 0;

                var status = swMeasure.Calculate(null);
                if ((status))
                {
                    if ((!(swMeasure.Length == -1)))
                    {
                        //Debug.Print("Length: " + swMeasure.Length);
                        e.Add(i, swMeasure.Length);
                    }
                }
                else
                {
                    Debug.Print("Invalid combination of selected entities.");
                }

                swModel.ClearSelection2(true);
            }

            //Console.WriteLine(e.Count);
            var maxKey = e.OrderByDescending(x => x.Value).First().Key;
            swEntity = (Entity)vEdges[maxKey];
            swEntity.Select4(true, null);

            swDraw.AlignHorz();

            //swModel.ForceRebuild3(true);
            swModel.EditRebuild3();

            // Clear all selections
            swModel.ClearSelection2(true);

            // Show all hidden edges
            //swView.HiddenEdges = vEdges;

        }

        public SldWorks swApp;

        public string SwViewName
        {
            get { return swViewName; }
            set
            {
                swViewName = value; RaisePropertyChanged();
            }
        }

        public string SwSheetName
        {
            get { return swSheetName; }
            set
            {
                swSheetName = value; RaisePropertyChanged();
            }
        }
    }
}


