using SolidWorks.Interop.sldworks;
using System.Diagnostics;

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
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System.Data.Entity.Infrastructure;
using System.Linq;
using WeldCutList;
using System.Windows.Shapes;
using System.Windows.Forms.VisualStyles;
using System.Windows;

namespace CenterOfMass_CSharp.csproj
{
    public partial class SolidWorksMacro
    {

        ModelDoc2 swModel;
        DrawingDoc swDrawDoc;
        View swView;
        public string swViewName;
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
        Sheet swSheet;
        public string swSheetName;
        DrawingComponent swDrawingComponent;
        Component2 swComponent;
        Entity swEntity;
        object[] vEdges;
        bool bRet;
        int i;

        object[] Bodies = new object[1];
        DispatchWrapper[] arrBodiesIn = new DispatchWrapper[1];

        public void Main()
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
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
            var arrBody = (object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);
            var queryArrBodyWithIndex = arrBody.Select((item, index) => new { index, Body = (Body2)item });

            #region autoballoons的参数

            AutoBalloonOptions autoballoonParams;
            autoballoonParams = swDrawDoc.CreateAutoBalloonOptions();

            #endregion


            using (CutListSample01Entities cutListSample01Entities1 = new CutListSample01Entities())
            {
                var queryDB =
                from product in cutListSample01Entities1.CutLists
                select new { product.Folder_Name, product.Body_Name, product.MaterialProperty };

                var resultquery =
                from c in queryArrBodyWithIndex
                join p in queryDB on c.Body.Name equals p.Body_Name
                select new { Index = c.index, BodyName = c.Body.Name, FolderName = p.Folder_Name };

                var arrayFromResultQuert = resultquery.ToArray();

                Console.WriteLine(queryDB.Count() + " " + arrBody.Count() + " " + queryArrBodyWithIndex.Count());

                for (sheetCount = ss.GetLowerBound(0); sheetCount <= ss.GetUpperBound(0); sheetCount++)
                {
                    vv = (object[])ss[sheetCount];

                    for (viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                    {
                        swView = (View)vv[viewCount];

                        swView.AlignWithView(0, swView);
                    }

                    for (viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                    {
                        Debug.Print(((View)vv[viewCount]).GetName2());
                        swView = (View)vv[viewCount];
                        try
                        {
                            Bodies[0] = arrBody[arrayFromResultQuert[sheetCount * 30 + viewCount - 1].Index];
                        }
                        catch (Exception)
                        {
                            System.Windows.Forms.MessageBox.Show(((30 * (ss.GetUpperBound(0) + 1)) - arrayFromResultQuert.Count()).ToString() + "个view无效");
                            return;
                        }
                        arrBodiesIn[0] = new DispatchWrapper(Bodies[0]);
                        swView.Bodies = (arrBodiesIn);

                        #region add balloons

                        //TODO: 增加气球功能
                        swModel.Extension.SelectByID2(((View)vv[viewCount]).GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                        swDrawDoc.AutoBalloon5(autoballoonParams);
                        //TODO: 视图breakalignment 和 position 功能
                        swModel.ClearSelection2(true);

                        //已经把 break alignment 提前到了前边, 所以这一行注释掉
                        //swView.AlignWithView(0, swView);

                        //根据长边调整一下方向
                        AlignViewWithTheLongestEdge(swModel, swView.Name);

                        swModel.EditRebuild3();
                        swModel.ForceRebuild3(true);

                        //193 136
                        double[] vPos = { 0, 0 };
                        vPos[0] = ((viewCount - 1) % 5) * 0.193 + 0.096;
                        vPos[1] = 0.82 - ((viewCount - 1) / 5) * 0.136 - 0.068;
                        Console.WriteLine($"第{viewCount}个视图的坐标:x={vPos[0]};y={vPos[1]}");
                        swView.Position = vPos;

                        swModel.EditRebuild3();
                        //AlignViewWithTheLongestEdge(swModel, swView.Name);
                        #endregion

                    }
                }
            }
        }

        public void AlignViewWithTheLongestEdge(ModelDoc2 swModel, string viewName)
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDraw = (DrawingDoc)swModel;

            // Get the current sheet
            swSheet = (Sheet)swDraw.GetCurrentSheet();
            //activate the current sheet 
            //swModel.SetPickMode();
            //var boolActivated= swDraw.ActivateSheet(swSheet.GetName());
            //swDraw.GetCurrentSheet();
            //Console.WriteLine(boolActivated);

            swSheetName = (string)swSheet.GetName();
            Debug.Print(swSheet.GetName());
            //Application.Current.Dispatcher.Invoke(new Action(() => {}));

            // Select Drawing View1
            bRet = swModel.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0.0, 0.0, 0.0, true, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swView = (View)((SelectionMgr)swModel.SelectionManager).GetSelectedObject6(1, -1);

            // Print the drawing view name and get the component in the drawing view
            swViewName = (string)swView.Name;
            Debug.Print(swView.GetName2());
            swDrawingComponent = swView.RootDrawingComponent;
            swComponent = swDrawingComponent.Component;

            // Get the component's visible entities in the drawing view
            int eCount = 0;
            eCount = swView.GetVisibleEntityCount2(swComponent, (int)swViewEntityType_e.swViewEntityType_Edge);
            vEdges = (object[])swView.GetVisibleEntities2(swComponent, (int)swViewEntityType_e.swViewEntityType_Edge);
            Debug.Print("Number of edges found: " + eCount);

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

            Console.WriteLine(e.Count);
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

    }
}


