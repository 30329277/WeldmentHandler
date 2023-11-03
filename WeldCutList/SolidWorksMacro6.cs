using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

//SOLIDWORKS API Help
//Hide and Show All Edges in Drawing View (C#)
//This example shows how to hide and then show all of the edges in the root component in a drawing view.

//---------------------------------------------------------------------------
// Preconditions: 
// 1. Verify that the specified drawing document to open exists.
// 2. Open the Immediate window.
//
// Postconditions: 
// 1. Opens the specified drawing document.
// 2. Hides and then shows all edges in the root component in 
//    Drawing View1.
// 3. Examine the drawing and Immediate window.
//----------------------------------------------------------------------------
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Data.Entity.Infrastructure;
using System.Linq;

namespace HideShowEdges_CSharp.csproj
{
    partial class SolidWorksMacro
    {

        ModelDoc2 swModel;
        DrawingDoc swDraw;
        DocumentSpecification swDocSpecification;
        Sheet swSheet;
        View swView;
        DrawingComponent swDrawingComponent;
        Component2 swComponent;
        Entity swEntity;
        object[] vEdges;
        bool bRet;
        int i;



        public void Main()
        {
            // Specify the drawing to open
            //swDocSpecification = (DocumentSpecification)swApp.GetOpenDocSpec
            //    (@"C:\Users\107895\Desktop\temp\temp swd\foodprocessor.slddrw");
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (swModel == null)
            {
                swModel = swApp.OpenDoc7(swDocSpecification);
            }

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDraw = (DrawingDoc)swModel;

            // Get the current sheet
            swSheet = (Sheet)swDraw.GetCurrentSheet();
            Debug.Print(swSheet.GetName());

            // Select Drawing View1
            bRet = swModel.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.0, 0.0, 0.0, true, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swView = (View)((SelectionMgr)swModel.SelectionManager).GetSelectedObject6(1, -1);

            // Print the drawing view name and get the component in the drawing view
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
                        Debug.Print("Length: " + swMeasure.Length);
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

            swModel.ForceRebuild3(true);

            //swModel.Rebuild(8);

            // Clear all selections
            swModel.ClearSelection2(true);

            // Show all hidden edges
            swView.HiddenEdges = vEdges;

        }

        public SldWorks swApp;

    }
}


