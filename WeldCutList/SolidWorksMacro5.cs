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

namespace CenterOfMass_CSharp.csproj
{
    partial class SolidWorksMacro
    {

        ModelDoc2 swModel;
        DrawingDoc swDrawDoc;
        View swView;
        CenterOfMass swCenterOfMass;
        Annotation swAnnotation;
        int sheetCount;
        int viewCount;
        object[] ss;
        object[] vv;
        double[] coord;

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
            var queryArrBodyWithIndex = arrBody.Select((item, index) => new {index, Body = (Body2)item });

            using (CutListSample01Entities cutListSample01Entities1 = new CutListSample01Entities())
            {
                var queryDB =
                from product in cutListSample01Entities1.CutLists
                select new { product.Folder_Name, product.Body_Name , product.MaterialProperty};

                var resultquery =
                from c in queryArrBodyWithIndex
                join p in queryDB on c.Body.Name equals p.Body_Name
                select new { Index = c.index, BodyName = c.Body.Name, FolderName = p.Folder_Name };

                var arrayFromResultQuert=resultquery.ToArray();

                Console.WriteLine(queryDB.Count()+" "+ arrBody.Count()+" "+ queryArrBodyWithIndex.Count());
                
                for (sheetCount = ss.GetLowerBound(0); sheetCount <= ss.GetUpperBound(0); sheetCount++)
                {
                    vv = (object[])ss[sheetCount];

                    for (viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                    {
                        Debug.Print(((View)vv[viewCount]).GetName2());
                        swView = (View)vv[viewCount];
                        try
                        {
                            Bodies[0] = arrBody[arrayFromResultQuert[sheetCount* 16 + viewCount - 1].Index];
                        }
                        catch (Exception)
                        {
                            //Console.WriteLine(arrayFromResultQuert.Count());
                            //Console.WriteLine(((16 * (ss.GetUpperBound(0)+1))- arrayFromResultQuert.Count()).ToString());
                            System.Windows.Forms.MessageBox.Show(((16 * (ss.GetUpperBound(0) + 1)) - arrayFromResultQuert.Count()).ToString() +"个view无效");
                            return;
                        }
                        arrBodiesIn[0] = new DispatchWrapper(Bodies[0]);
                        swView.Bodies = (arrBodiesIn);
                    }
                }
            }
        }

        public SldWorks swApp;

    }
}


