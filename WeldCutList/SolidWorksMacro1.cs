using SolidWorks.Interop.sldworks;
using System.Collections;
using System.Diagnostics;

//SOLIDWORKS API Help
//Create Unfolded View Example (C#)
//This example shows how to create an unfolded view from an existing view.

//----------------------------------------------------------------------------
// Preconditions: Open:
//    public_documents\samples\tutorial\advdrawings\foodprocessor.slddrw
//
// Postconditions: A new unfolded view is created from Drawing View1.
//
// NOTE: Because the model is used elsewhere,
// do not save changes when closing it.
// ---------------------------------------------------------------------------
namespace InsertUnfoldedView_CSharp.csproj
{
    /// <summary>
    /// 用于复制多个视图
    /// </summary>
    partial class SolidWorksMacro
    {
        public void Main()
        {
            DrawingDoc swDraw = default(DrawingDoc);
            View swView = default(View);
            ModelDoc2 Part;
            double[] outline = null;
            double[] pos = null;
            string fileName = null;
            int errors = 0;
            int warnings = 0;

            Part = (ModelDoc2)swApp.ActiveDoc;
            try
            {
                swDraw = (DrawingDoc)swApp.ActiveDoc;
            }
            catch (System.Exception)
            {
                System.Windows.Forms.MessageBox.Show("solidworks当前窗口必须是一个工程图!");
                return;
            }
            swView = (View)swDraw.GetFirstView();
            ArrayList arrayList = new ArrayList();
            while ((swView != null))
            {
                outline = (double[])swView.GetOutline();
                pos = (double[])swView.Position;
                arrayList.Add(swView.Name);
                arrayList.Add(pos[0]);
                arrayList.Add(pos[1]);

                //尝试把视图放到 0，0 位置
                pos[0] = 0;
                pos[1] = 0;
                swView.Position = pos;
                Debug.Print("View = " + swView.Name);
                Debug.Print("  X and Y positions = (" + pos[0] * 1000.0 + ", " + pos[1] * 1000.0 + ") mm");
                Debug.Print("  X and Y bounding box minimums = (" + outline[0] * 1000.0 + ", " + outline[1] * 1000.0 + ") mm");
                Debug.Print("  X and Y bounding box maximums = (" + outline[2] * 1000.0 + ", " + outline[3] * 1000.0 + ") mm");
                Debug.Print("  Position locked?" + swView.PositionLocked);
                swView = (View)swView.GetNextView();
            }
            if (arrayList.Count <= 3)
            {
                System.Windows.Forms.MessageBox.Show("当前工程图需要至少存在一个view!");
                return;
            }

            #region 新的 copy and paste 方案

            Part.Extension.SelectByID2(arrayList[3].ToString(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            for (int i = 0; i < 29; i++)
            {
                //用 editcut() 会报错
                //Part.EditCut();
                Part.EditCopy();
                Part.Paste();
            }

            #endregion



            #region 原来使用CreateUnfoldedViewAt3() 创建 viwe 的部分, 先注释掉

            //double xSpace = 0.19;
            //double ySpace = 0.13;
            //ArrayList verticalViewNames = new ArrayList();
            //for (int i = 0; i < 4; i++)
            //{
            //    Part.Extension.SelectByID2(arrayList[3].ToString(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            //    swDraw.CreateUnfoldedViewAt3((double)arrayList[4] + 0.025 + xSpace * (i + 1), (double)arrayList[5], 0, false);
            //}
            //for (int i = 0; i < 5; i++)
            //{
            //    Part.Extension.SelectByID2(arrayList[3].ToString(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            //    string tempVerticalName = swDraw.CreateUnfoldedViewAt3((double)arrayList[4], (double)arrayList[5] - 0.025 - ySpace * (i + 1), 0, false).Name;
            //    verticalViewNames.Add(tempVerticalName);
            //}
            //for (int i = 0; i < verticalViewNames.Count; i++)
            //{
            //    object item = verticalViewNames[i];
            //    if (item != null)
            //    {
            //        for (int j = 0; j < 4; j++)
            //        {
            //            Part.Extension.SelectByID2(item.ToString(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            //            swDraw.CreateUnfoldedViewAt3((double)arrayList[4] + 0.025 + xSpace * (j + 1), (double)arrayList[5] - 0.025 - ySpace * (i + 1), 0, false);
            //        }
            //    }
            //}

            #endregion

        }

        public SldWorks swApp;

    }
}

