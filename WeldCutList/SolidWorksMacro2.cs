//-------------------------------------------------------------
// Preconditions:
// 1. Verify that the specified drawing document to open exists.
// 2. Open the Immediate window.
//
// Postconditions:
// 1. Opens the specified drawing.
// 2. Horizontally aligns Drawing View6 with
//    the center of Drawing View3.
// 3. Examine the drawing and Immediate window.
//
// NOTE: Because the drawing is used elsewhere, do not
// save changes.
//-------------------------------------------------------------
//这个方法暂时没用
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;


namespace AlignWithViewCSharp.csproj
{
    public partial class SolidWorksMacro
    {
        ModelDoc2 swModel;
        DrawingDoc swDrawingDoc;
        View swView;
        View swViewAlignWith;
        ModelDocExtension swModelDocExtension;
        SelectionMgr swSelectionMgr;
        bool status;
        int errors;
        int warnings;
        string fileName;


        public void Main()
        {
            // Open drawing
            fileName = "C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2018\\samples\\tutorial\\advdrawings\\foodprocessor.slddrw";
            swModel = (ModelDoc2)swApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swDrawingDoc = (DrawingDoc)swModel;

            // Activate Sheet2
            status = swDrawingDoc.ActivateSheet("Sheet2");
            swModel.ClearSelection2(true);

            // Activate Drawing View6 and make it the active
            // drawing view
            status = swDrawingDoc.ActivateView("Drawing View6");
            swView = (View)swDrawingDoc.ActiveDrawingView;
            swModelDocExtension = (ModelDocExtension)swModel.Extension;

            // Select Drawing View3 and select it to be
            // the base view
            status = swModelDocExtension.SelectByID2("Drawing View3", "DRAWINGVIEW", 0.166130896593432, 0.265618781363279, 0, false, 0, null, 0);
            swSelectionMgr = (SelectionMgr)swModel.SelectionManager;
            swViewAlignWith = (View)swSelectionMgr.GetSelectedObject6(1, -1);

            // Drawing View6 aligns horizontally
            // with the center of Drawing View3
            status = swView.AlignWithView((int)swAlignViewTypes_e.swAlignViewHorizontalCenter, swViewAlignWith);

            Debug.Print("Did the specified drawing views align? " + status);


        }

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;
    }
}
