using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

//SOLIDWORKS API Help
//Copy and Paste Drawing Sheet Example (C#)
//This example shows how to copy and paste drawing sheets. 

//----------------------------------------------------------
// Preconditions:
// 1. Open a drawing document containing one sheet
//    named Sheet1.
// 2. Open the Immediate window.
//
// Postconditions:
// 1. Activates Sheet1.
// 2. Copy and pastes Sheet1 as Sheet1(2) and activates Sheet1(2).
// 3. Copy and pastes Sheet1 as Sheet1(3) and activates Sheet1(3).
// 4. Examine the FeatureManager design tree and Immediate window.
//----------------------------------------------------------
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Diagnostics;
 
namespace CopyAndPasteCsharp.csproj
{
    public partial class SolidWorksMacro
    {

        public void Main(int sheetQuantity=1)
        {
            DrawingDoc Part = default(DrawingDoc);
            ModelDoc2 swModel = default(ModelDoc2);
            bool boolstatus = false;

            swModel = (ModelDoc2)swApp.ActiveDoc;
            Part = (DrawingDoc)swModel;
            if ((Part == null))
            {
                Debug.Print(" Please open a drawing document. ");
                return;
            }

            Sheet currentsheet = default(Sheet);
            currentsheet = (Sheet)Part.GetCurrentSheet();
            Part.ActivateSheet(currentsheet.GetName());
            Debug.Print("Active sheet: " + currentsheet.GetName());

            //boolstatus = swModel.Extension.SelectByID2("Sheet1", "SHEET", 0.09205356547875, 0.10872368523, 0, false, 0, null, 0);
            //swModel.EditCopy();
            //boolstatus = Part.PasteSheet((int)swInsertOptions_e.swInsertOption_BeforeSelectedSheet, (int)swRenameOptions_e.swRenameOption_No);
            //currentsheet = (Sheet)Part.GetCurrentSheet();
            //Part.ActivateSheet(currentsheet.GetName());
            //Debug.Print("Active sheet: " + currentsheet.GetName());

            for (int i = 0; i < sheetQuantity; i++)
            {
                boolstatus = swModel.Extension.SelectByID2("Sheet1", "SHEET", 0.09205356547875, 0.10872368523, 0, false, 0, null, 0);
                swModel.EditCopy();
                boolstatus = Part.PasteSheet((int)swInsertOptions_e.swInsertOption_AfterSelectedSheet, (int)swRenameOptions_e.swRenameOption_No);
                currentsheet = (Sheet)Part.GetCurrentSheet();
                Part.ActivateSheet(currentsheet.GetName());
                Debug.Print("Active sheet: " + currentsheet.GetName()); 
            }

        }

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;
    }
}
