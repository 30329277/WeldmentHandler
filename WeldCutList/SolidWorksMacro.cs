using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.Runtime.InteropServices;
//SOLIDWORKS API Help
//Set Body for View Example (C#)
//This example shows how to show just one body of a multibody part in a drawing view. 
//------------------------------------------------------------------
// Preconditions:
// 1. Open public_documents\samples\tutorial\multibody\multi_inter.sldprt. 
// 2. Save the part document as a drawing document: 
//    a. Click File > Make Drawing from Part. 
//    b. Click OK on the Sheet Format/Size dialog. 
//    c. Drag the *Isometric view from the View Palette onto 
//       the drawing sheet. 
// 3. Select the drawing view.
// 4. Open the Immediate window.
// 
// Postconditions: 
// 1. Shows one body of the multibody part 
//    in the drawing view. 
// 2. Examine the drawing and the Immediate window.
//
// NOTE: Because the part document is used elsewhere, do not save 
// changes. 
//------------------------------------------------------------------ 
using System.Windows.Forms;
namespace BodiesViewCSharp.csproj
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            ModelDoc2 swModel = default(ModelDoc2);
            SelectionMgr swSelMgr = default(SelectionMgr);
            SolidWorks.Interop.sldworks.View swView = default(SolidWorks.Interop.sldworks.View);
            int nbrBodies = 0;
            object[] arrBody = null;
            Body2 swBody = default(Body2);
            Face2 swFace = default(Face2);
            Entity swEnt = default(Entity);
            SelectData swSelData = default(SelectData);
            PartDoc swPart = default(PartDoc);
            bool status = false;
            DispatchWrapper[] arrBodiesIn = new DispatchWrapper[1];
            object[] Bodies = new object[1];
            int i = 0;
            int objType = 0;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swView = (SolidWorks.Interop.sldworks.View)swSelMgr.GetSelectedObject6(1, -1);
            if ((swView == null))
            {
                MessageBox.Show("View not selected.");
                return;
            }
            nbrBodies = swView.GetBodiesCount();
            Debug.Print("Number of bodies: " + nbrBodies);
            if ((nbrBodies < 1))
            {
                MessageBox.Show("No bodies in selected view.");
                return;
            }
            arrBody = (object[])swView.Bodies;

            //for (i = 0; i < arrBody.Length; i++)
            //{
            //    swBody = (Body2)arrBody[i];
            //    swSelData = (SelectData)swSelMgr.CreateSelectData();
            //    swSelData.View = swView;
            //    status = swBody.Select2(false, swSelData);
            //    // Object type 76 is a solid body 
            //    objType = swSelMgr.GetSelectedObjectType3(1, -1);
            //    if ((objType == 76))
            //    {
            //        Debug.Print(" Object type: solid body");
            //    }
            //    if ((!((int)swSelectType_e.swSelSOLIDBODIES == swSelMgr.GetSelectedObjectType3(1, -1))))
            //    {
            //        MessageBox.Show("Solid body not found.");
            //    }
            //    swFace = (Face2)swBody.GetFirstFace();
            //    while ((swFace != null))
            //    {
            //        swEnt = (Entity)swFace;
            //        // Select using IEntity 
            //        status = swEnt.Select4(true, swSelData);
            //        Debug.Assert(status);
            //        swFace = (Face2)swFace.GetNextFace();
            //    }
            //    Debug.Print(" Name of body: " + swBody.GetSelectionId());
            //}

            swModel.ClearSelection2(true);
            // Get the bodies from referenced model 
            swModel = (ModelDoc2)swView.ReferencedDocument;
            swPart = (PartDoc)swModel;
            arrBody = (object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if ((nbrBodies == 1))
            {
                swView.Bodies = (arrBody);
            }
            else
            {
                // Set the body to include in the drawing view 
                Bodies[0] = arrBody[3];
                arrBodiesIn[0] = new DispatchWrapper(Bodies[0]);
                swView.Bodies = (arrBodiesIn);
            }
            swModel.ClearSelection2(true);
        }
        /// <summary> 
        /// The SldWorks swApp variable is pre-assigned for you. 
        /// </summary> 
        public SldWorks swApp;
    }
}
