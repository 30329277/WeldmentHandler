using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using WeldCutList;
using WeldCutList.ViewModel; // Ensure this using directive is present


namespace Dimensioning02.csproj
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

        // Change the Main method signature
        public void Main(DrawingViewModel drawingViewModel)
        {
            //尝试用方法 createUnfoldViewAt3() 添加project view, 但是无法对齐
            //用了copy paste 的方法, 然后 runcommand
            SolidWorksMacro8 macro8 = new SolidWorksMacro8();
            macro8.CreateAndAlignProjectedViews(drawingViewModel);
        }

        public SldWorks swApp;
    }
}
