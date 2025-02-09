﻿using SolidWorks.Interop.sldworks;
//SOLIDWORKS API Help
//Get the Total Number of Columns and Rows in a Table Example (C#)
//This example shows how to get the total number of columns and rows (visible and hidden) in a table annotation.

//--------------------------------------------------------------------------
// Preconditions: Open a drawing that contains a hole table
// with multiple columns and rows.
//
// Postconditions: Inspect the Immediate Window.
//--------------------------------------------------------------------------
using System.Collections.Generic;
using System.Diagnostics;

namespace AnnotationCounts_CSharp.csproj
{
    partial class SolidWorksMacro
    {
        public List<string> Main()
        {
            List<string> lines = new List<string>();
            ModelDoc2 swModel = default(ModelDoc2);
            DrawingDoc swDraw = default(DrawingDoc);
            View swView = default(View);
            TableAnnotation swTable = default(TableAnnotation);

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swDraw = (DrawingDoc)swModel;

            Debug.Print("File = " + swModel.GetPathName());

            // Get the first view
            swView = (View)swDraw.GetFirstView();

            while ((swView != null))
            {
                // Show the name of the view
                Debug.Print("View =   " + swView.Name);

                // Get the first table annotation for this view
                swTable = swView.GetFirstTableAnnotation();

                while ((swTable != null))
                {
                    //ProcessTable(swApp, swModel, swTable);
                    Annotation swAnn = default(Annotation);
                    int nNumCol = 0;
                    int nNumRow = 0;
                    int nTotalNumCol = 0;
                    int nTotalNumRow = 0;
                    swAnn = (Annotation)swTable.GetAnnotation();

                    // Show the name and type of table
                    Debug.Print("    " + swAnn.GetName() + " <" + swTable.Type + ">");

                    // Get the visible counts
                    nNumCol = swTable.ColumnCount;
                    Debug.Print("      Number of visible columns: " + nNumCol);
                    nNumRow = swTable.RowCount;
                    Debug.Print("      Number of visible rows: " + nNumRow);

                    // Get the total (hidden + visible) counts
                    nTotalNumCol = swTable.TotalColumnCount;
                    Debug.Print("      Total number of visible + hidden columns: " + nTotalNumCol);
                    nTotalNumRow = swTable.TotalRowCount;
                    Debug.Print("      Total number of visible + hidden rows: " + nTotalNumRow);

                    for (int i = 1; i < nTotalNumRow; i++)
                    {
                        lines.Add(swTable.DisplayedText2[i, 1, true]);
                    }

                    // Get next table annotation for this view
                    swTable = swTable.GetNext();
                }

                // Get the next view
                swView = (View)swView.GetNextView();

            }

            return lines;

        }

        public SldWorks swApp;

    }
}
