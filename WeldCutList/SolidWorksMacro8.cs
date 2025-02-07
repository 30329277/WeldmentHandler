using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using SolidWorks.Interop.swcommands;
using WeldCutList.ViewModel; // Add this using directive

namespace WeldCutList
{
    public partial class SolidWorksMacro8
    {
        public void CreateAndAlignProjectedViews(DrawingViewModel drawingViewModel) // Add DrawingViewModel parameter
        {
            int errors = 0;
            int warnings = 0;

            SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            DrawingDoc swDrawDoc = (DrawingDoc)swModel;
            ModelDocExtension swModelDocExt = swModel.Extension;
            object[] ss = (object[])swDrawDoc.GetViews();

            for (int sheetCount = ss.GetLowerBound(0); sheetCount <= ss.GetUpperBound(0); sheetCount++)
            {
                string[] sheetNames = (string[])swDrawDoc.GetSheetNames();
                swDrawDoc.ActivateSheet(sheetNames[sheetCount]); // Activate the current sheet
                
                // Update sheet name in view model
                drawingViewModel.SheetName = sheetNames[sheetCount];
                
                object[] vv = (object[])ss[sheetCount];
                for (int viewCount = 1; viewCount <= vv.GetUpperBound(0); viewCount++)
                {
                    View swView = (View)vv[viewCount];
                    string originalViewName = swView.GetName2();
                    
                    // Update view name in view model
                    drawingViewModel.ViewName = originalViewName;

                    // Skip if view is derived or already projected
                    if (swView.GetBaseView() != null || originalViewName.Contains("Projected"))
                        continue;

                    // Check view type
                    int viewType = swView.Type;
                    if (/*viewType == (int)swDrawingViewTypes_e.swDrawingAlternatePositionView ||
                        viewType == (int)swDrawingViewTypes_e.swDrawingAuxiliaryView ||
                        viewType == (int)swDrawingViewTypes_e.swDrawingDetachedView ||
                        viewType == (int)swDrawingViewTypes_e.swDrawingDetailView ||
                        viewType == (int)swDrawingViewTypes_e.swDrawingNamedView ||*/
                        viewType == (int)swDrawingViewTypes_e.swDrawingProjectedView /*||
                        viewType == (int)swDrawingViewTypes_e.swDrawingRelativeView ||
                        viewType == (int)swDrawingViewTypes_e.swDrawingSectionView ||
                        viewType == (int)swDrawingViewTypes_e.swDrawingSheet ||
                        viewType == (int)swDrawingViewTypes_e.swDrawingStandardView*/)
                    {
                        continue;
                    }

                    // Select and copy-paste the view
                    bool status = swModelDocExt.SelectByID2(originalViewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.EditCopy();
                    swModel.Paste();

                    // Get the newly pasted view
                    View pastedView = null;
                    object[] updatedViews = (object[])swDrawDoc.GetViews();
                    object[] currentSheetViews = (object[])updatedViews[sheetCount];

                    // Find the pasted view (it will be the last view in the list)
                    for (int i = currentSheetViews.GetUpperBound(0); i >= 1; i--)
                    {
                        View tempView = (View)currentSheetViews[i];
                        if (tempView.GetName2() != originalViewName)
                        {
                            pastedView = tempView;
                            break;
                        }
                    }

                    if (pastedView != null)
                    {
                        // Change orientation based on original view
                        string originalOrientation = swView.GetOrientationName().ToLower();
                        swModelDocExt.SelectByID2(pastedView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                        switch (originalOrientation)
                        {
                            case "*front":
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_ProjectedView, "28");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_3DDrawingView, "2988");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Left, "164");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Save, "");
                                status = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);
                                break;
                            case "*top":
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_ProjectedView, "28");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_3DDrawingView, "2988");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Right, "165");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Save, "");
                                status = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);
                                break;
                            case "*bottom":
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_ProjectedView, "28");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_3DDrawingView, "2988");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Left, "164");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Save, "");
                                status = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);
                                break;
                            case "*left":
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_ProjectedView, "28");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_3DDrawingView, "2988");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Left, "164");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Save, "");
                                status = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);
                                break;
                            default:
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_ProjectedView, "28");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_3DDrawingView, "2988");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Left, "164");
                                swModelDocExt.RunCommand((int)swCommands_e.swCommands_Save, "");
                                status = swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);
                                break;
                        }


                        // Check aspect ratio and adjust orientation if needed
                        double[] pastedViewOutline = (double[])pastedView.GetOutline();
                        double pastedViewWidth = Math.Abs(pastedViewOutline[2] - pastedViewOutline[0]);
                        double pastedViewHeight = Math.Abs(pastedViewOutline[3] - pastedViewOutline[1]);


                        double[] viewOutline = (double[])swView.GetOutline();
                        double viewWidth = Math.Abs((double)viewOutline[2] - (double)viewOutline[0]);

                        if (pastedViewWidth < pastedViewHeight)
                        {
                            // Rotate view to make the width greater than the height
                            //swModelDocExt.SelectByID2(pastedView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                            //pastedView.Angle = 90.0;
                            pastedView.Angle = Math.PI / 2; // Set angle to 90 degrees in radians

                        }

                        // Change alignment direction - make pasted view align with original view
                        status = pastedView.AlignWithView((int)swAlignViewTypes_e.swAlignViewHorizontalCenter, swView);

                        // Set up AutoBalloon parameters
                        AutoBalloonOptions autoballoonParams = swDrawDoc.CreateAutoBalloonOptions();
                        autoballoonParams.Layout = (int)swBalloonLayoutType_e.swDetailingBalloonLayout_Square;
                        /*autoballoonParams.ReverseDirection = false;
                        autoballoonParams.IgnoreMultiple = true;
                        autoballoonParams.InsertMagneticLine = true;
                        autoballoonParams.LeaderAttachmentToFaces = false;
                        autoballoonParams.Style = (int)swBalloonStyle_e.swBS_Circular;
                        autoballoonParams.Size = (int)swBalloonFit_e.swBF_5Chars;
                        autoballoonParams.UpperTextContent = (int)swBalloonTextContent_e.swBalloonTextItemNumber;
                        autoballoonParams.Layername = "-None-";
                        autoballoonParams.ItemNumberStart = 1;
                        autoballoonParams.ItemNumberIncrement = 1;
                        autoballoonParams.ItemOrder = (int)swBalloonItemNumbersOrder_e.swBalloonItemNumbers_DoNotChangeItemNumbers;
                        autoballoonParams.EditBalloons = true;
                        autoballoonParams.EditBalloonOption = (int)swEditBalloonOption_e.swEditBalloonOption_Resequence;*/

                        // Select the pasted view and add balloons
                        swModel.Extension.SelectByID2(pastedView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                        swDrawDoc.AutoBalloon5(autoballoonParams);

                        // Optional: Adjust position if needed

                        double[] originalPos = (double[])swView.Position;
                        double[] pastedPos = (double[])pastedView.Position;
                        pastedView.Position = new double[] { originalPos[0] + viewWidth + 0.1, originalPos[1], 0 }; // Align X and Y positions with the original view




                        // Determine alignment based on polyline count and edge type
                        int nCountShaded;
                        int polyLineCount = swView.GetPolyLineCount5(1, out nCountShaded);
                        // Replace the line that causes the error
                        // List<object> vEdges = new List<object>((object[])swView.GetEdges());

                        // With the following line
                        List<object> vEdges = new List<object>((object[])swView.GetVisibleEntities(null, (int)swSelectType_e.swSelEDGES));

                        if (polyLineCount == 16 || (polyLineCount == 17 && vEdges.Count(e => e is Edge && ((Edge)e).GetCurveParams3().CurveType == 3002) == 8))
                        {

                            // Align pasted view to the left of the original view
                            pastedView.Position = new double[] { originalPos[0] - viewWidth - 0.1, originalPos[1], 0 };
                        }
                        else
                        {
                            // Align pasted view to the right of the original view
                            pastedView.Position = new double[] { originalPos[0] + viewWidth + 0.1, originalPos[1], 0 };
                        }
                    }
                }
            }
        }
    }
}
