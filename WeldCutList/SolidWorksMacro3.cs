using SolidWorks.Interop.sldworks;
using System;
using System.Diagnostics;
using WeldCutList;


//SOLIDWORKS API Help
//Get Solid Bodies from Cut-list Folders and Get Custom Properties Example (C#)
//This example shows how to get the solid bodies from cut-list folders and how to get the custom properties for the solid bodies.

//---------------------------------------------------------------
// Preconditions:
// 1. Open public_documents\samples\tutorial\api\weldment_box3.sldprt.
// 2. Click Tools > Options > Document Properties > Weldments >
//    Rename cut list folders with Description property value > OK.
// 3. Right-click Cut list(31) in the FeatureManager design tree
//    and click Update.
// 4. Add a reference to Microsoft.VisualBasic (click Project >
//    Add Reference > Microsoft.VisualBasic).
// 5. Open the Immediate window.
//
// Postconditions:
// 1. Traverses the FeatureManager design tree.
// 2. Examine the Immediate window.
//
// NOTE: Because this part is used elsewhere, do not save changes.
//----------------------------------------------------------------

namespace Macro1CSharp.csproj
{
    public partial class SolidWorksMacro
    {
        ModelDoc2 swPart;
        Feature swFeat;

        public void Main()
        {
            swPart = (ModelDoc2)swApp.ActiveDoc;

            //Debug.Print("File: " + swPart.GetPathName());
            string ConfigName = null;

            try
            {
                ConfigName = swPart.ConfigurationManager.ActiveConfiguration.Name;
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("solidworks当前窗口必须是一个模型!");
                return;
            }
            //Debug.Print("Active configuration name: " + ConfigName);
            swFeat = (Feature)swPart.FirstFeature();
            TraverseFeatures(swFeat, true, "Root Feature");
        }
        public void GetFeatureCustomProps(Feature thisFeat)
        {
            CustomPropertyManager CustomPropMgr = default(CustomPropertyManager);
            CustomPropMgr = (CustomPropertyManager)thisFeat.CustomPropertyManager;
            string[] vCustomPropNames = null;
            vCustomPropNames = (string[])CustomPropMgr.GetNames();
            if ((vCustomPropNames != null))
            {
                Debug.Print("               Cut-list custom properties:");
                int i = 0;
                for (i = 0; i <= (vCustomPropNames.Length - 1); i++)
                {
                    string CustomPropName = null;
                    CustomPropName = (string)vCustomPropNames[i];
                    int CustomPropType = 0;
                    CustomPropType = CustomPropMgr.GetType2(CustomPropName);
                    string CustomPropVal = "";
                    string CustomPropResolvedVal = "";
                    CustomPropMgr.Get2(CustomPropName, out CustomPropVal, out CustomPropResolvedVal);
                    Debug.Print("                     Name: " + CustomPropName);
                    Debug.Print("                         Value: " + CustomPropVal);
                    Debug.Print("                         Resolved value: " + CustomPropResolvedVal);
                }
            }
        }

        bool static_DoTheWork_InBodyFolder;
        readonly Microsoft.VisualBasic.CompilerServices.StaticLocalInitFlag static_DoTheWork_BodyFolderType_Init = new Microsoft.VisualBasic.CompilerServices.StaticLocalInitFlag();
        string[] static_DoTheWork_BodyFolderType;
        bool static_DoTheWork_BeenHere;
        public void DoTheWork(Feature thisFeat, string ParentName)
        {
            lock (static_DoTheWork_BodyFolderType_Init)
            {
                try
                {
                    if (InitStaticVariableHelper(static_DoTheWork_BodyFolderType_Init))
                    {
                        static_DoTheWork_BodyFolderType = new string[6];
                    }
                }
                finally
                {
                    static_DoTheWork_BodyFolderType_Init.State = 1;
                }
            }
            bool bAllFeatures = false;
            bool bCutListCustomProps = false;
            object vSuppressed = null;
            int BodyCount = 0;

            if (!static_DoTheWork_BeenHere)
            {
                static_DoTheWork_BodyFolderType[0] = "dummy";
                static_DoTheWork_BodyFolderType[1] = "swSolidBodyFolder";
                static_DoTheWork_BodyFolderType[2] = "swSurfaceBodyFolder";
                static_DoTheWork_BodyFolderType[3] = "swBodySubFolder";
                static_DoTheWork_BodyFolderType[4] = "swWeldmentSubFolder";
                static_DoTheWork_BodyFolderType[5] = "swWeldmentCutListFolder";
                static_DoTheWork_InBodyFolder = false;
                static_DoTheWork_BeenHere = true;
                bAllFeatures = false;
                bCutListCustomProps = false;
            }

            //Comment out next line to print information for just BodyFolders
            //bAllFeatures = true;
            //True to print information about all features    

            //Comment out next line if you do not want cut list's custom properties
            //bCutListCustomProps = true;

            string FeatType = null;
            FeatType = thisFeat.GetTypeName();
            if ((FeatType == "SolidBodyFolder") & (ParentName == "Root Feature"))
            {
                static_DoTheWork_InBodyFolder = true;
            }
            if ((FeatType != "SolidBodyFolder") & (ParentName == "Root Feature"))
            {
                static_DoTheWork_InBodyFolder = false;
            }

            //Only consider the CutListFolders that are under SolidBodyFolder
            if ((static_DoTheWork_InBodyFolder == false) & (FeatType == "CutListFolder"))
            {
                //Skip the second occurrence of the CutListFolders during the feature traversal
                return;
            }

            //Only consider the SubWeldFolder that are under the SolidBodyFolder
            if ((static_DoTheWork_InBodyFolder == false) & (FeatType == "SubWeldFolder"))
            {
                //Skip the second occurrence of the SubWeldFolders during the feature traversal
                return;
            }

            bool IsBodyFolder = false;
            //if (FeatType == "SolidBodyFolder" | FeatType == "SurfaceBodyFolder" | FeatType == "CutListFolder" | FeatType == "SubWeldFolder" | FeatType == "SubAtomFolder")
            if (FeatType == "CutListFolder")
            {
                IsBodyFolder = true;
            }
            else
            {
                IsBodyFolder = false;
            }
            if (bAllFeatures & (!IsBodyFolder))
            {
                Debug.Print("Feature name: " + thisFeat.Name);
                Debug.Print("   Feature type: " + FeatType);
                //vSuppressed = thisFeat.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);
                //if ((vSuppressed == null))
                //{
                //    Debug.Print("        Suppression failed");
                //}
                //else
                //{
                //    Debug.Print("        Suppressed");
                //}
            }
            if (IsBodyFolder)
            {
                BodyFolder BodyFolder = default(BodyFolder);
                BodyFolder = (BodyFolder)thisFeat.GetSpecificFeature2();
                BodyCount = BodyFolder.GetBodyCount();
                if ((FeatType == "CutListFolder") & (BodyCount < 1))
                {
                    //When BodyCount = 0, this cut list folder is not displayed in the
                    //FeatureManager design tree, so skip it
                    return;
                }
                //else
                //{
                //    Debug.Print("Feature name: " + thisFeat.Name);
                //    Debug.Print("  Feature type: " + FeatType);
                //    vSuppressed = thisFeat.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);
                //    if ((vSuppressed == null))
                //    {
                //        Debug.Print("        Suppression failed");
                //    }
                //    else
                //    {
                //        Debug.Print("        Suppressed");
                //    }
                //}
                if (!bAllFeatures)
                {
                    //Debug.Print("Feature name: " + thisFeat.Name);
                    //Debug.Print("  Feature type: " + FeatType);
                    //vSuppressed = thisFeat.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);
                    //if ((vSuppressed == null))
                    //{
                    //    Debug.Print("        Suppression failed");
                    //}
                    //else
                    //{
                    //    Debug.Print("        Suppressed");
                    //}
                }
                int BodyFolderTypeE = 0;
                BodyFolderTypeE = BodyFolder.Type;
                //Debug.Print("        Body folder: " + static_DoTheWork_BodyFolderType[BodyFolderTypeE]);
                //Debug.Print("        Body folder type: " + BodyFolderTypeE);
                //Debug.Print("        Body count: " + BodyCount);
                object[] vBodies = null;
                vBodies = (object[])BodyFolder.GetBodies();
                int i = 0;

                using (CutListSample01Entities cutListSample01Entities1 = new CutListSample01Entities())
                {
                    if ((vBodies != null))
                    {
                        for (i = 0; i <= (vBodies.Length - 1); i++)
                        {
                            Body2 Body = default(Body2);
                            Body = (Body2)vBodies[i];
                            Debug.Print("Feature name: " + thisFeat.Name);
                            Debug.Print("          Body name: " + Body.Name);

                            CutList cutList = new CutList { Folder_Name = thisFeat.Name, Body_Name = Body.Name, MaterialProperty = thisFeat.GetTypeName2() };
                            cutListSample01Entities1.CutLists.Add(cutList);

                            try
                            {
                                cutListSample01Entities1.SaveChanges();
                            }
                            catch (System.Exception)
                            {
                                return;
                            }
                        }
                    }
                }
            }
            else
            {
                if (bAllFeatures)
                {
                    Debug.Print("");
                }
            }
            if ((FeatType == "CutListFolder"))
            {
                //When BodyCount = 0, this cut-list folder is not displayed
                //in the FeatureManager design tree, so skip it
                if (BodyCount > 0)
                {
                    if (bCutListCustomProps)
                    {
                        //Comment out this line if you do not want to
                        //print the cut-list folder's custom properties
                        GetFeatureCustomProps(thisFeat);
                    }
                }
            }

        }
        public void TraverseFeatures(Feature thisFeat, bool isTopLevel, string ParentName)
        {
            Feature curFeat = default(Feature);
            curFeat = thisFeat;
            while ((curFeat != null))
            {
                DoTheWork(curFeat, ParentName);
                Feature subfeat = default(Feature);
                subfeat = (Feature)curFeat.GetFirstSubFeature();
                while ((subfeat != null))
                {
                    TraverseFeatures(subfeat, false, curFeat.Name);
                    Feature nextSubFeat = default(Feature);
                    nextSubFeat = (Feature)subfeat.GetNextSubFeature();
                    subfeat = (Feature)nextSubFeat;
                    nextSubFeat = null;
                }
                subfeat = null;
                Feature nextFeat = default(Feature);
                if (isTopLevel)
                {
                    nextFeat = (Feature)curFeat.GetNextFeature();
                }
                else
                {
                    nextFeat = null;
                }
                curFeat = (Feature)nextFeat;
                nextFeat = null;
            }
        }

        static bool InitStaticVariableHelper(Microsoft.VisualBasic.CompilerServices.StaticLocalInitFlag flag)
        {
            if (flag.State == 0)
            {
                flag.State = 2;
                return true;
            }
            else if (flag.State == 2)
            {
                throw new Microsoft.VisualBasic.CompilerServices.IncompleteInitialization();
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;
    }
}
