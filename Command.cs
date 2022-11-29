#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using Forms = System.Windows.Forms;

#endregion

namespace Week_3
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand

       {
        internal double MakeMetric(double metres)
        {
            double feet = metres * 3.28084;

            return feet;
        }

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //                  ***** TRANSACTION START CODE  ***** 
           

            //                  ***** LEVELS  *****
            //              *****  open levels csv file  *****
            //       TODO add task dialogue for instructions 
            Forms.OpenFileDialog selectLevelfile = new Forms.OpenFileDialog();
            selectLevelfile.InitialDirectory = "R:\\revit\\dynamo\\WIP\\Laura\\C#\\";
            selectLevelfile.Filter = "CSV file|*.csv";

            string levelfilePATH = "";
            if (selectLevelfile.ShowDialog() == Forms.DialogResult.OK)
            {
                levelfilePATH = selectLevelfile.FileName;
            }

            if (levelfilePATH == "")
            {
                return Result.Failed;
            }

            //                    ***** VIEWS  *****
            //              *****  Create views - RCP and Plan  *****
            FilteredElementCollector ViewFilterCollector = new FilteredElementCollector(doc);
            ViewFilterCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType ViewFamTypePlan = null;
            ViewFamilyType ViewFamTypeRCP = null;

            foreach (ViewFamilyType VFT in ViewFilterCollector)
            {
                if (VFT.ViewFamily == ViewFamily.FloorPlan)
                    ViewFamTypePlan = VFT;
                if (VFT.ViewFamily == ViewFamily.CeilingPlan)
                    ViewFamTypeRCP = VFT;
            }

            //          *****  read data from excel - Levels  *****
            string[] fileArrayLvl = System.IO.File.ReadAllLines(levelfilePATH);

            Transaction trans = new Transaction(doc);
            trans.Start("Create level and sheet");
            //             *****  remove header row using split to separate text file data  *****
            foreach (string rowst in fileArrayLvl.Skip(1))
            {
                string[] cellst = rowst.Split(',');
                string LevelNames = cellst[0];
                int LevelElevationInt = Convert.ToInt32(cellst[2]);
                double lvlHeight = MakeMetric(LevelElevationInt);
                Level myLvl = Level.Create(doc, lvlHeight);
                myLvl.Name = LevelNames;
                ViewPlan planviews = ViewPlan.Create(doc, ViewFamTypePlan.Id, myLvl.Id);
                ViewPlan RCPviews = ViewPlan.Create(doc, ViewFamTypeRCP.Id, myLvl.Id);

            }
            //                  *****   END LEVELS AND VIEWS   *****       

            //                  *****  SHEETS  *****


            //           ***** open sheets csv file  *****
            Forms.OpenFileDialog selectsheetfile = new Forms.OpenFileDialog();
            selectsheetfile.InitialDirectory = "R:\\revit\\dynamo\\WIP\\Laura\\C#\\";
            selectsheetfile.Filter = "CSV file|*.csv";

            string sheetfilePATH = "";

            if (selectsheetfile.ShowDialog() == Forms.DialogResult.OK)
            {
                sheetfilePATH = selectsheetfile.FileName;
            }

            if (sheetfilePATH == "")
            {
                // return Result.Failed;
            }

            //          *****  read data from excel - sheets  *****

            string[] fileArraySh = System.IO.File.ReadAllLines(sheetfilePATH);

            Element TBlock = TitleBlockByName(doc, "A1 Landscape");

            //              *****  remove header row using split to separate text file data  *****
            foreach (string rowst in fileArraySh.Skip(1))
            {
                string[] cellst = rowst.Split(',');
                string SheetNumber = cellst[0];
                string SheetName = cellst[1];
                ViewSheet thesheets = ViewSheet.Create(doc, TBlock.Id);
                thesheets.SheetNumber = SheetNumber;
             
            }

            // Modify Revit document within transaction 

            trans.Commit();
            trans.Dispose();

            return Result.Succeeded;
        }

        //          *****  Get stuff from Revit API for sheets *****
        internal Element TitleBlockByName(Document doc, string TBlockName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                        
            foreach (Element TitleBlock in collector)
            {
                if (TitleBlock.Name == TBlockName)
                {
                    return TitleBlock;
                }
            }
            return null;
        }      

    }
}
