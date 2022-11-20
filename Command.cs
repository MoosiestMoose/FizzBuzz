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


            //          *****  Get stuff from Revit API  *****
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            //                  ***** TRANSACTION START CODE  ***** 
            Transaction trans = new Transaction(doc);
            trans.Start("Create level and sheet");

            //              *****  open levels csv file  *****
      
            Forms.OpenFileDialog selectLevelfile = new Forms.OpenFileDialog();
            selectLevelfile.InitialDirectory = "R:\\revit\\dynamo\\WIP\\Laura\\C#\\";
            selectLevelfile.Filter = "CSV file|*.csv";

            string levelfilePATH = "";

            if (selectLevelfile.ShowDialog() == Forms.DialogResult.OK)
            {
                levelfilePATH = selectLevelfile.FileName;
            }

            if (levelfilePATH != "")
            {
                // do something with file
            }

            //          *****  read data from excel - Levels  *****
            string[] fileArrayLvl = System.IO.File.ReadAllLines(levelfilePATH);


            //              *****  remove header row using split to separate text file data  *****
            foreach (string rowst in fileArrayLvl.Skip(1))
            {
                string[] cellst = rowst.Split(',');
                string LevelNames = cellst[0];
                //string LevelElevation = cellst[2];
                int LevelElevationInt = Convert.ToInt32(cellst[2]);
                //double elevationdbl = 0;
                double lvlHeight = MakeMetric(LevelElevationInt);
                Level myLvl = Level.Create(doc, lvlHeight);
                myLvl.Name = LevelNames;

            }


            //           ***** open sheets csv file  *****
            Forms.OpenFileDialog selectsheetfile = new Forms.OpenFileDialog();
            selectsheetfile.InitialDirectory = "R:\\revit\\dynamo\\WIP\\Laura\\C#\\";
            selectsheetfile.Filter = "CSV file|*.csv";

            string sheetfilePATH = "";

            if (selectsheetfile.ShowDialog() == Forms.DialogResult.OK)
            {
                sheetfilePATH = selectsheetfile.FileName;
            }

            if (sheetfilePATH != "")
            {
                // do something with file
            }

            //          *****  read data from excel - sheets  *****
            
            string[] fileArraySh = System.IO.File.ReadAllLines(sheetfilePATH);

            //              *****  remove header row using split to separate text file data  *****
            foreach (string rowst in fileArraySh.Skip(1))
            {
                string[] cellst = rowst.Split(',');
                string SheetNumber = cellst[0];
                string SheetName = cellst[1];
                ViewSheet thesheets = ViewSheet.Create(doc, collector.FirstElementId());
                thesheets.SheetNumber = SheetNumber;
                thesheets.Name = SheetName;
            }

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

            //Transaction Tran = new Transaction(doc);
            //Tran.Start("Create my things");

            Level LevelName = Level.Create(doc, 20);

            ViewPlan planviews = ViewPlan.Create(doc, ViewFamTypePlan.Id, LevelName.Id);
            ViewPlan RCPviews = ViewPlan.Create(doc, ViewFamTypeRCP.Id, LevelName.Id);




            // Modify Revit document within transaction 


            trans.Commit();
            trans.Dispose();



            return Result.Succeeded;
        }

       

    }
}