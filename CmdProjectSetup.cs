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

#endregion

namespace SessionTwoChallenge
{
    [Transaction(TransactionMode.Manual)]
    public class CmdProjectSetup : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)

          {

                      
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;



            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            Transaction trans = new Transaction(doc);
            trans.Start("Create level and sheet");

            //read data from excel - Levels
            string filepathLvl = @"R:\revit\dynamo\WIP\\Laura\C#\ArchSmWk2\RAB_Session_02_Challenge_Levels.csv";
            string[ ] fileArrayLvl = System.IO.File.ReadAllLines(filepathLvl);
            //string filetxt = System.IO.File.ReadAllText(filepathLvl);

            //read data from excel - Sheets
            string filepathSheets = @"R:\revit\dynamo\WIP\\Laura\C#\ArchSmWk2\RAB_Session_02_Challenge_Sheets.csv";
            string[] fileArraySh = System.IO.File.ReadAllLines(filepathSheets);
            //string filetxtSh = System.IO.File.ReadAllText(filepathSheets);
            
            //remove header row using split to separate text file data
            foreach (string rowst in fileArrayLvl.Skip(1))
            {
                string[] cellst = rowst.Split(',');
                string LevelName = cellst[0];
                //string LevelElevation = cellst[2];
                int LevelElevationInt = Convert.ToInt32(cellst[2]);
                //double elevationdbl = 0;
                //bool didItParse =  double.TryParse(LevelElevation, out elevationdbl);
                double lvlHeight = Metres2Feet(LevelElevationInt);
                Level myLvl = Level.Create(doc, lvlHeight);
                myLvl.Name = LevelName;

            }

            //remove header row using split to separate text file data
            foreach (string rowst in fileArraySh.Skip(1))
            {
                string[] cellst = rowst.Split(',');
                string SheetNumber = cellst[0];
                string SheetName = cellst[1];
                ViewSheet thesheets = ViewSheet.Create(doc, collector.FirstElementId());
                thesheets.SheetNumber = SheetNumber;
                thesheets.Name = SheetName;
            }

            trans.Commit();
            trans.Dispose();


            return Result.Succeeded;

        }

        internal double Metres2Feet(double meters)
        {
            double feet = meters * 3.28084;

            return feet;
        }
    }
}
