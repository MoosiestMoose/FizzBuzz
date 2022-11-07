#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace FizzBuzz
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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

            XYZ myPoint = new XYZ(10, 10, 0);
            XYZ myNextPoint = new XYZ();

            // 5. Filtered Element Collectors
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //collector.OfCategory(BuiltInCategory.OST_TextNotes);
            //collector.WhereElementIsElementType();
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc);
            t.Start("CreateTextNote");

            XYZ offset = new XYZ(0, -5, 0);
            XYZ newPoint = myPoint;

            for (int i = 1; i <= 100; i++)
            {
                newPoint = newPoint.Add(offset);

                string result = "";


                        if (i % 3 == 0)
                        {
                            if (i % 5 == 0)
                                result = "fizzbuzz";
                            else
                                result = "buzz";
                        }
                        else if (i % 5 == 0)
                            result = "fizz";
                        else
                            result = i.ToString();

                TextNote myTextNote = TextNote.Create(doc, doc.ActiveView.Id,
                newPoint, result + "\n",
                collector.FirstElementId());

            }
            t.Commit();
    
            return Result.Succeeded;
        }
    }
}

