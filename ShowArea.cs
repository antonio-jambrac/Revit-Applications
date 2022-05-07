using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ViewArea
{
    [Transaction(TransactionMode.Manual)]
    public class ShowArea : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Selection sel = uidoc.Selection;
            try
            {
                ICollection<ElementId> elementIds = sel.GetElementIds();
                List<double> areaList = new List<double>();

                foreach (ElementId id in elementIds)
                {
                    Element el = doc.GetElement(id);
                    string elAreaString = (el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsValueString());
                    string elAreastr = (Regex.Match(elAreaString, @"\d+\,\d{2}").Value);
                    double elArea = double.Parse(elAreastr);

                    if (elArea == 0)
                    {
                        double width = double.Parse(el.get_Parameter(BuiltInParameter.GENERIC_WIDTH).AsValueString()) / 1000;// u metrima
                        double height = double.Parse(el.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsValueString()) / 1000;
                        double areaCalc = width * height;

                        areaList.Add(areaCalc);
                    }
                    else
                    {
                        areaList.Add(elArea);
                    }
                }

                double result = 0;
                Transaction trans = new Transaction(doc);

                StringBuilder sb = new StringBuilder();
                trans.Start("Show Area of selected elemets");

                foreach (double n in areaList)
                {
                    result += n;
                }

                sb.Append(result.ToString());
                TaskDialog.Show("Selected area is :", sb.ToString() + " m2");
                trans.Commit();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
