/*
Solves the problem of room area calculation by moving the boundary line 
from the middle of the curtain grid to the edges of the element based on its thickness. 
*/
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace SpaceSeparator2021
{
    [Transaction(TransactionMode.Manual)]
    public class SpaceBoundary  : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            BuiltInCategory spaceSeparator = BuiltInCategory.OST_MEPSpaceSeparationLines;
            BuiltInCategory roomSeparator = BuiltInCategory.OST_RoomSeparationLines;

            FilteredElementCollector separatorCollector = new FilteredElementCollector(doc, doc.ActiveView.Id);

            try
            {
                TaskDialog dialog = new TaskDialog("Choose separator type");
                dialog.MainContent = "Select YES for SPACE and NO for ROOM";
                dialog.AllowCancellation = true;
                dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

                TaskDialogResult result = dialog.Show();

                if (result == TaskDialogResult.Yes)
                {
                    separatorCollector.OfCategory(spaceSeparator);
                }
                else
                {
                    separatorCollector.OfCategory(roomSeparator);
                }

                FilteredElementCollector wallCollector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                wallCollector.OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();

                ICollection<ElementId> elementIds = separatorCollector.ToElementIds();

                CurveArray crAr = new CurveArray();

                foreach (Wall wall in wallCollector)
                {
                    XYZ endpt, startpt;
                    XYZ z = XYZ.BasisZ;
                    double thickness = 0.0;
                    LocationCurve curve = wall.Location as LocationCurve;

                    if (curve == null)
                    {
                        message = "There are no walls in project";
                    }

                    string name = wall.Name;
                    string[] numbers = Regex.Split(name, @"\D+");

                    foreach (string value in numbers)
                    {
                        if (!string.IsNullOrEmpty(value))
                        {
                            switch (value)
                            {
                                case "42":
                                    thickness = 0.0689;
                                    break;

                                case "62":
                                    thickness = 0.1017;
                                    break;

                                case "82":
                                    thickness = 0.13452;
                                    break;
                            }
                        }
                    }

                    startpt = curve.Curve.GetEndPoint(0);
                    endpt = curve.Curve.GetEndPoint(1);
                    for (int i = 1; i != -3; i = i - 2)
                    {
                        Curve crOffset = curve.Curve.CreateOffset((i * thickness), z);
                        crAr.Append(crOffset);
                    }
                }

                Transaction trans = new Transaction(doc);
                trans.Start("Space");

                foreach (ElementId elid in elementIds)
                {
                    doc.Delete(elid);
                }

                if (result == TaskDialogResult.Yes)
                {
                    doc.Create.NewSpaceBoundaryLines(doc.ActiveView.SketchPlane, crAr, doc.ActiveView);

                }
                else
                {
                    doc.Create.NewRoomBoundaryLines(doc.ActiveView.SketchPlane, crAr, doc.ActiveView);
                }

                trans.Commit();

                return Result.Succeeded;
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
