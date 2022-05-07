/*
Disallow or allow join on both ends of a wall element. 
The app can be used on all elements on same or just on the selected element. 
*/
using System;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace disallow_join
{
    [Transaction(TransactionMode.Manual)]
    public class DisallowAllowJoin : IExternalCommand
    {

        public void DisallowOrAllowJoin (IList<Element> elem)
        {
            List<Wall> controlWall = new List<Wall>();

            foreach (Wall w in elem)
            {
                if (WallUtils.IsWallJoinAllowedAtEnd(w, 0) || WallUtils.IsWallJoinAllowedAtEnd(w, 1))
                {
                    controlWall.Add(w);
                }
            }

            if (controlWall.Count != 0)
            {
                foreach (Wall wall in controlWall)
                {
                    WallUtils.DisallowWallJoinAtEnd(wall, 0);
                    WallUtils.DisallowWallJoinAtEnd(wall, 1);
                }
            }
            else
            {
                foreach (Wall wl in elem)
                {
                    WallUtils.AllowWallJoinAtEnd(wl, 0);
                    WallUtils.AllowWallJoinAtEnd(wl, 1);
                }
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            try
            {
                // get Element ids from selected walls
                Selection sel = uidoc.Selection;
                ICollection<ElementId> elId = sel.GetElementIds();
                IList<Element> elem = new List<Element>();

                Transaction trans = new Transaction(doc);
                trans.Start("Disallow Join");

                if (elId.Count != 0) // disallow/allow join on selected elements
                {
                    foreach (ElementId id in elId) // get element from selected element id
                    {
                        elem.Add(doc.GetElement(id));
                    }

                    DisallowOrAllowJoin(elem);
                }

                else // disallow/allow join all all elements
                {
                    TaskDialog dialog = new TaskDialog("Are you sure");
                    dialog.MainContent = "Do you want to disallow/allow join on all elements?";
                    dialog.AllowCancellation = true;
                    dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

                    TaskDialogResult result = dialog.Show();

                    if (result == TaskDialogResult.Yes)
                    {
                        // collect all walls in project 
                        IList<Element> collector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();
                        DisallowOrAllowJoin(collector);
                    }

                    else
                    {
                        trans.RollBack();
                        return Result.Cancelled;
                    }
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
