using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Mark4
{
    public class MultiDictionary<TKey, TValue> : Dictionary<TKey, List<TValue>>
    {
        public void Add(TKey key, TValue value)
        {
            if (TryGetValue(key, out List<TValue> valueList))
            {
                valueList.Add(value);
            }
            else
            {
                Add(key, new List<TValue> { value });
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Mark4Number : IExternalCommand
    {

        public void SetMark (List<int> markValue, List<List<ElementId>> elementIds, Document doc)
        {
            for (int x = 0; x < markValue.Count; x++)
            {
                foreach(ElementId elID in elementIds[x])
                {
                    Element el = doc.GetElement(elID);
                    el.LookupParameter("Mark 4").Set(markValue[x]);           
                }
            } 
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            FilteredElementCollector panelCollector = new FilteredElementCollector(doc);
            panelCollector.OfCategory(BuiltInCategory.OST_CurtainWallPanels).WhereElementIsNotElementType().ToElements();

            //FilteredElementCollector panelId_col = new FilteredElementCollector(doc);
            //panelId_col.OfCategory(BuiltInCategory.OST_CurtainWallPanels).WhereElementIsNotElementType().ToElementIds();
            var errorElementId = new List<ElementId>();
            Selection sel = uidoc.Selection;
            List<ElementId> systemPanelIds = new List<ElementId>();
            MultiDictionary<string, ElementId> sortPanels = new MultiDictionary<string, ElementId>();
            List<Element> inputPanelList = new List<Element>();
            try
            {
                
                if (sel.GetElementIds().Count != 0) // Use selected elements
                {
                    foreach(ElementId eid in sel.GetElementIds())
                    {
                        inputPanelList.Add(doc.GetElement(eid));
                    }
                }
                else // use all elements
                {
                    TaskDialog dialog = new TaskDialog("Are you sure");
                    dialog.MainContent = "Do you want to generate Mark 4 on all elements?";
                    dialog.AllowCancellation = true;
                    dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                    TaskDialogResult result = dialog.Show();

                    if(result == TaskDialogResult.Yes)
                    {
                        foreach (Element eid in panelCollector)
                        {
                            inputPanelList.Add(eid);
                        }
                    }
                    else
                    {
                        return Result.Cancelled;
                    }
                }
                
                // Get Parameters
                foreach (Element el in inputPanelList)
                {
                    ElementId panelId = el.Id;
                    // Check for System panels
                    if (el.LookupParameter("Mark 4 double") == null)
                    {
                        //List of System Panel Id
                        systemPanelIds.Add(el.Id);
                        continue;
                    }
                    double mark4_double = double.Parse(el.LookupParameter("Mark 4 double").AsValueString());
                    sortPanels.Add(mark4_double.ToString(), panelId);
                }

                // Get Keys and Values from MultiDictionary
                var markKeyParam = new List<string>(sortPanels.Keys);
                var panelIdList = new List<List<ElementId>>(sortPanels.Values);

                // Csv file
                string revitProjectName = doc.Title;
                string originalProjectPath = doc.PathName;
                List<string> splitPath = new List<string>(originalProjectPath.Split(new String[] { "\\" }, StringSplitOptions.None));
                string pathForCsv = "";
                Dictionary<string, string> csvDictionary = new Dictionary<string, string>();

                // Create file path name
                for (int i = 0; i < (splitPath.Count - 1); i++)
                {
                    if (i == splitPath.Count - 2)
                    {
                        pathForCsv += splitPath[i] + "\\" + revitProjectName + "Mark4.txt";
                        break;
                    }
                    pathForCsv += splitPath[i] + "\\";
                }
                // Read csv text file
                if(File.Exists(pathForCsv))
                {
                    string readCsv = File.ReadAllText(pathForCsv);
                    if (readCsv != "")
                    {
                        List<string> csvTextList = new List<string>(readCsv.Split(new String[] { ";", "\r\n"}, StringSplitOptions.None));

                        if(csvTextList[csvTextList.Count - 1] == "")
                        {
                            csvTextList.RemoveAt(csvTextList.Count - 1);
                        }

                        if(csvTextList.Count%2 != 0)
                        {
                            message = "List of key, value for Mark 4 are not in pair, please contact developer team";
                            return Result.Failed;
                        }

                        for (int i = 0; i < csvTextList.Count; i += 2)
                        {
                            csvDictionary.Add(csvTextList[i], csvTextList[i + 1]);
                        }
                    }
                }

                List<int> mark4Numbers = new List<int>();
                Random ran = new Random();
                // Generate Mark 4 from CSV file or add random number
                List<string> csvMarkListString = new List<string>(csvDictionary.Values); //get mark 4 from csv as a string
                List<int> csvMarkList = csvMarkListString.Select(n => int.Parse(n)).ToList(); // convert to int

                for(int i = 0; i < markKeyParam.Count; i++)
                {
                    errorElementId.Add(panelIdList[i][0]);
                    string mark4valueFromCsv = "";
                    if (csvDictionary.TryGetValue(markKeyParam[i], out mark4valueFromCsv))
                    {
                        mark4Numbers.Add(int.Parse(mark4valueFromCsv));
                        continue;
                    }
                    else
                    {
                        double[] panel_w_h_t = { double.Parse(doc.GetElement(panelIdList[i][0]).get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH).AsValueString()),
                            double.Parse(doc.GetElement(panelIdList[i][0]).get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT).AsValueString()),
                            double.Parse(doc.GetElement(panelIdList[i][0]).LookupParameter("Thickness").AsValueString()) };

                        panel_w_h_t = panel_w_h_t.Select(el => Math.Round(el, MidpointRounding.AwayFromZero)).ToArray();

                        if(double.Parse(markKeyParam[i]) == panel_w_h_t.Sum()) // Only Standard panels
                        {
                            string strMark4 = "";

                            if ((panel_w_h_t[0] / 100 - (int)(panel_w_h_t[0] / 100)) == 0 & (panel_w_h_t[1] / 100 - (int)(panel_w_h_t[1] / 100) == 0))
                            {

                                if(panel_w_h_t[1] > 999)
                                {
                                    strMark4 = (panel_w_h_t[0]/100).ToString() + (panel_w_h_t[1] / 100).ToString();
                                }
                                else
                                {
                                    strMark4 = (panel_w_h_t[0]/100).ToString() + "0" + (panel_w_h_t[1] / 100).ToString();
                                }

                                // Add to Mark 4 list
                                mark4Numbers.Add(int.Parse(strMark4));
                            }
                            else
                            {
                                int keyParam = int.Parse(markKeyParam[i]); // Mark 4 KEY Value from csv file
                                while(keyParam.ToString().Length > 4)
                                {
                                    keyParam = keyParam / 10;
                                }
                                if (!mark4Numbers.Contains(keyParam) & !csvMarkList.Contains(keyParam))
                                {
                                    mark4Numbers.Add(keyParam);
                                }
                                else
                                {
                                    bool control = true;
                                    while (control)
                                    {
                                        int randomNumber = ran.Next(0, 9999);
                                        if (!mark4Numbers.Contains(randomNumber) & !csvMarkList.Contains(randomNumber))
                                        {
                                            mark4Numbers.Add(randomNumber);
                                        }
                                        control = false;
                                    }
                                }
                            }
                        }
                        else
                        {
                            int keyParam = int.Parse(markKeyParam[i]); // Mark 4 KEY Value from csv file
                            while (keyParam.ToString().Length > 4)
                            {
                                keyParam = keyParam / 10;
                            }
                            if (!mark4Numbers.Contains(keyParam) & !csvMarkList.Contains(keyParam))
                            {
                                mark4Numbers.Add(keyParam);
                            }
                            else
                            {
                                bool control = true;
                                while (control)
                                {
                                    int randomNumber = ran.Next(0, 9999);
                                    if (!mark4Numbers.Contains(randomNumber) & !csvMarkList.Contains(randomNumber))
                                    {
                                        mark4Numbers.Add(randomNumber);
                                    }
                                    control = false;
                                }
                            }
                        }
                    }
                }

                //------Store data in csv ------
                var csv = new StringBuilder();
                for (int i = 0; i < markKeyParam.Count; i++)
                {
                    if(!csvMarkList.Contains(mark4Numbers[i]))
                    {
                        var newLine = string.Format("{0};{1}", markKeyParam[i], mark4Numbers[i] + Environment.NewLine);
                        csv.Append(newLine);
                    }
                }
                File.AppendAllText(pathForCsv, csv.ToString());
                
                //------ Set new Mark 4 -------
                Transaction trans = new Transaction(doc);
                trans.Start("Mark 4");

                SetMark(mark4Numbers, panelIdList, doc);

                uidoc.Selection.SetElementIds(systemPanelIds); // Highlight Systam panels in Revit UI

                trans.Commit();

                return Result.Succeeded;
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                uidoc.Selection.SetElementIds(errorElementId);
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
