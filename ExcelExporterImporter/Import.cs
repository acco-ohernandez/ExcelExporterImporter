#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using static System.Net.Mime.MediaTypeNames;

#endregion

#region Begining of doc
namespace ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Import : MyUtils, IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            #region FormStuff
            //// open form
            //MyForm currentForm = new MyForm()
            //{
            //    Width = 800,
            //    Height = 450,
            //    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
            //    Topmost = true,
            //};

            //currentForm.ShowDialog();

            // get form data and do something
            #endregion
            #endregion

            #region Testing
            // This if statement is just to test some code without 
            // running the entire code.
            // Change to true or false depending if you want to execute or not.
            if (false) // <---- Don't for get to change it if you need to.
            {
                //string uid = "dc86627d-cf12-49fe-bdad-488a619b34a1-00060aca"; // row Elem UniqueId - NO GOOD
                string uid = "7a2419bd-e042-4b38-8b95-781bc33e7dd8-000854c1"; // schedule ID - GOOD
                var elementsreturned = _GetViewScheduleBasedOnUniqueId(doc, uid);
                foreach (var elem in elementsreturned)
                {
                    Debug.Print($"Elem Name: {elem.Name} " +
                                $"UniqueId : {elem.UniqueId}");
                }

                return Result.Cancelled;  
            }
            #endregion


            // ================= GetAllSchedules =================
            var _schedulesList = _GetSchedulesList(doc); // Get all the Schedules into a list
            //Get schedule by array possition in _schedulesList;
            var _selectedSchedule = _schedulesList[6];
            Debug.Print($"Selected Schedule Name: {_selectedSchedule.Name}");
            //return Result.Cancelled;

            // Get list of row elements
            List<Element> _elementsList = MyUtils.GetElementsOnScheduleRow(doc, _selectedSchedule);
            //List<ElementId> td = _elementsList.Select(e => e.Id).ToList();

            using (Transaction tx = new Transaction(doc, "Update Parameters")) // Start a new transaction to make changes to the elements in Revit
            {
                tx.Start(); // Lock the doc while changes are made in the transaction

                int numCount = 0; // Initialize a counter to keep track of the number of elements processed
                var rowElemIds = _elementsList.Select(e => e.Id);  // Get the IDs of all the row elements in the selected schedule
                foreach (var rowElemId in rowElemIds)
                {
                    var rowElem = doc.GetElement(rowElemId); // Get the element from its ID
                    Debug.Print($"{rowElem.Name} {rowElem.UniqueId}");


                    ParameterSet paramSet = rowElem.Parameters; // Get the parameters of the element
                    // Print the name and value of each parameter
                    foreach (Parameter parameter in paramSet)
                    {
                        string paramName = parameter.Definition.Name; // Get the name of the parameter
                        string paramValue = parameter.AsValueString(); // Get the value of the parameter as a string

                        Debug.Print($"Parameter_Name: {paramName} \n" +
                                   $"Parameter_Value: {paramValue} \n" +
                                        $"IsReadOnly: {parameter.IsReadOnly}\n" +
                            $"====================================");

                        numCount++;

                        // Update all the values from the Test2 Parameter to "Romeo"
                        //if (paramName == "Test2") // use this one if there is a parameter call Test2
                        if (paramName != null && parameter.IsReadOnly == false) // Only attempt to make shanges if the parameter Is Not ReadyOnly
                        {
                            // Set the value of the parameter 
                            parameter.Set($"{numCount} - ElemID: {parameter.Id}");
                        }
                    }
                }

                tx.Commit(); // Commit the changes made in the transaction
            }

            return Result.Succeeded;
        }


        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}