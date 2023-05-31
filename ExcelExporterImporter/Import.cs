#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Windows.Input;
using System.Windows.Shapes;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

using Microsoft.VisualBasic.FileIO;
#endregion

#region Begining of doc
namespace ORH_ExcelExporterImporter
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

            //Testing
            if (false)
            {
                //using (Transaction t = new Transaction(doc, "Added param to sched"))
                //{
                //    t.Start();
                //    var af = M_AddByNameAvailableFieldToSchedule(doc, "Mechanical Equipment Schedule", "Count");
                //    t.Commit();
                //}

                //var allBiltInCats = _GetAllBuiltInCategories();

                var schedule = M_GetViewScheduleByName(doc, "Mechanical Equipment Schedule");

                BuiltInCategory _scheduleBuiltInCategory = M_GetScheduleBuiltInCategory(doc, schedule);

                //M_Add_Dev_Text_1(app, doc, "Mechanical Equipment Schedule", biCat);
                M_Add_Dev_Text_1(app, doc, schedule.Name, _scheduleBuiltInCategory);


                //var schedule = M_GetViewScheduleByName(doc, "Mechanical Equipment Schedule");
                //var fistElemCategory = _GetElementsOnScheduleRow(doc, schedule).FirstOrDefault();

                //BuiltInCategory elementBuiltInCategory = _GetElementBuiltInCategory(fistElemCategory);
                //BuiltInCategory cat = _GetScheduleBuiltInCategory(schedule);


                //M_Add_Dev_Text_1(app, doc, "Mechanical Equipment Schedule", elementBuiltInCategory);
                return Result.Succeeded;
            }

            // ================= Import CSVs =================
            string[] csvFilePaths = GetCsvFilePath(); // Get CSV file paths
            if (csvFilePaths == null)
            {
                TaskDialog.Show("INFO", "You didn't select any CSV file");
                return Result.Cancelled;
            }//Tell user no file was selected and stop process

            //var _curDocScheduleNames = GetAllScheduleNames(doc); // Get all the schedules names in current doc
            var _curDocSchedulesUniqueIds = GetAllScheduleUniqueIds(doc); // Get all the schedules names in current doc

            string csvScheduleNamesFound = null;    // Track Found Schedules
            string csvScheduleNamesNotFound = null; // Track Not Found Schedules

            foreach (var csvFilePath in csvFilePaths)  // Loop Through all the selected CSVs
            {
                CheckAndPromptToCloseExcel(csvFilePath); // Tell the user to close excel before continueing

                string _viewScheduleUniqueIdFromCSV = null;
                try { _viewScheduleUniqueIdFromCSV = M_GetLinesFromCSV(csvFilePath, 1)[0]; }  // Get View schedule name from csv
                catch (Exception) { continue; }

                var _viewScheduleNameFromCSV = M_GetLinesFromCSV(csvFilePath, 1)[1];  // Get View schedule UniqueId from csv

                // Check if the current Schedule UniqueID from the CSV is found in _curDocSchedulesUniqueIds
                if (_curDocSchedulesUniqueIds.Contains(_viewScheduleUniqueIdFromCSV))
                {
                    Debug.Print($"Schedule: {_viewScheduleNameFromCSV} - Found in current document!");

                    var _headersFromCSV = M_GetLinesFromCSV(csvFilePath, 2);                   // Get Headers from csv
                    List<string[]> _viewScheduledata = ImportCSVToStringList2(csvFilePath);  // Get data from csv - skips the first 3 lines
                    csvScheduleNamesFound += $"{_viewScheduleNameFromCSV}\n";               // add found schedule to csvScheduleNamesFound for later report.


                    using (Transaction tx = new Transaction(doc, $"Update {_viewScheduleNameFromCSV} Parameters")) // Start a new transaction to make changes to the elements in Revit
                    {
                        tx.Start(); // Lock the doc while changes are made in the transaction

                        // UPDATE THE SCHEDULES FROM PRIVIOSLY EXPORTED CSV FILES.
                        // THIS WILL ONLY UPDATE STRING-TYPE FIELDS THAT ARE NOT READONLY.
                        var _viewScheduleUpdateResult = _UpdateViewSchedule(doc, _viewScheduleUniqueIdFromCSV, _headersFromCSV, _viewScheduledata);
                        tx.Commit();
                    }
                }
                else
                {
                    Debug.Print($"Schedule: {_viewScheduleNameFromCSV} - Not found in current document!");
                    csvScheduleNamesNotFound += $"{_viewScheduleNameFromCSV}\n"; // add found schedule to csvScheduleNamesNotFound for later report.
                }

            }
            if (csvScheduleNamesFound != null)
            {
                TaskDialog.Show("INFO", $"Updated the following Schedule(s):\n\n{csvScheduleNamesFound}");
            }
            if (csvScheduleNamesNotFound != null)
            {
                TaskDialog.Show("INFO", $"Could Not find the following Schedule(s):\n\n{csvScheduleNamesNotFound}");
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