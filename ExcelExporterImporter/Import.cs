#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
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

            // ================= Import CSVs =================
            string[] csvFilePaths = GetCsvFilePath(); // Get CSV file paths
            if (csvFilePaths == null)
            {
                TaskDialog.Show("INFO", "You didn't select any CSV file");
                return Result.Cancelled;
            }//Tell user no file was selected and stop process

            var _curDocScheduleNames = GetAllScheduleNames(doc); // Get all the schedules names in current doc

            string csvScheduleNamesFound = null;    // Tran Found Schedules
            string csvScheduleNamesNotFound = null; // Track Not Found Schedules
            foreach (var csvFilePath in csvFilePaths)  // Loop Through all the selected CSVs
            {
                var _viewScheduleNameFromCSV = GetLineFromCSV(csvFilePath, 1)[0];  // Get View schedule name from csv
                if (_curDocScheduleNames.Contains(_viewScheduleNameFromCSV))
                {
                    Debug.Print($"Schedule: {_viewScheduleNameFromCSV} - Found in current document!");

                    var _headersFromCSV = GetLineFromCSV(csvFilePath, 2);                   // Get Headers from csv
                    List<string[]> _viewScheduledata = ImportCSVToStringList(csvFilePath);  // Get data from csv - skips the first 3 lines
                    csvScheduleNamesFound += $"{_viewScheduleNameFromCSV}\n";               // add found schedule to csvScheduleNamesFound for later report.

                    using (Transaction tx = new Transaction(doc, $"Update {_viewScheduleNameFromCSV} Parameters")) // Start a new transaction to make changes to the elements in Revit
                    {
                        tx.Start(); // Lock the doc while changes are made in the transaction

                        // UPDATE THE SCHEDULES FROM PRIVIOSLY EXPORTED CSV FILES.
                        // THIS WILL ONLY UPDATE STRING-TYPE FIELDS THAT ARE NOT READONLY.
                        var _viewScheduleUpdateResult = _UpdateViewSchedule(doc, _viewScheduleNameFromCSV, _headersFromCSV, _viewScheduledata);
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