#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
#endregion

namespace ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Export : MyUtils, IExternalCommand
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

            // ================= GetAllSchedules =================
            var _schedulesList = _GetSchedulesList(doc); // Get all the Schedules into a list
            foreach (var _curViewSchedule in _schedulesList)
            {
                //// Create a ViewScheduleExportOptions object
                ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();
                exportOptions.FieldDelimiter = ",";

                string _path = @"C:\Users\ohernandez\Desktop\Revit_Exports";
                string _name = $"{_curViewSchedule.Name}.csv";
                _curViewSchedule.Export(_path, _name, exportOptions); // exports schedule
                //Process.Start(Path.Combine(_path, _name)); // opens exported file

                var _listOfUniqueIds = _listOfUniqueIdsInScheduleView(doc, _curViewSchedule);
                string _curFilePath = $"{_path}\\{_name}";
                AddUniqueIdColumnToViewScheduleCsv(_curFilePath, _listOfUniqueIds);
            }

            return Result.Succeeded;
        }

        //public void AddUniqueIdColumnToCsv(string filePath, string[] uniqueIds)
        public void AddUniqueIdColumnToViewScheduleCsv(string filePath, List<string> uniqueIds)
        {
            // Wait for 1 second
            Thread.Sleep(200);

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("The specified file does not exist.", filePath);
            }

            var csvLines = File.ReadAllLines(filePath);
            if (csvLines.Length < 3)
            {
                throw new InvalidOperationException("The specified file does not have enough rows.");
            }

            // Add the UniqueID header to the first row
            csvLines[0] = csvLines[0] + ",";       // Update Row 0
            csvLines[1] = "UniqueID," + csvLines[1];  // Update Row 1
            csvLines[2] = csvLines[2] + ",";        // Update Row 2

            // Add the UniqueID values to each subsequent row
            for (int i = 3; i < csvLines.Length; i++)
            {
                csvLines[i] = uniqueIds[i - 3] + "," + csvLines[i];
            }

            // Write the modified CSV data to the same file
            File.WriteAllLines(filePath, csvLines);
        }


        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
