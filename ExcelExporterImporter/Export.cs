#region Namespaces
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using OfficeOpenXml;

using ORH_ExcelExporterImporter.Forms;
#endregion

namespace ORH_ExcelExporterImporter
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
            // open form
            MyForm currentForm = new MyForm()
            {
                Width = 400,
                Height = 200,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            //Get form data and do something
            #endregion

            // Create Revit_Exports on desktop if it doesn't exist
            string _FolderName = "Revit_Exports";
            var _path = _CreateFolderOnDesktopByName(_FolderName);
            string docName = doc.Title; // Name of the revit document
            string _excelFilePath = $"{_path}\\{docName}_Schedules.xlsx"; // Path of the excel file to be output



            // ================= Get All Schedules =================
            var _schedulesList = _GetSchedulesList(doc); // Get all the Schedules into a list


            // Open schedulesImport_Form1
            SchedulesImport_Form schedulesImport_Form1 = new SchedulesImport_Form()
            {
                Width = 1000,
                Height = 800,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            schedulesImport_Form1.dataGrid.ItemsSource = _schedulesList.Select(schedule => schedule.Name).ToList();
            schedulesImport_Form1.btn_Import.Content = "Export";
            schedulesImport_Form1.ShowDialog();
            if (schedulesImport_Form1.DialogResult == true)
            {
                var selectedSchedulenames = schedulesImport_Form1.dataGrid.SelectedItems.Cast<string>().ToList();
                List<ViewSchedule> selectedSchedules = new List<ViewSchedule>();
                foreach (var scheduleName in selectedSchedulenames)
                {
                    // Find the selected schedule by name
                    selectedSchedules.Add(_schedulesList.FirstOrDefault(sch => sch.Name == scheduleName));
                }
                _schedulesList = selectedSchedules;
            }
            else
            {
                return Result.Cancelled;
            }

            //if (true) { return Result.Cancelled; }
            //
            bool userCancelled = M_TellTheUserIfFileExistsOrIsOpen(_excelFilePath);
            if (userCancelled) { return Result.Cancelled; }

            M_ShowCurrentFormForNSeconds(currentForm, 5);



            ExcelPackage excelFile = Create_ExcelFile(_excelFilePath);
            ExcelWorkbook workbook = excelFile.Workbook;  // Get the workbook from the Excel package
            int prefix = 1;
            using (Transaction t = new Transaction(doc, "Exported Schedules"))
            {
                t.Start();
                foreach (ViewSchedule schedule in _schedulesList)
                {
                    // set the schedule to show tile and headers returns de original schedule definition
                    ScheduleDefinition curScheduleDefinition = MyUtils.M_ShowHeadersAndTileOnSchedule(schedule);
                    // Get all the categegories that allow AllowsBoundParameters as a set 
                    CategorySet _scheduleBuiltInCategory = M_GetAllowBoundParamCategorySet(doc, schedule);
                    // Add the "Dev_Text_1" parameter to be used for the UniqueID of the row element during export.
                    M_Add_Dev_Text_4(app, doc, schedule, _scheduleBuiltInCategory);

                    // create excel sheet based on schedule name and number prefix to avoid duplicates
                    ExcelWorksheet worksheet = workbook.Worksheets.Add($"{prefix}_{schedule.Name}");
                    // load current schedule to its own excel sheet
                    ExportViewScheduleBasic(schedule, worksheet);
                    prefix++;
                }
                t.RollBack();
            }
            excelFile.Save();
            excelFile.Dispose();
            Process.Start(_excelFilePath);
            return Result.Succeeded;

        }



        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
