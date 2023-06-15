#region Namespaces
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using OfficeOpenXml;
#endregion

namespace ORH_ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Export_OLD2 : MyUtils, IExternalCommand
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

            // Create Revit_Exports on desktop if it doesn't exist
            string _FolderName = "Revit_Exports";
            var _path = _CreateFolderOnDesktopByName(_FolderName);

            // ================= Get All Schedules =================
            var _schedulesList = M_GetSchedulesList(doc); // Get all the Schedules into a list

            // ================= Get Specific Schedule =================
            //var _schedulesList = M_GetSchedulesList(doc).Where(x => x.Name == "Mechanical Equipment Schedule"); // Get specific Schedule into a list
            //var _schedulesList = M_GetSchedulesList(doc).Where(x => x.Name == "ACCO Drawing Index - Coordination"); // Get specific Schedule into a list
            //var _schedulesList = M_GetSchedulesList(doc).Where(x => x.Name == "ACCO Drawing Index - Construction Documents"); // Get specific Schedule into a list
            //var schedule = M_GetSchedulesList(doc).Where(x => x.Name == "ACCO Drawing Index - Coordination") as ViewSchedule; // Get specific Schedule into a list
            //var schedule = M_GetSchedulesList(doc).FirstOrDefault(x => x.Name == "ACCO Drawing Index - Coordination") as ViewSchedule;
            //var schedules = M_GetSchedulesList(doc).FirstOrDefault(x => x.Name == "ACCO Drawing Index - Construction Documents") as ViewSchedule;
            //var _schedulesList = M_GetSchedulesList(doc).Where(x => x.Name == "ACCO Drawing Index - Coordination"); // Get specific Schedule into a list
            //var _schedulesList = M_GetSchedulesList(doc).Where(x => x.Name == "VARIABLE VOLUME BOX - DDC HOT WATER REHEAT SCHEDULE"); // Get specific Schedule into a list


            //var schedule = _schedulesList[7];

            #region Testin
            if (true)
            {
                string docName = doc.Title;
                string _excelFilePath = $"{_path}\\{docName}.xlsx";
                if (File.Exists(_excelFilePath)) File.Delete(_excelFilePath); // If the file exists, delete it.

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
            #endregion

            // Exported schedule names will be added to "_exportedSchedules" to notify the user at the end.
            string _exportedSchedules = "";

            // set TRUE only for testing a piece of code without running the entire class
            if (false)
            {
                //Place the code you wanto test here 

                return Result.Succeeded;
            }

            // ================= ExportSelectedSchedules =================
            using (Transaction t = new Transaction(doc, "Exported Schedules"))
            {
                foreach (var _curViewSchedule in _schedulesList)
                {
                    t.Start();

                    // set the schedule to show tile and headers returns de original schedule definition
                    ScheduleDefinition curScheduleDefinition = MyUtils.M_ShowHeadersAndTileOnSchedule(_curViewSchedule);

                    // Get the BuiltInCategory of the current schedule
                    //BuiltInCategory _scheduleBuiltInCategory = M_GetScheduleBuiltInCategory(doc, _curViewSchedule);
                    //M_Add_Dev_Text_2(app, doc, _curViewSchedule, _scheduleBuiltInCategory);

                    // Get all the categegories that allow AllowsBoundParameters as a set 
                    CategorySet _scheduleBuiltInCategory = M_GetAllowBoundParamCategorySet(doc, _curViewSchedule);
                    // Add the "Dev_Text_1" parameter to be used for the UniqueID of the row element during export.
                    M_Add_Dev_Text_4(app, doc, _curViewSchedule, _scheduleBuiltInCategory);

                    // Create a new instance of ViewScheduleExportOptions
                    ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();

                    // Set the field delimiter to ","
                    exportOptions.FieldDelimiter = ",";

                    // Set the text qualifier to double quotes
                    exportOptions.TextQualifier = ExportTextQualifier.DoubleQuote;

                    // Set the column headers to export only one row
                    exportOptions.ColumnHeaders = ExportColumnHeaders.OneRow;

                    // Exclude group headers, footers, and blank lines from the export
                    exportOptions.HeadersFootersBlanks = false;
                    //

                    string _fileName = $"{_curViewSchedule.Name}.csv";
                    _curViewSchedule.Export(_path, _fileName, exportOptions); // exports schedule
                    Thread.Sleep(400);

                    //Process.Start(Path.Combine(_path, _name)); // opens exported file

                    //var _listOfUniqueIds = _listOfUniqueIdsInScheduleView(doc, _curViewSchedule);
                    //string _curFilePath = $"{_path}\\{_name}";
                    //AddUniqueIdColumnToViewScheduleCsv(_curFilePath, _listOfUniqueIds);

                    //Update the Result string for the output TaskDialog
                    _exportedSchedules += $"{_curViewSchedule.Name}\n";

                    #region Update the output CSV with additional changes.
                    string _fileFullPath = $"{_path}\\{_fileName}";
                    M_MoveCsvLastColumnToFirst(_fileFullPath);

                    M_AddScheduleUniqueIdToCellA1(_fileFullPath, $"{_curViewSchedule.UniqueId}");
                    #endregion

                    t.RollBack();
                }

                // Notify the user of the exported schedules.
                // using my _MyTaskDialog Method. Removes the prefix on the Title
                M_MyTaskDialog("Exported Schedules:", _exportedSchedules);
            }
            // Open Windows Explorer to the folder path
            Process.Start("explorer.exe", _path);

            return Result.Succeeded;
        }




        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
