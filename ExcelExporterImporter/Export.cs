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

            // Create Revit_Exports on desktop if it doesn't exist
            string _FolderName = "Revit_Exports";
            var _path = _CreateFolderOnDesktopByName(_FolderName);

            // ================= Get All Schedules =================
            var _schedulesList = _GetSchedulesList(doc); // Get all the Schedules into a list

            // ================= Get Specific Schedule =================
            //var _schedulesList = _GetSchedulesList(doc).Where(x => x.Name == "VARIABLE VOLUME BOX - DDC HOT WATER REHEAT SCHEDULE"); // Get specific Schedule into a list

            //var schedule = _schedulesList[7];


            string docName = doc.Title;
            string _excelFilePath = $"{_path}\\{docName}_Schedules.xlsx"; // Path of the excel file to be output
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




        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
