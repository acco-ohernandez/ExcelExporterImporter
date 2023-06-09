#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace ORH_ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Export_backups : MyUtils, IExternalCommand
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

            //string _path = @"C:\Users\ohernandez\Desktop\Revit_Exports";
            string _FolderName = "Revit_Exports";
            var _path = _CreateFolderOnDesktopByName(_FolderName);

            // ================= GetAllSchedules =================
            //var _schedulesList = M_GetSchedulesList(doc); // Get all the Schedules into a list
            var _schedulesList = M_GetSchedulesList(doc).Where(x => x.Name == "Mechanical Equipment Schedule"); // Get specific Schedule into a list

            string _exportedSchedules = "";
            using (Transaction t = new Transaction(doc, "Added param to sched"))
            {
                t.Start();
                foreach (var _curViewSchedule in _schedulesList)
                {
                    BuiltInCategory _scheduleBuiltInCategory = M_GetScheduleBuiltInCategory(doc, _curViewSchedule);
                    M_Add_Dev_Text_2(app, doc, _curViewSchedule, _scheduleBuiltInCategory);

                    //// Create a ViewScheduleExportOptions object
                    ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();
                    exportOptions.Title = true;
                    exportOptions.FieldDelimiter = ",";
                    //exportOptions.ColumnHeaders = ExportColumnHeaders.OneRow;
                    exportOptions.TextQualifier = ExportTextQualifier.DoubleQuote;
                    exportOptions.ColumnHeaders = ExportColumnHeaders.OneRow;
                    exportOptions.HeadersFootersBlanks = false;


                    string _fileName = $"{_curViewSchedule.Name}.csv";
                    _curViewSchedule.Export(_path, _fileName, exportOptions); // exports schedule
                    Thread.Sleep(500);

                    //Process.Start(Path.Combine(_path, _name)); // opens exported file

                    //var _listOfUniqueIds = _listOfUniqueIdsInScheduleView(doc, _curViewSchedule);
                    //string _curFilePath = $"{_path}\\{_name}";
                    //AddUniqueIdColumnToViewScheduleCsv(_curFilePath, _listOfUniqueIds);
                    //######
                    _exportedSchedules += $"{_curViewSchedule.Name}\n";

                    M_MoveCsvLastColumnToFirst($"{_path}\\{_fileName}");
                }
                // using my _MyTaskDialog Method. Removes the prefix on the Title
                M_MyTaskDialog("Exported Schedules:", _exportedSchedules);

                t.Commit();
            }
            // Open Windows Explorer to the folder path
            System.Diagnostics.Process.Start("explorer.exe", _path);

            return Result.Succeeded;
        }




        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
