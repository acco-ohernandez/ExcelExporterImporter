#region Namespaces
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

#endregion

namespace ORH_ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Export_Old : MyUtils, IExternalCommand
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
            var _schedulesList = M_GetSchedulesList(doc); // Get all the Schedules into a list

            // Uncommen the option you want to use
            //==== Option 1 ==== Get schedule by array possition in _schedulesList ;
            var _selectedSchedule = _schedulesList[6];
            //////==== Option 2 ==== Get schedule by name from _schedulesList
            ////string _scheduleNameToSelect = "Electrical Equipment Connection Schedule";
            ////var _selectedSchedule = _schedulesList.FirstOrDefault(x => x.Name == _scheduleNameToSelect); 
            Debug.Print($"Working on ViewSchedule: {_selectedSchedule.Name} ID: {_selectedSchedule.Id}"); // This line is only for debug output

            // Get the schedule rows data list
            var _scheduleTableDataAsString = _GetScheduleTableDataAsString(_selectedSchedule, doc);

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // Get Desktop Path
            string filePath = Path.Combine(desktopPath, $"{_selectedSchedule.Name}.txt");      // Combine DesktopPath with Schedule Name for filePath

            File.WriteAllLines(filePath, _scheduleTableDataAsString);  // Write the contents of the _scheduleRowList to the file
            Process.Start(filePath);                                   // open the outputed file

            return Result.Succeeded;
        }



        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
