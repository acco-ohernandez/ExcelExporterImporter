#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;

#endregion

namespace ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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
            var _selectedSchedule = _schedulesList[3];   // Get schedule by array possition in _schedulesList ;
            //var _selectedSchedule = _schedulesList.FirstOrDefault(x => x.Name == "Electrical Equipment Connection Schedule"); // Get schedule by name from _schedulesList
            Debug.Print($"Working on ViewSchedule: {_selectedSchedule.Name}"); // This line is only for debug output

            // Get the schedule rows data list
            var _scheduleTableDataAsString = _GetScheduleTableDataAsString(_selectedSchedule, doc);

            // This line is only for debug, outputs all the rows in _scheduleTableDataAsString
            int lineN = 0; foreach (var line in _scheduleTableDataAsString){lineN++; Debug.Print($"Row {lineN}: {line}");}

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // Get Desktop Path
            string filePath = Path.Combine(desktopPath, $"{_selectedSchedule.Name}.txt");      // Combine DesktopPath with Schedule Name for filePath

            File.WriteAllLines(filePath, _scheduleTableDataAsString);  // Write the contents of the _scheduleRowList to the file
            Process.Start(filePath);                                   // open the outputed file

            return Result.Succeeded;
        }

        /// <summary>
        /// Takes in one ViewSchedule and the Autodesk.Revit.DB.Document
        /// </summary>
        /// <param name="_selectedSchedule"></param>
        /// <param name="doc"></param>
        /// <returns>Returns a comma delimited string list of all the schedule data</returns>
        private List<string> _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule, Autodesk.Revit.DB.Document doc)
        {
            //---Gets the list of items/rows in the schedule---
            var _scheduleItemsCollector = new FilteredElementCollector(doc, _selectedSchedule.Id).WhereElementIsNotElementType();
            List<string> uid = new List<string>();
            foreach (var _scheduleRow in _scheduleItemsCollector)
            {
                Debug.Print($"{_scheduleRow.GetType()} - UID: {_scheduleRow.UniqueId}");
                uid.Add(_scheduleRow.UniqueId); // adds each GUID to the uid list
            }


            // Get the data from the ViewSchedule object
            TableData tableData = _selectedSchedule.GetTableData();
            TableSectionData data = tableData.GetSectionData(SectionType.Body);

            // Concatenate the data into a list of strings, with each string representing a row
            List<string> rows = new List<string>();
            for (int row = 0; row < data.NumberOfRows; row++)
            {
                int r = row - 2;
                StringBuilder sb = new StringBuilder();
                for (int col = 0; col < data.NumberOfColumns; col++)
                {
                    string cellText = data.GetCellText(row, col).ToString();
                    if (col == 0 && row >= 2) // if first column, add the UniqueID
                    {
                        string uniqueId = uid[r];
                        //string uniqueId = data.GetCellElement(row, col).UniqueId;
                        //sb.Append(uniqueId + ",");
                        sb.Append($"{uniqueId},");
                        // Debug.Print($"{_GetElementBasedOnUniqueId(doc, uniqueId)}"); // still working on method 
                    }
                    sb.Append(cellText + ",");
                }
                rows.Add(sb.ToString());
            }

            return rows;
        }
        #region Testing methods

        //private List<string> _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule)
        //{
        //    List<string> rows = new List<string>();

        //    // Get the data from the ViewSchedule object
        //    TableData tableData = _selectedSchedule.GetTableData();
        //    TableSectionData data = tableData.GetSectionData(SectionType.Body);

        //    // Get the first row
        //    TableRowData firstRow = data.GetCellElement(0, 0).Row;

        //    // Concatenate the data into a list of comma-separated strings
        //    for (int row = 0; row < data.NumberOfRows; row++)
        //    {
        //        // Get the row data
        //        TableRowData rowData = firstRow.Offset(row, 0);

        //        // Get the row ID
        //        string rowId = rowData.UniqueId;

        //        // Concatenate the data into a comma-separated string
        //        StringBuilder sb = new StringBuilder();
        //        for (int col = 0; col < data.NumberOfColumns; col++)
        //        {
        //            string cellText = data.GetCellText(row, col).ToString();
        //            sb.Append(cellText + ",");
        //        }

        //        // Add the row ID to the beginning of the string
        //        sb.Insert(0, rowId + ",");

        //        // Add the row to the list
        //        rows.Add(sb.ToString());
        //    }

        //    return rows;
        //}


        //private List<string> _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule)
        //{
        //    // Get the data from the ViewSchedule object
        //    TableData tableData = _selectedSchedule.GetTableData();
        //    TableSectionData data = tableData.GetSectionData(SectionType.Body);

        //    // Collect the data into a list of comma-separated strings per row
        //    List<string> rows = new List<string>();
        //    var scheduleItemsCollector = new FilteredElementCollector(_selectedSchedule.Document, _selectedSchedule.Id).WhereElementIsNotElementType();
        //    foreach (var scheduleRow in scheduleItemsCollector)
        //    {
        //        var scheduleRowId = scheduleRow.UniqueId;
        //        StringBuilder sb = new StringBuilder();
        //        sb.Append(scheduleRowId + ",");
        //        for (int col = 0; col < data.NumberOfColumns; col++)
        //        {
        //            string cellText = data.GetCellText(scheduleRow.RowNumber, col).ToString();
        //            sb.Append(cellText + ",");
        //        }
        //        rows.Add(sb.ToString());
        //    }

        //    return rows;
        //}


        //private List<string> _GetScheduleTableDataAsString2(ViewSchedule _selectedSchedule) //almost
        //{
        //    // Get the data from the ViewSchedule object
        //    TableData tableData = _selectedSchedule.GetTableData();
        //    TableSectionData data = tableData.GetSectionData(SectionType.Body);

        //    // Concatenate the data into a list of strings, with each string representing a row
        //    List<string> rows = new List<string>();
        //    for (int row = 0; row < data.NumberOfRows; row++)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        string uniqueId = data.GetCellText(row, 0).ToString(); // get UniqueID from first column of each row
        //        sb.Append(uniqueId + ",");
        //        for (int col = 1; col < data.NumberOfColumns; col++) // start from second column
        //        {
        //            string cellText = data.GetCellText(row, col).ToString();
        //            sb.Append(cellText + ",");
        //        }
        //        rows.Add(sb.ToString());
        //    }

        //    return rows;

        //}
        //private List<string> _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule) //works
        //{
        //    // Get the data from the ViewSchedule object
        //    TableData tableData = _selectedSchedule.GetTableData();
        //    TableSectionData data = tableData.GetSectionData(SectionType.Body);

        //    // Concatenate the data into a list of strings, with each string representing a row
        //    List<string> rows = new List<string>();
        //    for (int row = 0; row < data.NumberOfRows; row++)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        for (int col = 0; col < data.NumberOfColumns; col++)
        //        {
        //            string cellText = data.GetCellText(row, col).ToString();
        //            sb.Append(cellText + ",");
        //        }
        //        rows.Add(sb.ToString());
        //    }

        //    return rows;
        //}

        //private string _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule)
        //{
        //    // Get the data from the ViewSchedule object
        //    TableData tableData = _selectedSchedule.GetTableData();
        //    TableSectionData data = tableData.GetSectionData(SectionType.Body);

        //    // Concatenate the data into a comma-separated string
        //    StringBuilder sb = new StringBuilder();
        //    for (int row = 0; row < data.NumberOfRows; row++)
        //    {
        //        for (int col = 0; col < data.NumberOfColumns; col++)
        //        {
        //            string cellText = data.GetCellText(row, col).ToString();
        //            sb.Append(cellText + ",");
        //        }
        //        sb.AppendLine();
        //    }

        //    return sb.ToString();
        //}

        //private string _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule)
        //{
        //    // Get the data from the ViewSchedule object
        //    TableData tableData = _selectedSchedule.GetTableData();
        //    TableSectionData data = tableData.GetSectionData(SectionType.Body);

        //    // Concatenate the data into a comma-separated string
        //    StringBuilder sb = new StringBuilder();
        //    for (int row = 0; row < data.NumberOfRows; row++)
        //    {
        //        for (int col = 0; col < data.NumberOfColumns; col++)
        //        {
        //            string cellText = data.GetCellText(row, col).ToString();
        //            sb.Append(cellText + ",");
        //        }
        //        sb.AppendLine();
        //    }

        //    return sb.ToString();
        //} 
        #endregion end of testing methods


        private ViewSchedule _GetElementBasedOnUniqueId(Autodesk.Revit.DB.Document doc, string _uniqueIdString) // still not working
        {
            string uniqueId = _uniqueIdString; // Replace with the actual UniqueID of the element
            ElementId elementId = ElementId.InvalidElementId;

            // Loop through all elements in the document to find the one with the given UniqueID
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            foreach (Element _element in collector)
            {
                if (_element.UniqueId == uniqueId)
                {
                    elementId = _element.Id;
                    break;
                }
            }

            // Get the element using the ElementId
            //Element element = doc.GetElement(elementId);
            ViewSchedule viewScheduleElement = doc.GetElement(elementId) as ViewSchedule;
            return viewScheduleElement;
        }

        /// <summary>
        /// Takes a ViewSchedule Element and gets the TableData from it.
        /// </summary>
        /// <param name="curSchedule"></param>
        /// <returns>viewSchedule TableData</returns>
        private TableData _GetSchedulesTablesList(Element viewScheduleElement)
        {
            ViewSchedule _curViewSchedule = viewScheduleElement as ViewSchedule;
            TableData _scheduleTableData = _curViewSchedule.GetTableData() as TableData;
            return _scheduleTableData;
        }

        private List<ViewSchedule> _GetSchedulesList(Autodesk.Revit.DB.Document doc) // This method returns a list of all the schedule elements
        {
            List<ViewSchedule> _schedulesList = new List<ViewSchedule>();
            FilteredElementCollector _schedules = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
            foreach (ViewSchedule viewSchedule in _schedules)
            {
                if (viewSchedule.IsTitleblockRevisionSchedule) continue;
                _schedulesList.Add(doc.GetElement(viewSchedule.Id) as ViewSchedule);
            }
            foreach (Element element in _schedulesList) // Debug Print
            { Debug.Print($"Method: _GetSchedulesList : {element.Name}"); } // 
            return _schedulesList;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
