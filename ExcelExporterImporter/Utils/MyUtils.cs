using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;

namespace ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class MyUtils
    {
        //###########################################################################################

        public static List<ViewSchedule> _GetSchedulesList(Autodesk.Revit.DB.Document doc) // This method returns a list of all the schedule elements
        {
            int count = 0;
            List<ViewSchedule> _schedulesList = new List<ViewSchedule>();
            FilteredElementCollector _schedules = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
            foreach (ViewSchedule viewSchedule in _schedules)
            {
                if (viewSchedule.IsTitleblockRevisionSchedule) continue;
                _schedulesList.Add(doc.GetElement(viewSchedule.Id) as ViewSchedule);
            }
            foreach (Element element in _schedulesList) // Debug Print
            { Debug.Print($"Method: _GetSchedulesList : {count} - {element.Name}"); count++; } // 
            return _schedulesList;
        }
        //###########################################################################################


        //###########################################################################################
        /// <summary>
        /// _GetScheduleTableDataAsString Method: Takes in one ViewSchedule and the Autodesk.Revit.DB.Document
        /// </summary>
        /// <param name="_selectedSchedule"></param>
        /// <param name="doc"></param>
        /// <returns>Returns a comma delimited string list of all the schedule data</returns>
        /// 
        public static List<string> _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule, Autodesk.Revit.DB.Document doc)
        {
            //---Gets the list of items/rows in the schedule---
            var _scheduleItemsCollector = new FilteredElementCollector(doc, _selectedSchedule.Id).WhereElementIsNotElementType();
            List<string> uid = new List<string>(); // List to hold all the UniqueIDs in the schedule
            foreach (var _scheduleRow in _scheduleItemsCollector)
            {
                Debug.Print($"{_scheduleRow.GetType()} - UID: {_scheduleRow.UniqueId}"); // only for Debug
                uid.Add(_scheduleRow.UniqueId); // adds each UniqueID to the uid list
            }

            // Get the data from the ViewSchedule object
            TableData tableData = _selectedSchedule.GetTableData();
            TableSectionData _sectionData = tableData.GetSectionData(SectionType.Body);

            // Concatenate the data into a list of strings, with each string representing a row
            List<string> _rowsList = new List<string>();
            for (int row = 0; row < _sectionData.NumberOfRows; row++)
            {
                int _adjestedRow = row - 2;
                StringBuilder sb = new StringBuilder();
                if (row == 0)
                {
                    sb.Append("UniqueId,"); // adds the "UniqueId" header
                }
                if (row == 1)
                {
                    sb.Append(","); // adds the "," on the second line. This is because schedules second row is empty.
                }
                for (int col = 0; col < _sectionData.NumberOfColumns; col++)
                {
                    string cellText = _sectionData.GetCellText(row, col).ToString();
                    var cellTextColIndex = _sectionData.GetCellText(row, col);

                    if (col == 0 && row >= 2) // if first column, add the UniqueID
                    {
                        string uniqueId = uid[_adjestedRow];
                        sb.Append($"{uniqueId},");
                    }
                    sb.Append(cellText + ",");
                }
                _rowsList.Add(sb.ToString());
            }
            // This line is only for debug, outputs all the rows in _scheduleTableDataAsString
            int lineN = 0; foreach (var line in _rowsList) { lineN++; Debug.Print($"Row {lineN}: {line}"); }

            return _rowsList;
        }
        //###########################################################################################

        //###########################################################################################

        public List<ViewSchedule> _GetViewScheduleBasedOnUniqueId(Autodesk.Revit.DB.Document doc, string _elementUniqueIdString) // still not working
        {
            // Assuming you have an active Revit document
            //string uniqueId = "dc86627d-cf12-49fe-bdad-488a619b34a1-00060aca";
            string uniqueId = _elementUniqueIdString;

            int count = 0;
            List<ViewSchedule> _schedulesList = new List<ViewSchedule>();
            FilteredElementCollector _schedules = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
            foreach (ViewSchedule viewSchedule in _schedules)
            {
                Debug.Print("Method: _GetElementBasedOnUniqueId =================");

                if (viewSchedule.IsTitleblockRevisionSchedule) continue;

                //_schedulesList.Add(doc.GetElement(viewSchedule.Id) as ViewSchedule);
                if (viewSchedule.UniqueId == uniqueId && viewSchedule != null)
                {
                    _schedulesList.Add(doc.GetElement(viewSchedule.Id) as ViewSchedule);
                }

            }
            foreach (Element element in _schedulesList) // Debug Print
            {
                Debug.Print($"Method: _GetElementBasedOnUniqueId : {count} - {element.Name} " +
                            $"Element UniqueId: {element.UniqueId}");
                count++;
            } // Debug Print

            return _schedulesList;
        }

        //###########################################################################################
        public static ViewSchedule _GetViewScheduleByName(Document doc, string viewScheduleName)
        {
            FilteredElementCollector _schedules = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
            ViewSchedule _viewScheduleNotFound = null;
            foreach (ViewSchedule curViewScheduleInDoc in _schedules)
            {
                if (curViewScheduleInDoc.Name == viewScheduleName)
                {
                    return curViewScheduleInDoc;
                }

            }
            return _viewScheduleNotFound;
        }
        //###########################################################################################
        public static List<Element> _GetElementsOnScheduleRow(Document doc, ViewSchedule selectedSchedule)
        {
            // Got the idea from this Youtube video
            // https://www.youtube.com/watch?v=H1Z3f1pgyPE

            Debug.Print($"=========== Biginning GetElementsOnScheduleRow Method");

            TableData tabledata = selectedSchedule.GetTableData();
            TableSectionData tableSectionData = tabledata.GetSectionData(SectionType.Body);
            List<ElementId> elemIds = new FilteredElementCollector(doc, selectedSchedule.Id).ToElementIds().ToList();
            List<Element> elementOnRow = new List<Element>();

            foreach (var elemId in elemIds)
            {
                var elem = doc.GetElement(elemId);
                Debug.Print($"===========Row ElementID: {elem.Id} : Element Name: {elem.Name}");
                elementOnRow.Add(elem);
            }
            Debug.Print($"=========== End of GetElementsOnScheduleRow Method");

            return elementOnRow;
        }

        //###########################################################################################
        /// <summary>
        /// Takes a ViewSchedule Element and gets the TableData from it.
        /// </summary>
        /// <param name="curSchedule"></param>
        /// <returns>viewSchedule TableData</returns>
        public TableData _GetSchedulesTablesList(Element viewScheduleElement)
        {
            ViewSchedule _curViewSchedule = viewScheduleElement as ViewSchedule;
            TableData _scheduleTableData = _curViewSchedule.GetTableData() as TableData;
            return _scheduleTableData;
        }
        //###########################################################################################

        public static List<string> _listOfUniqueIdsInScheduleView(Document doc, ViewSchedule _selectedSchedule)
        {
            //---Gets the list of items/rows in the schedule---
            var _scheduleItemsCollector = new FilteredElementCollector(doc, _selectedSchedule.Id).WhereElementIsNotElementType();
            List<string> uid = new List<string>(); // List to hold all the UniqueIDs in the schedule
            foreach (var _scheduleRow in _scheduleItemsCollector)
            {
                Debug.Print($"{_scheduleRow.GetType()} - UID: {_scheduleRow.UniqueId}"); // only for Debug
                uid.Add(_scheduleRow.UniqueId); // adds each UniqueID to the uid list
            }
            return uid;
        }


        //###########################################################################################
        //public static Element _getElementOnViewScheduleRowByUniqueId(Document doc, ViewSchedule _selectedSchedule, string _uniqueId)
        //{
        //    var _scheduleItemsCollector = new FilteredElementCollector(doc, _selectedSchedule.Id);
        //    var element = _scheduleItemsCollector.FirstOrDefault(e => e.UniqueId == _uniqueId);

        //    if (element != null)
        //    {
        //        Debug.Print($"Method _getElementOnViewScheduleRowByUniqueId: {element.GetType()} - UID: {element.UniqueId} === ELEMENT RETURNED!"); // only for Debug
        //    }
        //    else
        //    {
        //        Debug.Print($"Method _getElementOnViewScheduleRowByUniqueId: Element not found with UniqueId {_uniqueId}");
        //    }

        //    return element;
        //}
        public static Element _getElementOnViewScheduleRowByUniqueId(Document doc, ViewSchedule _selectedSchedule, string _uniqueId)
        {
            //---Gets the list of items/rows in the schedule---
            var _scheduleItemsCollector = new FilteredElementCollector(doc, _selectedSchedule.Id).Where(e => e.UniqueId == _uniqueId);

            if (_scheduleItemsCollector.Any())
            {
                int count = _scheduleItemsCollector.Count();
                Element _scheduleRow = _scheduleItemsCollector.FirstOrDefault() as Element;
                Debug.Print($"Method _getElementOnViewScheduleRowByUniqueId: {_scheduleRow.GetType()} - UID: {_scheduleRow.UniqueId} === ELEMENT RETURNED!");
                return _scheduleRow;
            }
            else
            {
                Debug.Print($"Method _getElementOnViewScheduleRowByUniqueId: Element with UID {_uniqueId} not found!");
                return null;
            }
        }
        //public static Element _getElementOnViewScheduleRowByUniqueId(Document doc, ViewSchedule _selectedSchedule, string _uniqueId)
        //{
        //    //---Gets the list of items/rows in the schedule---
        //    var _scheduleItemsCollector = new FilteredElementCollector(doc, _selectedSchedule.Id);//.WhereElementIsNotElementType();
        //    foreach (Element _scheduleRow in _scheduleItemsCollector)
        //    {
        //        if (_scheduleRow.UniqueId == _uniqueId)
        //        {
        //            Debug.Print($"Method _getElementOnViewScheduleRowByUniqueId: {_scheduleRow.GetType()} - UID: {_scheduleRow.UniqueId} === ELEMENT RETURNED!"); // only for Debug

        //            return _scheduleRow;
        //        }
        //        Debug.Print($"Method _getElementOnViewScheduleRowByUniqueId: {_scheduleRow.GetType()} - UID: {_scheduleRow.UniqueId} === Not Returned"); // only for Debug
        //    }
        //    return null;
        //}



        //###########################################################################################
        public static string[] GetCsvFilePath()
        {
            // Create a new OpenFileDialog object.
            OpenFileDialog dialog = new OpenFileDialog();

            // Set the dialog's filter to CSV files.
            dialog.Filter = "CSV Files (*.csv)|*.csv";
            dialog.Multiselect = true;
            dialog.Title = "Select CSV File";
            dialog.RestoreDirectory = false;

            // Show the dialog to the user.
            dialog.ShowDialog();
            var _fileNames = dialog.FileNames;
            // If the user selected a file, return its path.
            if (_fileNames.Count() > 0)
            {
                return _fileNames;
            }

            // Otherwise, return null.
            return null;
        }
        //###########################################################################################
        public static List<string[]> ImportCSVToStringList(string csvFilePath)
        {
            var dataList = new List<string[]>();

            using (StreamReader reader = new StreamReader(csvFilePath))
            {
                // Skip the first three lines (header and empty lines)
                reader.ReadLine();
                reader.ReadLine();
                reader.ReadLine();

                // Use a regular expression to split the line into fields only on commas that are not inside quotes
                Regex csvParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    string[] fields = csvParser.Split(line);

                    // Remove quotes from each field
                    for (int i = 0; i < fields.Length; i++)
                    {
                        fields[i] = fields[i].Trim('"');
                    }

                    dataList.Add(fields);
                }
            }

            return dataList;
        }

        //public static List<string[]> ImportCSVToStringList(string csvFilePath)
        //{
        //    var dataList = new List<string[]>();

        //    using (StreamReader reader = new StreamReader(csvFilePath))
        //    {
        //        // Skip the first three lines (header and empty lines)
        //        reader.ReadLine();
        //        reader.ReadLine();
        //        reader.ReadLine();

        //        while (!reader.EndOfStream)
        //        {
        //            string line = reader.ReadLine();
        //            string[] fields = line.Split(',');

        //            // Remove quotes from each field
        //            for (int i = 0; i < fields.Length; i++)
        //            {
        //                fields[i] = fields[i].Trim('"');
        //            }

        //            dataList.Add(fields);
        //        }
        //    }

        //    return dataList;
        //}
        //###########################################################################################


        public static string[] GetLineFromCSV(string csvFilePath, int lineNumber)
        {
            string[] lineFields = null;

            using (StreamReader reader = new StreamReader(csvFilePath))
            {
                // Read lines until we reach the specified line number
                for (int i = 1; i <= lineNumber; i++)
                {
                    string line = reader.ReadLine();
                    if (i == lineNumber)
                    {
                        lineFields = line.Split(',');

                        // Remove quotes from each field
                        for (int j = 0; j < lineFields.Length; j++)
                        {
                            lineFields[j] = lineFields[j].Trim('"');
                        }
                    }
                }
            }

            return lineFields;
        }

        public static List<string> GetAllScheduleNames(Document doc)
        {
            var _curDocScheduleNames = new List<string>();
            foreach (var vs in _GetSchedulesList(doc))
            {
                _curDocScheduleNames.Add(vs.Name); // Get all the Schedules in the current Document
            }
            return _curDocScheduleNames;
        }

        public static string _UpdateViewSchedule(Autodesk.Revit.DB.Document doc, string viewScheduleNameFromCSV, string[] headersFromCSV, List<string[]> viewScheduledataRows)
        {
            string _updatesResult = null;
            ViewSchedule _viewScheduleToUpdate = _GetViewScheduleByName(doc, viewScheduleNameFromCSV);
            if (_viewScheduleToUpdate != null)
            {
                _updatesResult += $"\n=== Updating ViewSchedule:{_viewScheduleToUpdate.Name} ===\n";

                List<Element> _rowsElementsOnViewSchedule =
                _GetElementsOnScheduleRow(doc, _viewScheduleToUpdate); // Get the list of Elements on Rows of the _viewScheduleToUpdate

                foreach (var _viewScheduledataRow in viewScheduledataRows) // Loop through the dataRows from viewScheduledataRows
                {
                    string _curCsvRowUniqueId = _viewScheduledataRow[0];  // Get the Unique Id from the current Row





                    Element _curRowElement =
                        _getElementOnViewScheduleRowByUniqueId(doc, _viewScheduleToUpdate, _curCsvRowUniqueId); // Get Element on ViewSchedule Row by UniqueId

                    // if the Element from the _curCsvRowUniqueId does not exist in the current schedule, skip it.
                    if (_curRowElement != null)
                    {
                        ParameterSet paramSet = _curRowElement.Parameters; // Get the parameters of the current row element in the _viewScheduleToUpdate

                        int headerCount = headersFromCSV.Count();
                        for (int i = 1; i < headerCount; i++)
                        {
                            int _headerColumnNumber = i;
                            string _curCsvHeaderName = headersFromCSV[_headerColumnNumber];
                            _curCsvHeaderName = _curCsvHeaderName.Trim();
                            Debug.Print(_curCsvHeaderName);

                            foreach (Parameter parameter in paramSet)
                            {
                                string paramName = null;
                                paramName = parameter.Definition.Name; // Get the name of the parameter
                                Debug.Print(paramName);

                                if (paramName == _curCsvHeaderName && parameter != null && parameter.StorageType == StorageType.String && !parameter.IsReadOnly)
                                {

                                    string _valueFromCsv = _viewScheduledataRow[_headerColumnNumber];
                                    parameter.Set(_valueFromCsv);
                                    //// Check if the CSV value is not empty before updating the parameter
                                    //if (!string.IsNullOrEmpty(_viewScheduledataRow[_headerColumnNumber]))
                                    //{
                                    //    string _valueFromCsv = _viewScheduledataRow[_headerColumnNumber];
                                    //    parameter.Set(_valueFromCsv);
                                    //}
                                    //else
                                    //{
                                    //    if (parameter.HasValue && parameter.IsShared)
                                    //        parameter.ClearValue();
                                    //}

                                }

                            }
                        }

                    }

                }
            }

            return _updatesResult;
        }
        //###########################################################################################

        public static void _MyTaskDialog(string Title, string MainContent)
        {
            TaskDialog _taskScheduleResult = new TaskDialog(Title); 
            _taskScheduleResult.TitleAutoPrefix = false;
            _taskScheduleResult.MainContent = MainContent;
            _taskScheduleResult.Show();
        }
        //###########################################################################################

        public static string _CreateFolderOnDesktopByName(string name)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderPath = System.IO.Path.Combine(desktopPath, name);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                Debug.Print("Directory created successfully.");
            }
            else
            {
                Debug.Print("Directory already exists.");
            }
            return folderPath;
        }

    }
}