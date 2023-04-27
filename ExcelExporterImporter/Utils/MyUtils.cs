//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using System.Collections.Generic;
//using System.Diagnostics;
//using System.Text;
//using System.Windows.Controls;
//using System.Windows.Shapes;

//namespace ExcelExporterImporter
//{
//    [Transaction(TransactionMode.Manual)]
//    public class MyUtils
//    {
//        //###########################################################################################
//        /// <summary>
//        /// _GetScheduleTableDataAsString Method: Takes in one ViewSchedule and the Autodesk.Revit.DB.Document
//        /// </summary>
//        /// <param name="_selectedSchedule"></param>
//        /// <param name="doc"></param>
//        /// <returns>Returns a comma delimited string list of all the schedule data</returns>
//        /// 
//        public List<string> _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule, Autodesk.Revit.DB.Document doc)
//        {
//            //---Gets the list of items/rows in the schedule---
//            var _scheduleItemsCollector = new FilteredElementCollector(doc, _selectedSchedule.Id).WhereElementIsNotElementType();
//            List<string> uid = new List<string>(); // List to hold all the UniqueIDs in the schedule
//            foreach (var _scheduleRow in _scheduleItemsCollector)
//            {
//                Debug.Print($"{_scheduleRow.GetType()} - UID: {_scheduleRow.UniqueId}"); // only for Debug
//                uid.Add(_scheduleRow.UniqueId); // adds each UniqueID to the uid list
//            }

//            // Get the data from the ViewSchedule object
//            TableData tableData = _selectedSchedule.GetTableData();
//            TableSectionData _sectionData = tableData.GetSectionData(SectionType.Body);

//            // Concatenate the data into a list of strings, with each string representing a row
//            List<string> _rowsList = new List<string>();
//            for (int row = 0; row < _sectionData.NumberOfRows; row++)
//            {
//                int _adjestedRow = row - 2;
//                StringBuilder sb = new StringBuilder();
//                if (row == 0)
//                {
//                    sb.Append("UniqueId,"); // adds the "UniqueId" header
//                }
//                if (row == 1)
//                {
//                    sb.Append(","); // adds the "," on the second line. This is because schedules second row is empty.
//                }
//                for (int col = 0; col < _sectionData.NumberOfColumns; col++)
//                {
//                    string cellText = _sectionData.GetCellText(row, col).ToString();
//                    var cellTextColIndex = _sectionData.GetCellText(row, col);

//                    if (col == 0 && row >= 2) // if first column, add the UniqueID
//                    {
//                        string uniqueId = uid[_adjestedRow];
//                        sb.Append(uniqueId + ",");
//                    }
//                    else if (IsScheduleFieldEditable(_selectedSchedule, col))
//                    {
//                        sb.Append("{");
//                        sb.Append(cellText);
//                        sb.Append("}");
//                    }
//                    else
//                    {
//                        sb.Append(cellText);
//                    }
//                    sb.Append(",");
//                }
//                _rowsList.Add(sb.ToString());
//            }
//            // This line is only for debug, outputs all the rows in _scheduleTableDataAsString
//            int lineN = 0; foreach (var line in _rowsList) { lineN++; Debug.Print($"Row {lineN}: {line}"); }

//            return _rowsList;
//        }
//        //###########################################################################################


//        public bool IsScheduleFieldEditable(ViewSchedule viewSchedule, int columnIndex)
//        {
//            var td = viewSchedule.GetTableData();

//            var sd= td.GetSectionData(columnIndex);
//            Debug.Print(sd.GetCellText(3,3));
//            //// Get the column IDs
//            //int[] columnIds = viewSchedule.GetColumnIds(SectionType.Body);

//            //// Get the column data
//            //object[] columnData = viewSchedule.GetColumnData(columnIds);

//            //// Check if the column is editable
//            //if (columnData[columnIndex] != null && columnData[columnIndex] is bool)
//            //{
//            //    return (bool)columnData[columnIndex];
//            //}
//            //else
//            //{
//               return false;
//            //}
//        }

//    }
//}
///#########
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;

namespace ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class MyUtils
    {
        //###########################################################################################

        public List<ViewSchedule> _GetSchedulesList(Autodesk.Revit.DB.Document doc) // This method returns a list of all the schedule elements
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
        public List<string> _GetScheduleTableDataAsString(ViewSchedule _selectedSchedule, Autodesk.Revit.DB.Document doc)
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
                if(viewSchedule.UniqueId == uniqueId && viewSchedule != null)
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
        protected static List<Element> GetElementsOnScheduleRow(Document doc, ViewSchedule selectedSchedule)
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



    }
}