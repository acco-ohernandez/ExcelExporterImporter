using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExporterImporter.Utils
{
    internal class TestCode
    {
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

}
