#region Namespaces
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;

using static System.Net.Mime.MediaTypeNames;
#endregion

#region Begining of doc
namespace ORH_ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Import_Old : MyUtils, IExternalCommand
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
            #endregion

            #region Testing
            // This if statement is just to test some code without 
            // running the entire code.
            // Change to true or false depending if you want to execute or not.
            //if (false)
            //{
            //    string[] csvFiles = GetCsvFilePath();
            //    return Result.Cancelled;
            //}

            if (false) // <---- Don't for get to change it if you need to.
            {
                // ================= GetAllSchedules =================
                var _schedulesList1 = _GetSchedulesList(doc); // Get all the Schedules into a list
                // Uncommen the option you want to use
                //==== Option 1 ==== Get schedule by array possition in _schedulesList ;
                var uid = _schedulesList1[6].UniqueId;
                var vsName = _schedulesList1[6].Name;


                //string uid = "dc86627d-cf12-49fe-bdad-488a619b34a1-00060aca"; // row Elem UniqueId - NO GOOD
                //string uid = "7a2419bd-e042-4b38-8b95-781bc33e7dd8-000854c1"; // schedule ID - GOOD

                var elementsReturned = _GetViewScheduleBasedOnUniqueId(doc, uid);
                foreach (var _viewSchedule in elementsReturned)
                {
                    Debug.Print($"Elem Name: {_viewSchedule.Name} " +
                                $"UniqueId : {_viewSchedule.UniqueId} =========");


                    //#region get table - tesing
                    //var _td = _viewSchedule.GetTableData().GetSectionData(SectionType.Body);
                    //int Rows = _td.NumberOfRows;
                    //int Columns = _td.NumberOfColumns;
                    //string data = "";
                    //for (int i = 0; i < Rows; i++)
                    //{
                    //    for (int j = 0; j < Columns; j++)
                    //    {
                    //        data += _td.GetCellText(i, j) + ",";
                    //    }
                    //    // remove the trailing "," after the last cell and add a newline for the end of this row
                    //    data = data.TrimEnd(',') + "\n";

                    //}
                    //    Debug.Print($"{_viewSchedule.Name}, {data}");
                    //Debug.Print("wait");
                    //#endregion get table - tesing



                    ////// Create a ViewScheduleExportOptions object
                    //ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();
                    //string _path = @"C:\Users\ohernandez\Desktop\RevitAPI_Testing";
                    //string _name = "svExport.csv";
                    //_viewSchedule.Export(_path, _name, exportOptions); // exports schedule
                    //Process.Start(Path.Combine(_path, _name)); // opens exported file


                    //// Iterate through the fields
                    //foreach (Field field in fields)
                    //{
                    //    // Check if the field is a schedule field
                    //    if (field.IsScheduleField)
                    //    {
                    //        // Print the field name
                    //        Console.WriteLine(field.Name);
                    //    }
                    //}

                    var _familyInstances = _GetElementsOnScheduleRow(doc, _viewSchedule);
                    //var vsElem = doc.GetElement(parameter.Id);
                    foreach (var rowElem in _familyInstances)
                    {
                        var vsELP = rowElem.LookupParameter("Test2");
                        //var vsVal = vsELP.AsValueString();
                        //var vsVal2 = rowElem.Id;
                        var i = doc.GetElement(rowElem.Id);
                        var t = i.ParametersMap;
                        var s = t.get_Item("Test2");
                    }

                    foreach (FamilyInstance _familyInstance in _familyInstances)
                    {
                        var _params = _familyInstance.Parameters;


                        // Iterate through the parameters
                        foreach (Parameter parameter in _params)
                        {
                            if (_params != null)
                            {


                                var p0 = parameter.Definition.Name;
                                var p1 = parameter.AsValueString();
                                var p2 = parameter.Definition;
                                var p3 = parameter.Definition.ParameterType;
                                var p4 = p2.Name;
                                var p5 = parameter.Definition.ParameterGroup;

                                // Print the parameter name and value
                                Debug.Print($"{p0},{p1},{p2},{p3},{p4},{p5} =+=+=+==++==+====");
                            }

                        }

                    }

                }


                return Result.Cancelled;
            }
            #endregion

            // ================= Import CSVs =================
            string[] csvFilePaths = GetCsvFilePath(); // Get CSV file paths
            if (csvFilePaths == null)
            {
                TaskDialog.Show("INFO", "You didn't select any CSV file");
                return Result.Cancelled;
            }//Tell user no file was selected and stop process

            var _curDocScheduleNames = GetAllScheduleNames(doc); // Get all the schedules names in current doc

            string csvScheduleNamesFound = null;    // Tran Found Schedules
            string csvScheduleNamesNotFound = null; // Track Not Found Schedules
            foreach (var csvFilePath in csvFilePaths)  // Loop Through all the selected CSVs
            {
                var _viewScheduleNameFromCSV = M_GetLinesFromCSV(csvFilePath, 1)[0];  // Get View schedule name from csv
                if (_curDocScheduleNames.Contains(_viewScheduleNameFromCSV))
                {
                    Debug.Print($"Schedule: {_viewScheduleNameFromCSV} - Found in current document!");

                    var _headersFromCSV = M_GetLinesFromCSV(csvFilePath, 2);                   // Get Headers from csv
                    List<string[]> _viewScheduledata = ImportCSVToStringList(csvFilePath);  // Get data from csv - skips the first 3 lines
                    csvScheduleNamesFound += $"{_viewScheduleNameFromCSV}\n";               // add found schedule to csvScheduleNamesFound for later report.

                    using (Transaction tx = new Transaction(doc, $"Update {_viewScheduleNameFromCSV} Parameters")) // Start a new transaction to make changes to the elements in Revit
                    {
                        tx.Start(); // Lock the doc while changes are made in the transaction
                        var _viewScheduleUpdateResult = _UpdateViewSchedule(doc, _viewScheduleNameFromCSV, _headersFromCSV, _viewScheduledata);
                        tx.Commit();
                    }
                }
                else
                {
                    Debug.Print($"Schedule: {_viewScheduleNameFromCSV} - Not found in current document!");
                    csvScheduleNamesNotFound += $"{_viewScheduleNameFromCSV}\n"; // add found schedule to csvScheduleNamesNotFound for later report.
                }

            }
            if (csvScheduleNamesFound != null)
            {
                TaskDialog.Show("INFO", $"Found the following Schedule(s):\n\n{csvScheduleNamesFound}");
            }
            if (csvScheduleNamesNotFound != null)
            {
                TaskDialog.Show("INFO", $"Could Not find the following Schedule(s):\n\n{csvScheduleNamesNotFound}");
            }

            return Result.Succeeded;

            // ================= GetAllSchedules =================
            var _schedulesList = _GetSchedulesList(doc); // Get all the Schedules into a list
            //Get schedule by array possition in _schedulesList;
            var _selectedSchedule = _schedulesList[6];
            Debug.Print($"Selected Schedule Name: {_selectedSchedule.Name}");
            //return Result.Cancelled;

            // Get list of row elements
            List<Element> _elementsList = MyUtils._GetElementsOnScheduleRow(doc, _selectedSchedule);
            //List<ElementId> td = _elementsList.Select(e => e.Id).ToList();

            using (Transaction tx = new Transaction(doc, "Update Parameters")) // Start a new transaction to make changes to the elements in Revit
            {
                tx.Start(); // Lock the doc while changes are made in the transaction

                int numCount = 0; // Initialize a counter to keep track of the number of elements processed
                var rowElemIds = _elementsList.Select(e => e.Id);  // Get the IDs of all the row elements in the selected schedule
                foreach (var rowElemId in rowElemIds)
                {
                    var rowElem = doc.GetElement(rowElemId); // Get the element from its ID
                    Debug.Print($"{rowElem.Name} {rowElem.UniqueId}");


                    ParameterSet paramSet = rowElem.Parameters; // Get the parameters of the element
                    // Print the name and value of each parameter
                    foreach (Parameter parameter in paramSet)
                    {
                        string paramName = parameter.Definition.Name; // Get the name of the parameter
                        string paramValue = parameter.AsValueString(); // Get the value of the parameter as a string

                        if (paramName == "Annotation Description")
                        {
                            Debug.Print(parameter.Definition.Name + "==============================================");


                            Debug.Print($"Parameter_Name: {paramName} \n" +
                                       $"Parameter_Value: {paramValue} \n" +
                                            $"IsReadOnly: {parameter.IsReadOnly}\n" +
                                $"====================================");

                            numCount++;

                            // Update all the values from the Test2 Parameter to "Romeo"
                            //if (paramName == "Test2") // use this one if there is a parameter call Test2
                            if (paramName != null && parameter.IsReadOnly == false) // Only attempt to make shanges if the parameter Is Not ReadyOnly
                            {
                                // Set the value of the parameter 
                                var result = parameter.Set($"{numCount} - ElemID: {parameter.Id}");

                            }
                        }
                    }
                }

                tx.Commit(); // Commit the changes made in the transaction
            }

            return Result.Succeeded;
        }







        //public void ExportScheduleWithUniqueId(Autodesk.Revit.DB.Document doc, ViewSchedule viewSchedule)
        //{
        //    // Get the Unique ID parameter
        //    var uniqueIdParam = doc.GetElement(viewSchedule.UniqueId) as Parameter;
        //    if (uniqueIdParam == null || uniqueIdParam.StorageType != StorageType.String)
        //    {
        //        TaskDialog.Show("Error", "The Unique ID parameter is not available for this schedule.");
        //        return;
        //    }

        //    // Set up the CSV export options
        //    var csvExportOptions = new CsvExportOptions();
        //    csvExportOptions.SetSeparator(",");
        //    csvExportOptions.ExportUrls = false;

        //    // Create a memory stream to hold the exported data
        //    var memoryStream = new MemoryStream();

        //    // Export the schedule to CSV format
        //    viewSchedule.ExportTo(new ScheduleExportOptions(csvExportOptions), memoryStream);

        //    // Rewind the memory stream to the beginning
        //    memoryStream.Seek(0, SeekOrigin.Begin);

        //    // Create a stream reader to read the exported CSV data
        //    var streamReader = new StreamReader(memoryStream);

        //    // Create a memory stream to hold the modified CSV data
        //    var modifiedMemoryStream = new MemoryStream();

        //    // Create a stream writer to write the modified CSV data
        //    var streamWriter = new StreamWriter(modifiedMemoryStream);

        //    // Read the CSV data line by line and add the Unique ID column
        //    while (!streamReader.EndOfStream)
        //    {
        //        var line = streamReader.ReadLine();
        //        if (line == null) continue;

        //        // Split the line into fields
        //        var fields = line.Split(csvExportOptions.GetSeparator());

        //        // Get the element ID from the first field
        //        var elementIdString = fields[0];
        //        if (!int.TryParse(elementIdString, out var elementId)) continue;

        //        // Get the element by ID
        //        var element = viewSchedule.Document.GetElement(new ElementId(elementId));
        //        if (element == null) continue;

        //        // Get the Unique ID value for the element
        //        var uniqueId = uniqueIdParam.AsString();

        //        // Add the Unique ID field to the line
        //        var modifiedLine = $"{elementIdString},{uniqueId},{string.Join(",", fields, 1, fields.Length - 1)}";

        //        // Write the modified line to the output stream
        //        streamWriter.WriteLine(modifiedLine);
        //    }

        //    // Flush the stream writer to ensure that all data is written to the output stream
        //    streamWriter.Flush();

        //    // Rewind the modified memory stream to the beginning
        //    modifiedMemoryStream.Seek(0, SeekOrigin.Begin);

        //    // Save the modified CSV data to a file
        //    var filePath = @"C:\Users\ohernandez\Desktop\RevitAPI_Testing\svExport.csv";
        //    using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        //    {
        //        modifiedMemoryStream.CopyTo(fileStream);
        //    }

        //    // Open the exported file
        //    System.Diagnostics.Process.Start(filePath);
        //}



        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }

    }
}