using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Lifetime;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;

using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;

using OfficeOpenXml;

namespace ORH_ExcelExporterImporter
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
        public static ViewSchedule M_GetViewScheduleByName(Document doc, string viewScheduleName)
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
        public static ViewSchedule M_GetViewScheduleByUniqueId(Document doc, string viewScheduleUniqueId)
        {
            FilteredElementCollector _schedules = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
            ViewSchedule _viewScheduleNotFound = null;
            foreach (ViewSchedule curViewScheduleInDoc in _schedules)
            {
                if (curViewScheduleInDoc.UniqueId == viewScheduleUniqueId)
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
        public static List<string[]> ImportCSVToStringList2(string csvFilePath)
        {
            var dataList = new List<string[]>();

            // Create a TextFieldParser to parse the CSV file
            using (TextFieldParser parser = new TextFieldParser(csvFilePath)) // Requires the Microsoft.VisualBasic namespace reference and using Microsoft.VisualBasic.FileIO;
            {
                // Set the delimiter to comma
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                // Set the option to handle fields enclosed in quotes
                parser.HasFieldsEnclosedInQuotes = true;

                // Skip the first two lines
                if (!parser.EndOfData)
                {
                    parser.ReadLine();
                }

                if (!parser.EndOfData)
                {
                    parser.ReadLine();
                }

                // Read each line of the CSV file
                while (!parser.EndOfData)
                {
                    // Read the fields of the current line
                    string[] fields = parser.ReadFields();

                    // Add the fields to the dataList collection
                    dataList.Add(fields);
                }
            }

            // Return the parsed data as a list of string arrays
            return dataList;
        }

        //###########################################################################################


        public static string[] M_GetLinesFromCSV(string csvFilePath, int lineNumber)
        {
            string[] lineFields = null;

            try
            {
                using (StreamReader reader = new StreamReader(csvFilePath))
                {
                    // Read lines until we reach the specified line number
                    for (int i = 1; i <= lineNumber; i++)
                    {
                        string line = reader.ReadLine();

                        if (line == null)
                        {
                            // Reached the end of the file before reaching the specified line number
                            Debug.Print("Specified line number is out of range.");
                            return null;
                        }

                        if (i == lineNumber)
                        {
                            lineFields = line.Split(',');

                            // Remove quotes from each field
                            for (int j = 0; j < lineFields.Length; j++)
                            {
                                //string skipDoubleQuoteinText = SkipDoubleQuoteInText(lineFields[j]);
                                lineFields[j] = lineFields[j].Trim('"');
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.Print($"An error occurred while reading the CSV file: \n{csvFilePath}");
                Debug.Print(ex.ToString());
                lineFields = null; // Set lineFields to null in case of an error
            }

            return lineFields;
        }

        //###########################################################################################

        public static List<string> GetAllScheduleNames(Document doc)
        {
            var _curDocScheduleNames = new List<string>();
            foreach (var vs in _GetSchedulesList(doc))
            {
                _curDocScheduleNames.Add(vs.Name); // Get all the Schedules in the current Document
            }
            return _curDocScheduleNames;
        }
        public static List<string> GetAllScheduleUniqueIds(Document doc)
        {
            var _curDocScheduleUniqueIds = new List<string>();
            foreach (var vs in _GetSchedulesList(doc))
            {
                _curDocScheduleUniqueIds.Add($"{vs.UniqueId}"); // Get all the Schedules UniqueIds in the current Document
            }
            return _curDocScheduleUniqueIds;
        }
        //###########################################################################################
        public static string _UpdateViewSchedule(Autodesk.Revit.DB.Document doc, string viewScheduleUniqueIdFromCSV, string[] headersFromCSV, List<string[]> viewScheduledataRows)
        {
            string _updatesResult = null;
            ViewSchedule _viewScheduleToUpdate = M_GetViewScheduleByUniqueId(doc, viewScheduleUniqueIdFromCSV);
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


                                if (parameter.IsShared)// test
                                {
                                    var paramGuid = parameter.GUID; // Get the name of the parameter
                                }

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
        public static string _UpdateViewSchedule_KindaWorking(Autodesk.Revit.DB.Document doc, string viewScheduleUniqueIdFromCSV, string[] headersFromCSV, List<string[]> viewScheduledataRows)
        {
            string _updatesResult = null;
            ViewSchedule _viewScheduleToUpdate = M_GetViewScheduleByUniqueId(doc, viewScheduleUniqueIdFromCSV);
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

                                if (parameter.IsShared)// test
                                {
                                    var paramGuid = parameter.GUID; // Get the name of the parameter
                                }

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

        public static void M_MyTaskDialog(string Title, string MainContent)
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

        //###########################################################################################
        //add shared parameter to schedule
        public static void _AddSharedParameterToSchedule(Document doc, string _viewScheduleName)
        {
            var viewSchedule = M_GetViewScheduleByName(doc, _viewScheduleName);
            var field = viewSchedule.Definition.GetSchedulableFields().FirstOrDefault(x => IsSharedParameterSchedulableField(viewSchedule.Document, x.ParameterId, new Guid("<your guid>")));
            //.FirstOrDefault(x => x.ParameterId == destParameterId);


        }
        public static bool IsSharedParameterSchedulableField(Document document, ElementId parameterId, Guid sharedParameterId)
        {
            var sharedParameterElement = document.GetElement(parameterId) as SharedParameterElement;

            return sharedParameterElement?.GuidValue == sharedParameterId;
        }


        public static ScheduleField GetScheduleFieldByName(Document doc, ViewSchedule viewSchedule, string fieldName)
        {
            ScheduleDefinition definition = viewSchedule.Definition;
            var scheduleFieldIds = definition.GetFieldOrder();

            // Loop through the schedule fields and find the one with matching name
            foreach (ScheduleFieldId fieldId in scheduleFieldIds)
            {
                ScheduleField field = definition.GetField(fieldId); ;
                var ParamId = field.ParameterId;
                if (ParamId.ToString() == fieldName)
                {
                    return field as ScheduleField;
                }
            }

            // If no matching field is found, return null
            return null;
        }



        //public static ScheduleField M_AddByNameAvailableFieldToSchedule(Document doc, string scheduleName, string fieldName)
        public static void M_AddByNameAvailableFieldToSchedule(Document doc, string scheduleName, string fieldName)
        {
            // Get the schedule by name.
            ViewSchedule schedule = M_GetViewScheduleByName(doc, scheduleName);

            // Get the definition of the schedule.
            ScheduleDefinition definition = schedule.Definition;

            // Get the list of available fields.
            IList<SchedulableField> availableFields = definition.GetSchedulableFields();

            // Find the field you want to add.
            SchedulableField field = availableFields.FirstOrDefault(f => f.GetName(doc) == fieldName);

            // Add the field to the schedule.
            ScheduleField fieldAdded = null;
            try
            {
                fieldAdded = definition.AddField(field);
            }
            catch
            {
                Debug.Print("====    Did not add the dev_Text_1 field because it already existed.    ====");

                //var fields =  GetScheduleFieldByName(doc, schedule, "dev_Text_1");
                // fieldAdded = definition.GetField(fields.FieldId);
            }


            //return fieldAdded;
        }
        //###########################################################################################
        public static void _UpdateMyUniqueIDColumn(Document doc, string _viewScheduleName)
        {
            ViewSchedule _viewScheduleToUpdate = M_GetViewScheduleByName(doc, _viewScheduleName); // Get the schedule by name
            List<Element> _rowsElementsOnViewSchedule = _GetElementsOnScheduleRow(doc, _viewScheduleToUpdate); // Get the list of Elements on Rows of the _viewScheduleToUpdate
            foreach (Element _rowElement in _rowsElementsOnViewSchedule)
            {
                // Get the parameter set for the current row element.
                ParameterSet paramSet = _rowElement.Parameters; // Get the parameters of the current row element in the _viewScheduleToUpdate

                // paramSet.
                //paramSet.Where(p => p.Definition.Name == "MyUniqueId");


                // Iterate through the parameters in the parameter set.
                foreach (Parameter param in paramSet)
                {
                    // Check if the parameter's name is equal to MyUniqueId.
                    //if (param.Definition.Name == "MyUniqueId")
                    if (param.Definition.Name == "Dev_Text_1")
                    {
                        // Get the parameter's value.
                        param.Set($"{_rowElement.UniqueId}");
                        break;
                    }
                }

            }


            //TaskDialog.Show("Info", "Added UniqueIDs to MyUniqueId Column");
            Debug.Print("Added UniqueIDs to MyUniqueId Column");
        }
        //###########################################################################################
        private void ShowDefinitionFileInfo(DefinitionFile myDefinitionFile)
        {
            StringBuilder fileInformation = new StringBuilder(500);

            // get the file name 
            fileInformation.AppendLine("File Name: " + myDefinitionFile.Filename);

            // iterate the Definition groups of this file
            foreach (DefinitionGroup myGroup in myDefinitionFile.Groups)
            {
                // get the group name
                fileInformation.AppendLine("Group Name: " + myGroup.Name);

                // iterate the difinitions
                foreach (Definition definition in myGroup.Definitions)
                {
                    // get definition name
                    fileInformation.AppendLine("Definition Name: " + definition.Name);
                }
            }
            TaskDialog.Show("Revit", fileInformation.ToString());
        }
        //###########################################################################################

        public static void AddNewParameterToSchedule(Document doc, string _viewScheduleName, string parameterName)
        {
            // Get the schedule by name.
            ViewSchedule schedule = M_GetViewScheduleByName(doc, _viewScheduleName); // Get the schedule by name

            // Get the definition of the schedule.
            ScheduleDefinition definition = schedule.Definition;

            // Create a new parameter.



            // Update the schedule.
            // schedule.Update();
        }
        public static void M_GetSharedParameterFile(Autodesk.Revit.ApplicationServices.Application app)
        {
            var originalSharedParametersFilename = app.SharedParametersFilename;
            // Create a new Revit application object.
            //Application app = new Application();

            // Set the SharedParametersFilename property to the path of the shared parameters file.
            //app.SharedParametersFilename = @"C:\Users\ohernandez\Desktop\Revit_Exports\SharedParams\ACCO -- Dev_Revit Shared Parameters.txt";
            app.SharedParametersFilename = M_CreateSharedParametersFile();

            // Open the shared parameters file.
            DefinitionFile definitionFile = app.OpenSharedParameterFile();

            // Get the DefinitionGroups collection for the DefinitionFile object.
            DefinitionGroups groups = definitionFile.Groups;

            // Iterate through the DefinitionGroups collection to get the DefinitionGroup objects.
            foreach (DefinitionGroup group in groups)
            {
                // Iterate through the DefinitionGroup objects to get the Definition objects.
                foreach (Definition definition in group.Definitions)
                {
                    // Use the Definition objects to access the shared parameters.
                    Console.WriteLine(definition.Name);
                }
            }

            // Close the shared parameters file.
            definitionFile.Dispose(); //Close();

            // Dispose the Revit application object.
            app.Dispose();

            app.SharedParametersFilename = originalSharedParametersFilename;
        }

        public static Definition GetParameterDefinitionFromFile(DefinitionFile defFile, string groupName, string paramName)
        {
            // iterate the Definition groups of this file
            foreach (DefinitionGroup group in defFile.Groups)
            {
                if (group.Name == groupName)
                {
                    // iterate the difinitions
                    foreach (Definition definition in group.Definitions)
                    {
                        if (definition.Name == paramName)
                            return definition;
                    }
                }
            }
            return null;
        }

        public static string M_MyAddNewParameterToSchedule(Autodesk.Revit.ApplicationServices.Application app)
        {
            string originalSharedParametersFile = app.SharedParametersFilename;

            // Set the SharedParametersFilename property to the path of the shared parameters file.
            //app.SharedParametersFilename = @"C:\Users\ohernandez\Desktop\Revit_Exports\SharedParams\ACCO -- Dev_Revit Shared Parameters.txt";
            app.SharedParametersFilename = M_CreateSharedParametersFile();

            // Open the shared parameters file.
            DefinitionFile definitionFile = app.OpenSharedParameterFile();
            Definition paramDef = null;
            // iterate the Definition groups of this file
            foreach (DefinitionGroup group in definitionFile.Groups)
            {
                if (group.Name == "Dev_Group_Common")
                {
                    // iterate the difinitions
                    foreach (Definition definition in group.Definitions)
                    {
                        if (definition.Name == "Dev_Text_1")
                            paramDef = definition;
                    }
                }
            }

            return originalSharedParametersFile;
        }


        public static void M_Add_Dev_Text_4(Autodesk.Revit.ApplicationServices.Application app, Document doc, ViewSchedule _curSchedule, CategorySet myCatSet)
        {
            string originalSharedParametersFilename = null;
            try
            {
                // Define category for shared parameter
                // ...

                // Save the original shared parameters filename
                originalSharedParametersFilename = app.SharedParametersFilename;

                // Set the path of the shared parameters file
                app.SharedParametersFilename = M_CreateSharedParametersFile();

                // Open the shared parameters file
                DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();

                // Get the parameter definition from the shared parameters file
                var curDef = MyUtils.GetParameterDefinitionFromFile(sharedParameterFile, "Dev_Group_Common", "Dev_Text_1");

                // Create the binding
                ElementBinding curBinding = doc.Application.Create.NewInstanceBinding(myCatSet);

                // Insert the parameter into the document
                var paramAdded = doc.ParameterBindings.Insert(curDef, curBinding, BuiltInParameterGroup.PG_IDENTITY_DATA);

                // Add the available field to the schedule
                M_AddByNameAvailableFieldToSchedule(doc, _curSchedule.Name, "Dev_Text_1");

                // Update the "Dev_Text_1" column with the UniqueID 
                _UpdateMyUniqueIDColumn(doc, _curSchedule.Name);

                // Restore the original shared parameters filename
                app.SharedParametersFilename = originalSharedParametersFilename;

                // Regenerate the document
                _curSchedule.Document.Regenerate();
            }
            catch (Exception ex)
            {
                app.SharedParametersFilename = originalSharedParametersFilename;

                // Handle the exception
                Debug.Print("Error occurred in M_Add_Dev_Text_4: " + ex.Message);
            }
        }


        public static void M_Add_Dev_Text_2(Autodesk.Revit.ApplicationServices.Application app, Document doc, ViewSchedule _curSchedule, BuiltInCategory _builtInCat)
        {
            //define category for shared param
            //Category myCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment);
            Category myCat = doc.Settings.Categories.get_Item(_builtInCat);
            CategorySet myCatSet = doc.Application.Create.NewCategorySet();
            myCatSet.Insert(myCat);

            //app.SharedParametersFilename = @"Y:\\DATABASES\\ACCORevit\\02-SHARED PARAMETERS\\ACCO -- Revit Shared Parameters.txt";
            var originalSharedParametersFilename = app.SharedParametersFilename;

            // Set the SharedParametersFilename property to the path of the shared parameters file.
            //app.SharedParametersFilename = @"C:\Users\ohernandez\Desktop\Revit_Exports\SharedParams\ACCO -- Dev_Revit Shared Parameters.txt";
            app.SharedParametersFilename = M_CreateSharedParametersFile();

            // Open the shared parameters file.
            DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();
            var curDef = MyUtils.GetParameterDefinitionFromFile(sharedParameterFile, "Dev_Group_Common", "Dev_Text_1");
            //create binding
            ElementBinding curBinding = doc.Application.Create.NewInstanceBinding(myCatSet);

            var paramAdded = doc.ParameterBindings.Insert(curDef, curBinding, BuiltInParameterGroup.PG_IDENTITY_DATA);

            //var af =
            M_AddByNameAvailableFieldToSchedule(doc, _curSchedule.Name, "Dev_Text_1");
            _UpdateMyUniqueIDColumn(doc, _curSchedule.Name);

            app.SharedParametersFilename = originalSharedParametersFilename;

            _curSchedule.Document.Regenerate();
        }
        public static BuiltInCategory _GetScheduleBuiltInCategory(ViewSchedule schedule)
        {
            Category scheduleCategory = schedule.Category;
            BuiltInCategory builtInCategory = (BuiltInCategory)scheduleCategory.Id.IntegerValue;

            if (Enum.IsDefined(typeof(BuiltInCategory), builtInCategory))
            {
                return builtInCategory;
            }
            else
            {
                // Return a default value or throw an exception, based on your requirements
                return BuiltInCategory.INVALID;
            }
        }


        public static BuiltInCategory _GetElementBuiltInCategory(Element element)
        {
            Document doc = element.Document;
            Category category = doc.GetElement(element.GetTypeId()).Category;
            BuiltInCategory builtInCategory = (BuiltInCategory)category.Id.IntegerValue;

            if (Enum.IsDefined(typeof(BuiltInCategory), builtInCategory))
            {
                return builtInCategory;
            }
            else
            {
                // Return a default value or throw an exception, based on your requirements
                return BuiltInCategory.INVALID;
            }
        }

        public static Category _GetScheduleCategory(Document doc, ViewSchedule schedule)
        {
            var CatID = schedule.Definition.CategoryId;
            Category _scheduleCategory = null;
            foreach (Category c in doc.Settings.Categories)
            {
                if (c.Id == CatID)
                {
                    Debug.Print($"CategoryName:{c.Name} ID:{c.Id} CategoryType:{c.CategoryType} ");
                    _scheduleCategory = c;
                    return _scheduleCategory;
                }
            }
            return null;
        }


        public List<BuiltInCategory> _GetAllBuiltInCategories()
        {
            List<BuiltInCategory> builtInCategories = new List<BuiltInCategory>();

            foreach (BuiltInCategory category in Enum.GetValues(typeof(BuiltInCategory)))
            {
                if (category != BuiltInCategory.INVALID)
                {
                    builtInCategories.Add(category);
                }
            }

            return builtInCategories;
        }


        public static BuiltInCategory _GetBuiltInCategoryFromCategory(Category category)
        {
            if (category != null && category.Id != null && category.Id.IntegerValue >= 0)
            {
                BuiltInCategory builtInCategory = (BuiltInCategory)category.Id.IntegerValue;
                if (Enum.IsDefined(typeof(BuiltInCategory), builtInCategory))
                {
                    return builtInCategory;
                }
            }

            // Return a default value or throw an exception, based on your requirements
            return BuiltInCategory.INVALID;
        }

        public static BuiltInCategory _GetBuiltInCategoryById(int categoryId)
        {
            BuiltInCategory builtInCategory = (BuiltInCategory)categoryId;

            if (Enum.IsDefined(typeof(BuiltInCategory), builtInCategory))
            {
                return builtInCategory;
            }
            else
            {
                // Return a default value or throw an exception, based on your requirements
                return BuiltInCategory.INVALID;
            }
        }

        //public void AddUniqueIdColumnToCsv(string filePath, string[] uniqueIds)
        public static void AddUniqueIdColumnToViewScheduleCsv(string filePath, List<string> uniqueIds)
        {
            // Wait for 1 second
            Thread.Sleep(200);

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("The specified file does not exist.", filePath);
            }

            var csvLines = File.ReadAllLines(filePath);
            if (csvLines.Length < 3)
            {
                throw new InvalidOperationException("The specified file does not have the correct format of rows.");
            }

            // Add the UniqueID header to the first row
            csvLines[0] = csvLines[0] + ",";       // Update Row 0
            csvLines[1] = "UniqueID," + csvLines[1];  // Update Row 1
            csvLines[2] = csvLines[2] + ",";        // Update Row 2

            // Add the UniqueID values to each subsequent row
            for (int i = 3; i < csvLines.Length; i++)
            {
                csvLines[i] = uniqueIds[i - 3] + "," + csvLines[i];
            }

            // Write the modified CSV data to the same file
            File.WriteAllLines(filePath, csvLines, Encoding.UTF8);
        }

        //###########################################################################################
        public static CategorySet M_GetAllowBoundParamCategorySet(Document doc, ViewSchedule schedule)
        {
            CategorySet newCatSet = new CategorySet();
            foreach (Category _settingsCategory in doc.Settings.Categories)
            {
                if (_settingsCategory.AllowsBoundParameters)
                {
                    newCatSet.Insert(_settingsCategory);
                }
            }

            return newCatSet;
        }
        public static BuiltInCategory M_GetScheduleBuiltInCategory(Document doc, ViewSchedule schedule)
        {
            ElementId _scheduleDefinitionCategoryId = schedule.Definition.CategoryId;

            BuiltInCategory _scheduleBuiltInCategory = new BuiltInCategory();
            foreach (Category _settingsCategory in doc.Settings.Categories)
            {
                var curC = _GetBuiltInCategoryFromCategory(_settingsCategory);
                //if (c.Id.IntegerValue == cId.IntegerValue)
                if (_settingsCategory.Id == _scheduleDefinitionCategoryId)
                {
                    _scheduleBuiltInCategory = _GetBuiltInCategoryById(_settingsCategory.Id.IntegerValue);
                    break;
                }
            }

            return _scheduleBuiltInCategory;
        }
        //###########################################################################################
        public static void M_MoveCsvLastColumnToFirst(string filePath)
        {
            // Read all lines from the CSV file
            string[] lines = File.ReadAllLines(filePath);
            //string[] lines = File.ReadAllLines(@"C:\Users\ohernandez\Desktop\Revit_Exports\Mechanical Equipment Schedule.csv");

            // Get the header line
            string headerLine = lines[0];

            // Get the data lines excluding the header line
            string[] dataLines = lines.Skip(1).ToArray();

            // Split the header line and data lines by comma
            string[] headerColumns = headerLine.Split(',');
            string[][] dataColumns = dataLines.Select(line => line.Split(',')).ToArray();

            // Move the last column to column A
            for (int i = 0; i < dataColumns.Length; i++)
            {
                string lastColumn = dataColumns[i][dataColumns[i].Length - 1];

                for (int j = dataColumns[i].Length - 1; j > 0; j--)
                {
                    dataColumns[i][j] = dataColumns[i][j - 1];
                }

                dataColumns[i][0] = lastColumn;
            }

            // Merge the updated header and data columns
            string[] updatedLines = new string[dataColumns.Length + 1];
            updatedLines[0] = string.Join(",", headerColumns);

            for (int i = 0; i < dataColumns.Length; i++)
            {
                updatedLines[i + 1] = string.Join(",", dataColumns[i]);
            }

            // Write the updated lines back to the file
            File.WriteAllLines(filePath, updatedLines, Encoding.UTF8);

        }



        public static void M_MoveCsvLastColumnToFirst_temp(string filePath)
        {
            // Read all lines from the CSV file
            string[] lines = File.ReadAllLines(filePath);

            // Get the header line
            string headerLine = lines[0];

            // Get the data lines excluding the header line
            string[] dataLines = lines.Skip(1).ToArray();

            // Split the header line and data lines by comma
            string[] headerColumns = headerLine.Split(',');

            // Split each data line into columns
            string[][] dataColumns = dataLines.Select(line => SplitLine(line)).ToArray();

            // Move the last column to the first column
            for (int i = 0; i < dataColumns.Length; i++)
            {
                string lastColumn = dataColumns[i][dataColumns[i].Length - 1];

                for (int j = dataColumns[i].Length - 1; j > 0; j--)
                {
                    dataColumns[i][j] = dataColumns[i][j - 1];
                }

                dataColumns[i][0] = lastColumn;
            }

            // Merge the updated header and data columns
            string[] updatedLines = new string[dataColumns.Length + 1];
            updatedLines[0] = string.Join(",", EscapeColumns(headerColumns));

            for (int i = 0; i < dataColumns.Length; i++)
            {
                updatedLines[i + 1] = string.Join(",", EscapeColumns(dataColumns[i]));
            }

            // Write the updated lines back to the file
            File.WriteAllLines(filePath, updatedLines, Encoding.UTF8);
        }
        private static string[] SplitLine(string line)
        {
            var columns = new List<string>();
            StringBuilder columnBuilder = new StringBuilder();
            bool inQuotes = false;

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    columns.Add(columnBuilder.ToString().Trim().Trim('"'));
                    columnBuilder.Clear();
                }
                else
                {
                    columnBuilder.Append(c);
                }
            }

            columns.Add(columnBuilder.ToString().Trim().Trim('"'));
            return columns.ToArray();
        }

        private static string[] SplitLine_Old(string line)
        {
            // Custom implementation to split a CSV line while handling quotes and special characters
            // This implementation assumes the CSV follows the standard CSV format

            var columns = new List<string>();
            StringBuilder columnBuilder = new StringBuilder();
            bool inQuotes = false;

            foreach (char c in line)
            {
                if (c == '"')
                {
                    // Toggle inQuotes flag when encountering a quote
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    // Add column to the list when encountering a comma outside quotes
                    columns.Add(columnBuilder.ToString().Trim());
                    columnBuilder.Clear();
                }
                else
                {
                    // Append character to the current column
                    columnBuilder.Append(c);
                }
            }

            // Add the last column to the list
            columns.Add(columnBuilder.ToString().Trim());

            return columns.ToArray();
        }

        private static string[] EscapeColumns(string[] columns)
        {
            // Custom implementation to escape quotes and special characters in CSV columns
            // Modify this method according to your specific escaping requirements

            for (int i = 0; i < columns.Length; i++)
            {
                // Escape double quotes by doubling them
                columns[i] = columns[i].Replace("\"", "\"\"");

                // Add additional escaping logic for special characters if needed
                // Modify the escaping rules based on your specific use case
                // For example, you may need to escape newlines, tabs, or specific symbols
            }

            return columns;
        }

        //###########################################################################################
        public void AddScheduleUniqueIdToA1_Old(string filePath, string newText)
        {
            // Read the existing CSV file
            string[] lines = File.ReadAllLines(filePath);

            // Move the existing text from A1 to B1 and shift other cells to the right
            for (int i = lines.Length - 1; i >= 0; i--)
            {
                string[] cells = lines[i].Split(',');

                // Shift cells to the right
                for (int j = cells.Length - 1; j >= 0; j--)
                {
                    if (j > 0 && j < cells.Length)
                    {
                        cells[j] = cells[j - 1];
                    }
                }

                // Update A1 with the new text
                if (i == 0)
                {
                    cells[0] = $"\"{newText}\"";
                }

                // Join cells back into a line
                lines[i] = string.Join(",", cells);
            }

            // Write the modified lines back to the CSV file
            File.WriteAllLines(filePath, lines, Encoding.UTF8);
        }
        public void M_AddScheduleUniqueIdToCellA1(string filePath, string newText)
        {
            // Read the existing CSV file
            string[] lines = File.ReadAllLines(filePath);

            // Move the existing text from A1 to B1 and insert new text in A1
            for (int i = 0; i < lines.Length; i++)
            {
                string[] cells = lines[i].Split(',');

                // Move the existing value from A1 to B1
                if (i == 0 && cells.Length > 1)
                {
                    cells[1] = cells[0];
                    // Insert the new text in A1
                    cells[0] = $"\"{newText}\"";
                }


                // Join cells back into a line
                lines[i] = string.Join(",", cells);
            }

            // Write the modified lines back to the CSV file
            File.WriteAllLines(filePath, lines, Encoding.UTF8);
        }

        //###########################################################################################
        /// <summary>
        /// This updated version takes into account quoted text fields by parsing each line character by character and keeping track of whether the current character is within quotes or not. It correctly handles commas within quoted text fields and ensures that the fields are split and joined correctly when moving the last column to the first position.

        ///Please note that this approach assumes that quoted text fields do not contain any escaped quotes.If your CSV file contains escaped quotes (e.g., "This is a ""quoted"" field"), additional handling may be required.
        /// </summary>
        /// <param name="csvFilePath"></param>
        public static void M_MoveCsvLastColumnToFirst2(string csvFilePath)
        {
            // Read all lines from the CSV file
            string[] lines = System.IO.File.ReadAllLines(csvFilePath);

            if (lines.Length > 0)
            {
                // Split the first line to get the column headers
                string[] headers = lines[0].Split(',');

                // Find the index of the last column
                int lastColumnIndex = headers.Length - 1;

                // Move the last column to the first position
                string lastColumnHeader = headers[lastColumnIndex];
                for (int i = lastColumnIndex; i > 0; i--)
                {
                    headers[i] = headers[i - 1];
                }
                headers[0] = lastColumnHeader;

                // Modify each line by moving the last column to the first position
                for (int i = 1; i < lines.Length; i++) // Start from index 1 to skip the header row
                {
                    List<string> values = new List<string>();
                    StringBuilder fieldValue = new StringBuilder();
                    bool withinQuotes = false;

                    foreach (char c in lines[i])
                    {
                        if (c == '"' && !withinQuotes)
                        {
                            withinQuotes = true;
                            fieldValue.Append(c);
                        }
                        else if (c == '"' && withinQuotes)
                        {
                            withinQuotes = false;
                            fieldValue.Append(c);
                        }
                        else if (c == ',' && !withinQuotes)
                        {
                            values.Add(fieldValue.ToString());
                            fieldValue.Clear();
                        }
                        else
                        {
                            fieldValue.Append(c);
                        }
                    }

                    values.Add(fieldValue.ToString());

                    string lastValue = values[lastColumnIndex];
                    for (int j = lastColumnIndex; j > 0; j--)
                    {
                        values[j] = values[j - 1];
                    }
                    values[0] = lastValue;

                    lines[i] = string.Join(",", values);
                }

                // Write the modified lines back to the CSV file
                System.IO.File.WriteAllLines(csvFilePath, lines, Encoding.UTF8);
            }
        }
        //##############################################################

        public static string M_CreateSharedParametersFile()
        {
            string data = @"# This is a Revit shared parameter file.
# Do not edit manually.
*META	VERSION	MINVERSION
META	2	1
*GROUP	ID	NAME
GROUP	1	Dev_Group_Common
*PARAM	GUID	NAME	DATATYPE	DATACATEGORY	GROUP	VISIBLE	DESCRIPTION	USERMODIFIABLE	HIDEWHENNOVALUE
PARAM	31fa72f6-6cd4-4ea8-9998-8923afa881e3	Dev_Text_1	TEXT		1	1		1	0";

            //33
            string outputFileName = @"ACCO -- Dev_Revit Shared Parameters.txt";
            string tempDirectory = System.IO.Path.GetTempPath();
            string tempFilePath = System.IO.Path.Combine(tempDirectory, outputFileName);

            try
            {
                // Write the data to the temp file
                System.IO.File.WriteAllText(tempFilePath, data);
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during file writing
                Console.WriteLine("An error occurred while writing the temp file: " + ex.Message);
                return null;
            }

            return tempFilePath;
        }

        public static ScheduleDefinition M_ShowHeadersAndTileOnSchedule(ViewSchedule curViewSchedule)
        {
            // get the current definitions of the schedule
            ScheduleDefinition curScheduleDefinition = curViewSchedule.Definition;

            // set the ShowTitle to True
            curScheduleDefinition.ShowTitle = true;

            // set the ShowHeaders to True
            curScheduleDefinition.ShowHeaders = true;

            // Return the original definitions
            return curScheduleDefinition;
        }

        public void CheckAndPromptToCloseExcel(string filePath)
        {
            FileStream stream = null;

            string fileName = System.IO.Path.GetFileName(filePath);

            try
            {
                Debug.Print($"Checking if file: {fileName} is currently locked by excel");
                // Attempt to open the CSV file using a FileStream to check for exclusive access
                stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException ex)
            {
                // The file is locked, prompt the user to close Excel before continuing
                M_MyTaskDialog("Warning:", $"The CSV file below is currently locked by Excel: " +
                                           $"\n{fileName}\n{ex} " +
                                           $"\n\nPlease close Excel, \nthen click CLOSE to continue!");
                //return;
            }
            finally
            {
                stream?.Dispose();
            }

            // Continue with processing the CSV file
            // ...
        }


        // Testing ======

        //public List<string> GetVisibleParametersInSchedule(ViewSchedule schedule)
        //{
        //    List<string> visibleParameters = new List<string>();

        //    // Get the schedule fields
        //    IList<SchedulableField> schedulableFields = schedule.Definition.GetSchedulableFields();

        //    ParameterSet paramSet = schedule.Parameters;

        //    // Iterate over each field in the schedule
        //    foreach (SchedulableField field in schedulableFields)
        //    {
        //        // Get the parameter associated with the field
        //        //Parameter parameter = field.GetFieldId().GetField(schedule.Document).GetDefinition();
        //        Parameter parameter = schedule.Parameters;


        //        if (parameter != null)
        //        {
        //            string parameterName = parameter.Name;

        //            // Add the parameter name to the list of visible parameters
        //            visibleParameters.Add(parameterName);
        //        }
        //    }

        //    return visibleParameters;
        //}


        // End Testing ======
        public void ExportViewScheduleBasic(ViewSchedule schedule, ExcelWorksheet worksheet) // bookmark CTRL + K K  . next K N
        {
            string scheduleUniqueId = schedule.UniqueId;
            var dt = CreateDataTable(schedule);

            if (dt.Rows.Count > 0)
            {
                var dtRearranged = RearrangeColumns(dt);
                dtRearranged = InsertUniqueId(dtRearranged, scheduleUniqueId);
                //var dtRearranged = InsertUniqueId(dt, scheduleUniqueId);

                worksheet.Cells.LoadFromDataTable(dtRearranged, true);
                WorksheetFormatting(worksheet);
            }
        }

        private void AutoFitColumns(ExcelWorksheet worksheet, int rowIndexToFitTo)
        {
            for (int columnIndex = 1; columnIndex <= worksheet.Dimension.Columns; columnIndex++)
            {
                var cellValue = worksheet.Cells[rowIndexToFitTo, columnIndex].Value;
                if (cellValue != null)
                {
                    var cellTextLength = cellValue.ToString().Length;
                    var column = worksheet.Column(columnIndex);
                    column.Width = cellTextLength + 2; // Adjust the value as needed
                }
            }
        }
        public DataTable CreateDataTable(ViewSchedule schedule)
        {
            var dt = new DataTable();

            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var columnName = field.GetName(); // Parameter names
                var fieldType = typeof(string);

                // Ensure column names are unique by appending a number if necessary
                var i = 1;
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;

            if (nRows > 1)
            {
                // Set the values of the first row to the column headings
                var columnNameRow = dt.NewRow();

                int actualIndex = 0;
                for (var j = 0; j < nColumns; j++)
                {
                    //var field = viewSchedule.Definition.GetField(j);
                    var field = viewSchedule.Definition.GetField(actualIndex);
                    actualIndex++;
                    if (field.IsHidden)
                    {
                        j--;
                        continue;
                    }
                    columnNameRow[j] = field.ColumnHeading;
                }
                dt.Rows.Add(columnNameRow);

                // Populate data rows
                for (var i = 2; i < nRows; i++) // start at row index 2. to skip schedule title and header rows
                {
                    var dataRow = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j); // Gets the displayed schedule data
                        if (val.ToString() != "")
                            dataRow[j] = val;
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            return dt;
        }
        public DataTable RearrangeColumns(DataTable dt)
        {
            var lastColumnIndex = dt.Columns.Count - 1;
            var lastColumn = dt.Columns[lastColumnIndex];

            // Create a new DataTable with the columns rearranged
            var dtRearranged = new DataTable();
            dtRearranged.Columns.Add(lastColumn.ColumnName, lastColumn.DataType);
            foreach (DataColumn column in dt.Columns)
            {
                if (column != lastColumn)
                    dtRearranged.Columns.Add(column.ColumnName, column.DataType);
            }

            // Copy the data rows with the columns rearranged
            foreach (DataRow row in dt.Rows)
            {
                var newRow = dtRearranged.NewRow();
                newRow[0] = row[lastColumnIndex];
                for (var i = 0; i < lastColumnIndex; i++)
                    newRow[i + 1] = row[i];
                dtRearranged.Rows.Add(newRow);
            }

            return dtRearranged;
        }
        public DataTable InsertUniqueId(DataTable dt, string scheduleUniqueId)
        {
            var newRowAtIndex0 = dt.NewRow();
            newRowAtIndex0[0] = scheduleUniqueId;
            dt.Rows.InsertAt(newRowAtIndex0, 0);

            return dt;
        }
        #region Excel Formatting methods
        public void WorksheetFormatting(ExcelWorksheet worksheet)
        {
            FormatRow3Style(worksheet);
            AutoFitColumns(worksheet, 3);
            HideFirstTwoRows(worksheet);
            HideFirstColumn(worksheet);
            FreezeFirstThreeRows(worksheet);
        }
        public void HideFirstTwoRows(ExcelWorksheet worksheet)
        {
            // Hide the first two rows
            worksheet.Row(1).Hidden = true;
            worksheet.Row(2).Hidden = true;
        }
        public void HideFirstColumn(ExcelWorksheet worksheet)
        {
            // Hide the first column
            worksheet.Column(1).Hidden = true;
        }
        public void FreezeFirstThreeRows(ExcelWorksheet worksheet)
        {
            // Freeze the first three rows
            worksheet.View.FreezePanes(4, 1);
        }
        public void FormatRow3Style(ExcelWorksheet worksheet)
        {
            var lastColumnIndex = worksheet.Dimension.End.Column;

            // Apply bold font to the row
            worksheet.Cells[3, 1, 3, lastColumnIndex].Style.Font.Bold = true;

            // Set the background color to light gray for cells with content
            for (int columnIndex = 1; columnIndex <= lastColumnIndex; columnIndex++)
            {
                var cell = worksheet.Cells[3, columnIndex];
                if (cell.Value != null)
                {
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }
            }
        }
        #endregion








        public void ExportViewScheduleBasic9(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            string scheduleUniqueId = schedule.UniqueId;
            var dt = new DataTable();

            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var columnName = field.GetName();
                var fieldType = typeof(string);

                // Ensure column names are unique by appending a number if necessary
                var i = 1;
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;

            if (nRows > 1)
            {
                // Set the values of the first row to the column headings
                var columnNameRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var field = viewSchedule.Definition.GetField(j);
                    columnNameRow[j] = field.ColumnHeading;
                }
                dt.Rows.Add(columnNameRow);

                // Populate data rows
                for (var i = 2; i < nRows; i++)
                {
                    var dataRow = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "")
                            dataRow[j] = val;
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Rearrange columns: Move last column to the first column
                var lastColumnIndex = dt.Columns.Count - 1;
                var lastColumn = dt.Columns[lastColumnIndex];

                // Create a new DataTable with the columns rearranged
                var dtRearranged = new DataTable();
                dtRearranged.Columns.Add(lastColumn.ColumnName, lastColumn.DataType);
                foreach (DataColumn column in dt.Columns)
                {
                    if (column != lastColumn)
                        dtRearranged.Columns.Add(column.ColumnName, column.DataType);
                }

                // Copy the data rows with the columns rearranged
                foreach (DataRow row in dt.Rows)
                {
                    var newRow = dtRearranged.NewRow();
                    newRow[0] = row[lastColumnIndex];
                    for (var i = 0; i < lastColumnIndex; i++)
                        newRow[i + 1] = row[i];
                    dtRearranged.Rows.Add(newRow);
                }

                // Insert a new row at index 0 and add scheduleUniqueId to column index 0
                var newRowAtIndex0 = dtRearranged.NewRow();
                newRowAtIndex0[0] = scheduleUniqueId;
                dtRearranged.Rows.InsertAt(newRowAtIndex0, 0);

                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dtRearranged, true);

                //worksheet.Cells.AutoFitColumns();
                // Auto-fit columns to the third row text
                int rowIndexToFitTo = 3;
                M_AutoFitColumns(worksheet, rowIndexToFitTo);
            }
        }
        public void M_AutoFitColumns(ExcelWorksheet worksheet, int rowIndex)
        {
            for (int columnIndex = 1; columnIndex <= worksheet.Dimension.Columns; columnIndex++)
            {
                var cellValue = worksheet.Cells[rowIndex, columnIndex].Value;

                if (cellValue != null)
                {
                    var cellTextLength = cellValue.ToString().Length;
                    var column = worksheet.Column(columnIndex);
                    column.Width = cellTextLength + 2; // Adjust the value as needed
                }
            }
        }

        public void ExportViewScheduleBasic8(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            string scheduleUniqueId = schedule.UniqueId;
            var dt = new DataTable();

            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var columnName = field.GetName();
                var fieldType = typeof(string);

                // Ensure column names are unique by appending a number if necessary
                var i = 1;
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;

            if (nRows > 1)
            {
                // Set the values of the first row to the column headings
                var columnNameRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var field = viewSchedule.Definition.GetField(j);
                    columnNameRow[j] = field.ColumnHeading;
                }
                dt.Rows.Add(columnNameRow);

                // Populate data rows
                for (var i = 2; i < nRows; i++)
                {
                    var dataRow = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "")
                            dataRow[j] = val;
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Rearrange columns: Move last column to the first column
                var lastColumnIndex = dt.Columns.Count - 1;
                var lastColumn = dt.Columns[lastColumnIndex];

                // Create a new DataTable with the columns rearranged
                var dtRearranged = new DataTable();
                dtRearranged.Columns.Add(lastColumn.ColumnName, lastColumn.DataType);
                foreach (DataColumn column in dt.Columns)
                {
                    if (column != lastColumn)
                        dtRearranged.Columns.Add(column.ColumnName, column.DataType);
                }

                // Copy the data rows with the columns rearranged
                foreach (DataRow row in dt.Rows)
                {
                    var newRow = dtRearranged.NewRow();
                    newRow[0] = row[lastColumnIndex];
                    for (var i = 0; i < lastColumnIndex; i++)
                        newRow[i + 1] = row[i];
                    dtRearranged.Rows.Add(newRow);
                }

                // Insert a new row at index 0 and add scheduleUniqueId to column index 0
                var newRowAtIndex0 = dtRearranged.NewRow();
                newRowAtIndex0[0] = scheduleUniqueId;
                dtRearranged.Rows.InsertAt(newRowAtIndex0, 0);

                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dtRearranged, true);
                // RevitUtilities.AutoFitAllCol(worksheet);
            }
        }

        public void ExportViewScheduleBasic7(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            string scheduleUniqueId = schedule.UniqueId;
            var dt = new DataTable();

            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var columnName = field.GetName();
                var fieldType = typeof(string);

                // Ensure column names are unique by appending a number if necessary
                var i = 1;
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;

            if (nRows > 1)
            {
                // Set the values of the first row to the column headings
                var columnNameRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var field = viewSchedule.Definition.GetField(j);
                    columnNameRow[j] = field.ColumnHeading;
                }
                dt.Rows.Add(columnNameRow);

                // Populate data rows
                for (var i = 2; i < nRows; i++)
                {
                    var dataRow = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "")
                            dataRow[j] = val;
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Rearrange columns: Move last column to the first column
                var lastColumnIndex = dt.Columns.Count - 1;
                var lastColumn = dt.Columns[lastColumnIndex];

                // Create a new DataTable with the columns rearranged
                var dtRearranged = new DataTable();
                dtRearranged.Columns.Add(lastColumn.ColumnName, lastColumn.DataType);
                foreach (DataColumn column in dt.Columns)
                {
                    if (column != lastColumn)
                        dtRearranged.Columns.Add(column.ColumnName, column.DataType);
                }

                // Copy the data rows with the columns rearranged
                foreach (DataRow row in dt.Rows)
                {
                    var newRow = dtRearranged.NewRow();
                    newRow[0] = row[lastColumnIndex];
                    for (var i = 0; i < lastColumnIndex; i++)
                        newRow[i + 1] = row[i];
                    dtRearranged.Rows.Add(newRow);
                }

                // ChatGPT: insert a new row at index 0 and add the scheduleUniqueId to column index 0

                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dtRearranged, true);
                // RevitUtilities.AutoFitAllCol(worksheet);
            }
        }




        public void ExportViewScheduleBasic6(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            string scheduleUniqueId = schedule.UniqueId;

            var dt = new DataTable();

            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var columnName = field.GetName(); // Use the field name as the columnName
                var fieldType = typeof(string);
                var i = 1;
                // Ensure column names are unique by appending a number if necessary
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;

            if (nRows > 1)
            {
                // Set the values of the first row to the column headings
                var columnNameRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var field = viewSchedule.Definition.GetField(j);
                    columnNameRow[j] = field.ColumnHeading;
                }
                dt.Rows.Add(columnNameRow);

                // Populate data rows
                for (var i = 2; i < nRows; i++)
                {
                    var dataRow = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "") dataRow[j] = val;
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dt, true);
                // RevitUtilities.AutoFitAllCol(worksheet); - Uncomment or implement the AutoFitAllCol method if needed
            }
        }

        public void ExportViewScheduleBasic5(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            var dt = new DataTable();
            var emptyRow = dt.NewRow();
            dt.Rows.InsertAt(emptyRow, 0); // Insert an empty row at index 0

            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var fieldType = typeof(string);
                var columnName = field.GetName(); // Use the field name as the columnName
                var i = 1;
                // Ensure column names are unique by appending a number if necessary
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var viewScheduleDefinition = viewSchedule.Definition;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;
            if (nRows > 1)
            {
                // Set the values of the first row to the field names from the definition
                //var fieldRow = dt.NewRow();
                //for (var j = 0; j < nColumns; j++)
                //{
                //    var field = viewScheduleDefinition.GetField(j);
                //    fieldRow[j] = field.GetName();
                //}
                //dt.Rows.Add(fieldRow);

                // Set the values of the second row to the column names

                var columnNameRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var field = viewScheduleDefinition.GetField(j);
                    columnNameRow[j] = field.ColumnHeading;
                }
                dt.Rows.Add(columnNameRow);

                // Starts at 2 to skip the header rows
                for (var i = 2; i < nRows; i++)
                {
                    var data = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "") data[j] = val;
                    }
                    dt.Rows.Add(data);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dt, true);
                // RevitUtilities.AutoFitAllCol(worksheet); - Uncomment or implement the AutoFitAllCol method if needed
            }
        }




        public void ExportViewScheduleBasic4(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            var dt = new DataTable();
            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var fieldType = typeof(string);
                var columnName = field.ColumnHeading; // Use the field name as the columnName
                var i = 1;
                // Ensure column names are unique by appending a number if necessary
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.ColumnHeading}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var viewScheduleDefinition = viewSchedule.Definition;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;
            if (nRows > 1)
            {
                // Set the values of the first row to the field names from the definition
                var fieldRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var field = viewScheduleDefinition.GetField(j);
                    fieldRow[j] = field.GetName();
                }
                dt.Rows.Add(fieldRow);

                // Set the values of the second row to the values from the first row of the schedule
                var firstDataRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var value = viewSchedule.GetCellText(SectionType.Body, 1, j);
                    firstDataRow[j] = value.ToString();
                }
                dt.Rows.Add(firstDataRow);

                // Starts at 2 to skip the first two rows
                for (var i = 2; i < nRows; i++)
                {
                    var data = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "") data[j] = val;
                    }
                    dt.Rows.Add(data);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dt, true);
                // RevitUtilities.AutoFitAllCol(worksheet); - Uncomment or implement the AutoFitAllCol method if needed
            }
        }

        public void ExportViewScheduleBasic3(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            var dt = new DataTable();
            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var fieldType = typeof(string);
                var columnName = field.GetName(); // Use the parameter name as the columnName
                var i = 1;
                // Ensure column names are unique by appending a number if necessary
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var viewScheduleDefinition = viewSchedule.Definition;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;
            if (nRows > 1)
            {
                // Set the values of the first row to the current parameter values
                var parameterRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    var parameter = viewSchedule.GetCellText(SectionType.Body, 1, j);
                    parameterRow[j] = parameter.ToString();
                }
                dt.Rows.Add(parameterRow);

                // Set the values of the second row to the current column names
                var headerRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    headerRow[j] = dt.Columns[j].ColumnName;
                }
                dt.Rows.Add(headerRow);

                // Starts at 2 to skip the header rows
                for (var i = 2; i < nRows; i++)
                {
                    var data = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "") data[j] = val;
                    }
                    dt.Rows.Add(data);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dt, true);
                // RevitUtilities.AutoFitAllCol(worksheet); - Uncomment or implement the AutoFitAllCol method if needed
            }
        }

        public void ExportViewScheduleBasic2(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            var dt = new DataTable();
            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var fieldType = typeof(string);
                var columnName = field.GetName(); // Use the parameter name as the columnName
                var i = 1;
                // Ensure column names are unique by appending a number if necessary
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var viewScheduleDefinition = viewSchedule.Definition;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;
            if (nRows > 1)
            {
                // Insert a new row before the current row 1
                var newRow = dt.NewRow();
                dt.Rows.InsertAt(newRow, 0);

                // Set the values of the second row to the current column names
                var headerRow = dt.NewRow();
                for (var j = 0; j < nColumns; j++)
                {
                    headerRow[j] = dt.Columns[j].ColumnName;
                }
                dt.Rows.InsertAt(headerRow, 1);

                // Starts at 2 to skip the header rows
                for (var i = 2; i < nRows; i++)
                {
                    var data = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        // Retrieve the cell value for each column
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "") data[j] = val;
                    }
                    dt.Rows.Add(data);
                }
            }

            if (dt.Rows.Count > 0)
            {
                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dt, true);
                // RevitUtilities.AutoFitAllCol(worksheet); - Uncomment or implement the AutoFitAllCol method if needed
            }
        }


        public void ExportViewScheduleBasic_Old(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            var dt = new DataTable();
            // Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var fieldType = typeof(string);
                var columnName = field.ColumnHeading;
                var i = 1;
                // Ensure column names are unique by appending a number if necessary
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            // Content display
            var viewSchedule = schedule;
            var viewScheduleDefinition = viewSchedule.Definition; //<-- Added
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;
            if (nRows > 1)
            {
                // Starts at 1 to skip the header row
                for (var i = 1; i < nRows; i++)
                {
                    var data = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        //if (i == 1)
                        //{
                        //    try
                        //    {
                        //        // Retrieve the parameter name for the current column index
                        //        var field = viewScheduleDefinition.GetField(j);
                        //        if (field.IsHidden)
                        //            continue;
                        //        object val = field.GetName();

                        //        //if (field.ColumnHeading == "Dev_Text_1") //<--- Left off here!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        //        //{
                        //        //    val = "Dev_Text_1";
                        //        //    if (val.ToString() != "") data[j] = val;
                        //        //    continue;
                        //        //}

                        //        if (val.ToString() != "") data[j] = val;
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Debug.Print($"Error retrieving parameter name at column index {j}: {ex.Message}");
                        //        // Handle the exception or log the error as needed
                        //    }
                        //}
                        //else
                        //{
                        //    // Retrieve the cell value for non-header rows
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "") data[j] = val;
                        //}
                    }
                    dt.Rows.Add(data);
                }
            }

            //<--
            if (dt.Rows.Count > 0)
            {
                // Load the data into the worksheet
                worksheet.Cells.LoadFromDataTable(dt, true);
                // RevitUtilities.AutoFitAllCol(worksheet); - Uncomment or implement the AutoFitAllCol method if needed
            }
        }

        //public void ExportViewScheduleBasic(ViewSchedule schedule, ExcelWorksheet worksheet)
        //{
        //    var dt = new DataTable();
        //    // Definition of columns
        //    var fieldsCount = schedule.Definition.GetFieldCount();
        //    for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
        //    {
        //        var field = schedule.Definition.GetField(fieldIndex);
        //        if (field.IsHidden) continue;
        //        var fieldType = typeof(string);
        //        var columnName = field.ColumnHeading;
        //        var i = 1;
        //        // Ensure column names are unique by appending a number if necessary
        //        while (dt.Columns.Contains(columnName))
        //        {
        //            columnName = $"{field.GetName()}({i})";
        //            i++;
        //        }

        //        dt.Columns.Add(columnName, fieldType);
        //    }

        //    // Content display
        //    var viewSchedule = schedule;
        //    var viewScheduleDefinition = viewSchedule.Definition; //<-- Added
        //    var table = viewSchedule.GetTableData();
        //    var section = table.GetSectionData(SectionType.Body);
        //    var nRows = section.NumberOfRows;
        //    var nColumns = section.NumberOfColumns;
        //    if (nRows > 1)
        //    {
        //        // Starts at 1 to skip the header row
        //        for (var i = 1; i < nRows; i++)
        //        {
        //            var data = dt.NewRow();
        //            for (var j = 0; j < nColumns; j++)
        //            {
        //                //if (i == 1)
        //                //{
        //                //    try
        //                //    {
        //                //        // Retrieve the parameter name for the current column index
        //                //        var field = viewScheduleDefinition.GetField(j);
        //                //        object val = field.GetName();

        //                //        if (field.ColumnHeading == "Dev_Text_1") //<--- Left off here!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        //                //        {
        //                //            val = "Dev_Text_1";
        //                //            if (val.ToString() != "") data[j] = val;
        //                //            continue;
        //                //        }

        //                //        if (val.ToString() != "") data[j] = val;
        //                //    }
        //                //    catch (Exception ex)
        //                //    {
        //                //        Debug.Print($"Error retrieving parameter name at column index {j}: {ex.Message}");
        //                //        // Handle the exception or log the error as needed
        //                //    }
        //                //}
        //                //else
        //                //{
        //                    //// Retrieve the cell value for non-header rows
        //                    object val = viewSchedule.GetCellText(SectionType.Body, i, j);
        //                if (val.ToString() != "") data[j] = val;
        //            //}
        //        }

        //            dt.Rows.Add(data);
        //        }
        //    }


        //    ////Content display
        //    //var viewSchedule = schedule;
        //    //var table = viewSchedule.GetTableData();
        //    //var section = table.GetSectionData(SectionType.Body);
        //    //var nRows = section.NumberOfRows;
        //    //var nColumns = section.NumberOfColumns;
        //    //if (nRows > 1)
        //    //    //Starts at 1 so as not to display the header
        //    //    for (var i = 1; i < nRows; i++)
        //    //    {
        //    //        var data = dt.NewRow();
        //    //        for (var j = 0; j < nColumns; j++)
        //    //        {
        //    //            object val = viewSchedule.GetCellText(SectionType.Body, i, j);
        //    //            if (val.ToString() != "") data[j] = val;
        //    //        }

        //    //        dt.Rows.Add(data);
        //    //    }
        //    if (dt.Rows.Count > 0)
        //    {
        //        // Load the data into the worksheet
        //        worksheet.Cells.LoadFromDataTable(dt, true);
        //        // RevitUtilities.AutoFitAllCol(worksheet); - Uncomment or implement the AutoFitAllCol method if needed
        //    }
        //}



        public static ExcelPackage Create_ExcelFile(string filePath)
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;  // Set the license context for EPPlus to NonCommercial

            // Open Excel file using EPPlus library
            ExcelPackage excelFile = new ExcelPackage(filePath);  // Create an instance of ExcelPackage by providing the file path
                                                                  //ExcelWorkbook workbook = excelFile.Workbook;  // Get the workbook from the Excel package
                                                                  // ExcelWorksheet worksheet = workbook.Worksheets[1];  // Get the first worksheet (index 0) from the workbook
                                                                  //ExcelWorksheet worksheet = workbook.Worksheets.Add("Sheet1");

            return excelFile;
        }

    }
}