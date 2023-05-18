using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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

using Microsoft.Win32;

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
            var viewSchedule = _GetViewScheduleByName(doc, _viewScheduleName);
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
            ViewSchedule schedule = _GetViewScheduleByName(doc, scheduleName);

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
            //using (Transaction tx = new Transaction(doc, $"Update Parameters")) // Start a new transaction to make changes to the elements in Revit
            //{
            //    tx.Start(); // Lock the doc while changes are made in the transaction


            ViewSchedule _viewScheduleToUpdate = _GetViewScheduleByName(doc, _viewScheduleName); // Get the schedule by name
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
                    }
                }

            }

            //    tx.Commit();
            //}


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
            ViewSchedule schedule = _GetViewScheduleByName(doc, _viewScheduleName); // Get the schedule by name

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

        public static void addSharedParamToSchedule(Document doc, Autodesk.Revit.ApplicationServices.Application app, string scheduleName)
        {
            var originalSharedParametersFilename = app.SharedParametersFilename;
            // Create a new Revit application object.
            //Application app = new Application();

            // Set the SharedParametersFilename property to the path of the shared parameters file.
            //app.SharedParametersFilename = @"C:\Users\ohernandez\Desktop\Revit_Exports\SharedParams\ACCO -- Dev_Revit Shared Parameters.txt";
            app.SharedParametersFilename = M_CreateSharedParametersFile();

            // Open the shared parameters file.
            DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();

            DefinitionGroup definitionGroup = sharedParameterFile.Groups.get_Item("Dev_Group_Common");
            Definition definition = definitionGroup.Definitions.get_Item("Dev_Text_1");
            GetParameterDefinitionFromFile(sharedParameterFile, "Dev_Group_Common", "Dev_Text_1");

            ViewSchedule scheduleView = _GetViewScheduleByName(doc, scheduleName);

            // Get the schedule definition from the view
            ScheduleDefinition scheduleDef = scheduleView.Definition;

            // Create the schedule field using the shared parameter
            //ScheduleField scheduleField = scheduleDef.AddField(definition);

            // Modify any additional properties of the schedule field if needed
            // scheduleField.ColumnHeading = "Dev Text 1";

            // Refresh the schedule to reflect the changes
            scheduleView.Document.Regenerate();

            //// Save the document if necessary
            //scheduleView.Document.Save();



            /////////////////////////
            ////(Document doc, string scheduleName, string fieldName)

            //ViewSchedule schedule = _GetViewScheduleByName(doc, scheduleName);

            //// Get the definition of the schedule.
            //ScheduleDefinition definition = schedule.Definition;

            //var t = definition.par

            //// Get the list of available fields.
            //IList<SchedulableField> availableFields = definition.GetSchedulableFields();

            //// Find the field you want to add.
            //SchedulableField field = availableFields.FirstOrDefault(f => f.GetName(doc) == "Dev_Text_1");//fieldName);

            //// Add the field to the schedule.
            //var fieldAdded = definition.AddField(field);

            //app.SharedParametersFilename = originalSharedParametersFilename;
            ////return fieldAdded;


        }


        public static void M_Add_Dev_Text_1(Autodesk.Revit.ApplicationServices.Application app, Document doc, string _curScheduleName, BuiltInCategory _builtInCat)
        {
            var sv = _GetViewScheduleByName(doc, _curScheduleName);
            var i = sv.Category;

            //define category for shared param
            //Category myCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment);
            Category myCat = doc.Settings.Categories.get_Item(_builtInCat);
            CategorySet myCatSet = doc.Application.Create.NewCategorySet();
            myCatSet.Insert(myCat);

            app.SharedParametersFilename = @"Y:\\DATABASES\\ACCORevit\\02-SHARED PARAMETERS\\ACCO -- Revit Shared Parameters.txt";

            var originalSharedParametersFilename = app.SharedParametersFilename;
            // Set the SharedParametersFilename property to the path of the shared parameters file.
            //app.SharedParametersFilename = @"C:\Users\ohernandez\Desktop\Revit_Exports\SharedParams\ACCO -- Dev_Revit Shared Parameters.txt";
            app.SharedParametersFilename = M_CreateSharedParametersFile();

            // Open the shared parameters file.
            DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();
            var curDef = MyUtils.GetParameterDefinitionFromFile(sharedParameterFile, "Dev_Group_Common", "Dev_Text_1");
            //create binding
            ElementBinding curBinding = doc.Application.Create.NewInstanceBinding(myCatSet);

            using (Transaction tx = new Transaction(doc, $"AddParam")) // Start a new transaction to make changes to the elements in Revit
            {
                tx.Start();
                //insert definition into binding
                //using (Transaction curTrans = new Transaction(doc, "Added Shared Parameter"))
                //{
                //    if (curTrans.Start() == TransactionStatus.Started)
                //    {
                //        //do something
                var paramAdded = doc.ParameterBindings.Insert(curDef, curBinding, BuiltInParameterGroup.PG_IDENTITY_DATA);
                //    }

                //    //commit changes
                //    //curTrans.RollBack();
                //    curTrans.Commit();
                //}

                //var af = 
                M_AddByNameAvailableFieldToSchedule(doc, _curScheduleName, "Dev_Text_1");
                _UpdateMyUniqueIDColumn(doc, _curScheduleName);
                tx.Commit();
                //tx.RollBack();
            }

            // Set the SharedParametersFilename property to the path of the shared parameters file.
            app.SharedParametersFilename = originalSharedParametersFilename;
            //app.SharedParametersFilename = @"Y:\\DATABASES\\ACCORevit\\02-SHARED PARAMETERS\\ACCO -- Revit Shared Parameters.txt";
            //app.SharedParametersFilename = @"C:\Users\ohernandez\Desktop\Revit_Exports\SharedParams\ACCO -- Dev_Revit Shared Parameters.txt";

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
            File.WriteAllLines(filePath, csvLines);
        }

        //###########################################################################################
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
        public static void MoveCsvLastColumnToFirst(string filePath)
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
            File.WriteAllLines(filePath, updatedLines);

        }

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


    }
}