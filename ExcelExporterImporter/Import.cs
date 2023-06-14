#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Windows.Input;
using System.Windows.Shapes;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

using Microsoft.VisualBasic.FileIO;

using OfficeOpenXml;
#endregion

#region Begining of doc
namespace ORH_ExcelExporterImporter
{
    [Transaction(TransactionMode.Manual)]
    public class Import : MyUtils, IExternalCommand
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


            string excelFilePath = @"C:\Users\ohernandez\Desktop\Revit_Exports\rme_advanced_sample_project_2020.xlsx";
            //var excelFile = M_ReadExcelFile(excelFilePath);
            // Set EPPlus license context
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;  // Set the license context for EPPlus to NonCommercial

            using (var excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var excelFile = M_ReadExcelFile(excelPackage);
                if (excelFile == null) { return Result.Cancelled; }

                var excelSheetData = GetScheduleDataFromSheet(excelFile[6]);


                using (Transaction trans = new Transaction(doc, "Import Schedules"))
                {
                    trans.Start();

                    ImportSchedules(doc, excelSheetData);

                    trans.Commit();
                }
            }

            return Result.Succeeded;
        }



        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}