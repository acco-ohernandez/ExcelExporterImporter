#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.Versioning;
using System.Windows.Markup;

#endregion

namespace ExcelExporterImporter
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            // 1. Create ribbon tab
            string TabName = "ExporterImporter";
            try
            {
                app.CreateRibbonTab(TabName);
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists.");
            }

            // 2. Create ribbon panel 
            RibbonPanel panel = Utils.CreateRibbonPanel(app, TabName, "Revit Tools");

            // 3. Create button data instances
            ButtonDataClass myButtonData = new ButtonDataClass("btnExcelExporterExport", "Exporter", Export.GetMethod(), Properties.Resources.Blue_32, Properties.Resources.Blue_16, "This is a tooltip");

            ButtonDataClass myButtonData2 = new ButtonDataClass("btnExcelExporterImporter", "Importer", Import.GetMethod(), Properties.Resources.Green_32, Properties.Resources.Green_16, "This is a tooltip");

            // 4. Create buttons
            PushButton myButton = panel.AddItem(myButtonData.Data) as PushButton;
            PushButton myButton2 = panel.AddItem(myButtonData2.Data) as PushButton;
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }


    }
}
