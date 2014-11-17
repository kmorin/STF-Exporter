#region Namespaces
using System;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO;
using System.Windows.Media.Imaging;
#endregion

namespace STFExporter
{
    class App : IExternalApplication
    {
        static readonly string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        static readonly string assyPath = Path.Combine(dir, "STFExporter.dll");
        static readonly string _imgFolder = Path.Combine(dir, "Images");

        public Result OnStartup(UIControlledApplication a)
        {
            try
            {
                AddRibbonPanel(a);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ribbon", ex.ToString());
            }

            return Result.Succeeded;
        }

        private void AddRibbonPanel(UIControlledApplication app)
        {
            RibbonPanel panel = app.CreateRibbonPanel("STF Exporter: v" + Assembly.GetExecutingAssembly().GetName().Version);

            PushButtonData pbd_STF = new PushButtonData("STFExport", "Export STF File", assyPath, "STFExporter.Command");
            PushButton pb_STFbutton = panel.AddItem(pbd_STF) as PushButton;
            pb_STFbutton.LargeImage = NewBitmapImage("stfexport.png");
            pb_STFbutton.ToolTip = "Export Revit Spaces to STF File";
            pb_STFbutton.LongDescription = "Exports Spaces in Revit model to STF file for use in application such as DIALux";

            //ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.ChmFile, dir + "/Resources/STFExporter Help.htm");
            ContextualHelp contextHelp = new ContextualHelp(ContextualHelpType.Url, "https://github.com/kmorin/STF-Exporter");

            pb_STFbutton.SetContextualHelp(contextHelp);
        }

        BitmapImage NewBitmapImage(string imgName)
        {
            return new BitmapImage(new Uri(Path.Combine(_imgFolder, imgName)));
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
