#region Header
// The MIT License (MIT)
//
// Copyright (c) 2013 Kyle T. Morin
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
#endregion

#region Namespaces
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Lighting;
using System.Windows.Forms;
#endregion

namespace STFExporter
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Application _app;
        public Document _doc;
        public string writer = "";
        public double meterMultiplier = 0.3048;
        public List<ElementId> distinctLuminaires = new List<ElementId>();
        public string stfVersionNum = "1.0.5";

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            _app = app;
            _doc = doc;

            //Set project units to Meters then back after

            Units pUnit = doc.GetUnits();
            FormatOptions formatOptions = pUnit.GetFormatOptions(UnitType.UT_Length);

            DisplayUnitType curUnitType = pUnit.GetDisplayUnitType();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("STF EXPORT");
                DisplayUnitType meters = DisplayUnitType.DUT_METERS;
                formatOptions.DisplayUnits = meters;                
                //formatOptions.Units = meters;
                //formatOptions.Rounding = 0.0000000001;
                formatOptions.Accuracy = 0.0000000001;

                ElementLevelFilter filter = new ElementLevelFilter(doc.ActiveView.GenLevel.Id);

                FilteredElementCollector fec = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WherePasses(filter);

                int numOfRooms = fec.Count();
                writer += "[VERSION]\n"
                    + "STFF=" + stfVersionNum + "\n"
                    + "Progname=Revit\n"
                    + "Progvers=" + app.VersionNumber + "\n"
                    + "[Project]\n"
                    + "Name=" + _doc.ProjectInformation.Name + "\n"
                    + "Date=" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "\n"
                    + "Operator=" + app.Username + "\n"
                    + "NrRooms=" + numOfRooms + "\n";

                for (int i = 1; i < numOfRooms + 1; i++)
                {
                    writer += "Room" + i.ToString() + "=ROOM.R" + i.ToString() + "\n";
                }

                int increment = 1;

                //Space writer                
                foreach (Space s in fec)
                {
                    writer += "[ROOM.R" + increment.ToString() + "]\n";
                    SpaceInfoWriter(s.Id);
                    increment++;                
                }

                //Writout Luminaires to bottom
                writeLumenairs();

                //reset back to original
                formatOptions.DisplayUnits = curUnitType;

                tx.Commit();

                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "STF File | *.stf",
                    FilterIndex = 2,
                    RestoreDirectory = true
                };

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter sw = new StreamWriter(dialog.FileName);
                    string[] ar = writer.Split('\n');
                    for (int i = 0; i < ar.Length; i++)
                    {
                        sw.WriteLine(ar[i]);
                    }
                    sw.Close();
                }

                return Result.Succeeded;
            }
        }

        private void writeLumenairs()
        {
            FilteredElementCollector fecFixtures = new FilteredElementCollector(_doc)
            .OfCategory(BuiltInCategory.OST_LightingFixtures)
            .OfClass(typeof(FamilySymbol));

            foreach (FamilySymbol fs in fecFixtures)
            {
                string load = "";
                string flux = "";
                ParameterSet pset = fs.Parameters;

                Parameter pload = fs.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD);
                Parameter pflux = fs.get_Parameter(BuiltInParameter.FBX_LIGHT_LIMUNOUS_FLUX);
                if (pflux != null)
                {
                    load = pload.AsValueString();
                    flux = pflux.AsValueString();

                    writer += "[" + fs.Name.Replace(" ","") + "]\n";
                    writer += "Manufacturer=" + "\n"
                        + "Name=" + "\n"
                        + "OrderNr=" + "\n"
                        + "Box=1 1 0" + "\n" //need to fix per bounding box size (i guess);
                        + "Shape=0" + "\n"
                        + "Load=" + load.Remove(load.Length - 3) + "\n"
                        + "Flux=" + flux.Remove(flux.Length - 3) + "\n"
                        + "NrLamps=" + getNumLamps(fs) + "\n"
                        + "MountingType=1\n";
                }
            }
            

            
        }

        private string getNumLamps(FamilySymbol fs)
        {
            string results;
            if (fs.get_Parameter("Number of Lamps") != null)
            {
                results = fs.get_Parameter("Number of Lamps").ToString();
                return results;
            }
            else
            {
                //Parse from IES file
                var file = fs.get_Parameter(BuiltInParameter.FBX_LIGHT_PHOTOMETRIC_FILE);
                return "1"; //for now
            }
        }    

        private void SpaceInfoWriter(ElementId spaceID)
        {
            const double MAX_ROUNDING_PRECISION = 0.000000000001;
            
            //Get info from Space
            Space roomSpace = _doc.GetElement(spaceID) as Space;
            //Space roomSpace = _doc.get_Element(spaceID) as Space;

            //VARS
            string name = roomSpace.Name;
            double height = roomSpace.UnboundedHeight * meterMultiplier;
            double workPlane = roomSpace.LightingCalculationWorkplane * meterMultiplier;
            
            //Get room vertices
            List<String> verticies = new List<string>();
            verticies = getVertexPoints(roomSpace);

            int numPoints = getVertexPointNums(roomSpace);

            //Writeout Top part of room entry
            writer += "Name=" + name + "\n"
                + "Height=" + height.ToString() + "\n"
                + "WorkingPlane=" + workPlane.ToString() + "\n"
                + "NrPoints=" + numPoints.ToString() + "\n";

            //Write vertices for each point in vertex numbers
            for (int i = 0; i < numPoints; i++)
            {
                int i2 = i + 1;
                writer += "Point" + i2 + "=" + verticies.ElementAt(i) + "\n";
            }

            double cReflect = roomSpace.CeilingReflectance;
            double fReflect = roomSpace.FloorReflectance;
            double wReflect = roomSpace.WallReflectance;
            
            //Write out ceiling reflectance
            writer += "R_Ceiling=" + cReflect.ToString() + "\n";

            IList<ElementId> elemIds = roomSpace.GetMonitoredLocalElementIds();
            foreach (ElementId e in elemIds)
            {
                TaskDialog.Show("s", _doc.GetElement(e).Name);
            }

            //Get fixtures within space
            FilteredElementCollector fec = new FilteredElementCollector(_doc)
            .OfCategory(BuiltInCategory.OST_LightingFixtures)
            .OfClass(typeof(FamilyInstance));

            int count = 0;
            foreach (FamilyInstance fi in fec)
            {
                if (fi.Space.Id == spaceID)
                {
                    FamilySymbol fs = _doc.GetElement(fi.GetTypeId()) as FamilySymbol;
                    //FamilySymbol fs = _doc.get_Element(fi.GetTypeId()) as FamilySymbol;

                    int lumNum = count + 1;
                    string lumName = "Lum" + lumNum.ToString();
                    LocationPoint locpt = fi.Location as LocationPoint;
                    XYZ fixtureloc = locpt.Point;
                    double X = fixtureloc.X * meterMultiplier;
                    double Y = fixtureloc.Y * meterMultiplier;
                    double Z = fixtureloc.Z * meterMultiplier;

                    double rotation = locpt.Rotation;
                    writer += lumName + "=" + fs.Name.Replace(" ","") + "\n";
                    writer += lumName + ".Pos=" + X.ToString() + " " + Y.ToString() + " " + Z.ToString() + "\n";
                    writer += lumName + ".Rot=0 0 0" + "\n"; //need to figure out this rotation; Update: cannot determine. Almost impossible for Dialux

                    count++;
                }
            }

            //Writeout Lums part
            writer += "NrLums=" + count.ToString() + "\n"
                + "NrStruct=0\n" + "NrFurns=0\n";
        }

        private List<string> getVertexPoints(Space roomSpace)
        {
            List<string> verticies = new List<string>();
            SpatialElementBoundaryOptions opts = new SpatialElementBoundaryOptions();
            IList<IList<Autodesk.Revit.DB.BoundarySegment>> bsa = roomSpace.GetBoundarySegments(opts);
            if (bsa.Count > 0)
            {
                foreach (Autodesk.Revit.DB.BoundarySegment bs in bsa[0])
                {
                    var X = bs.Curve.get_EndPoint(0).X * meterMultiplier;
                    var Y = bs.Curve.get_EndPoint(0).Y * meterMultiplier;
                    verticies.Add(X.ToString() + " " + Y.ToString());
                }

            }
            else
            {
                verticies.Add("0 0");
            }
            return verticies;
        }

        private int getVertexPointNums(Space roomSpace)
        {
            SpatialElementBoundaryOptions opts = new SpatialElementBoundaryOptions();
            IList<IList<Autodesk.Revit.DB.BoundarySegment>> bsa = roomSpace.GetBoundarySegments(opts);

            return bsa[0].Count;

        }
    }
}