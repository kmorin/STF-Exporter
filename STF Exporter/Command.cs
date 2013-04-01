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
            ProjectUnit pUnit = doc.ProjectUnit;
            FormatOptions formatOptions = pUnit.get_FormatOptions(UnitType.UT_Length);

            DisplayUnitType curUnitType = formatOptions.Units;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("STF EXPORT");
                DisplayUnitType meters = DisplayUnitType.DUT_METERS;
                formatOptions.Units = meters;
                formatOptions.Rounding = 0.0000000001;
                
                FilteredElementCollector fec = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces);

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
                    writer += "[ROOM.R"+increment.ToString()+"]\n";
                    SpaceInfoWriter(s.Id);
                    increment++;
                }

                //Writout Luminaires to bottom
                writeLumenairs();

                //reset back to original
                formatOptions.Units = curUnitType;

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
                string load = fs.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD).AsValueString();
                string flux = fs.get_Parameter(BuiltInParameter.FBX_LIGHT_LIMUNOUS_FLUX).AsValueString();

                writer += "[" + fs.Name + "]\n";
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

        // TODO: Implement AnalysisVisualizationFramework
        /*
        public void RoomPointsVis(double spacing, Reference reference, Space roomSpace, List<FamilyInstance> lightfixtures, SpatialFieldManager sfm)
        {
            string s = "";

            Reference _ref = reference;
            IList<UV> uvPts = new List<UV>();
            List<double> doubleList = new List<double>();
            IList<ValueAtPoint> valList = new List<ValueAtPoint>();

            Options opt = new Options();
            opt.ComputeReferences = true;

            SpatialElement spaceElement = (SpatialElement)roomSpace;
            GeometryElement ge = spaceElement.get_Geometry(opt);
            GeometryObjectArray geoArr = ge.Objects;
            Solid sol = geoArr.get_Item(0) as Solid;
            Face _face = sol.Faces.get_Item(1) as Face;           
            
            BoundingBoxUV bb = _face.GetBoundingBox();
            UV min = bb.Min;
            UV max = bb.Max;

            //// _face transform
            UV faceCenter = new UV((max.U + min.U) / 2, (max.V + min.V) / 2);
            Transform computeDerivatives = _face.ComputeDerivatives(faceCenter);
            XYZ faceCenterNormal = computeDerivatives.BasisZ;

            // Normalize the normal vector and multiply by 2.5 
            XYZ faceCenterNormalMultiplied = faceCenterNormal.Normalize().Multiply(2.5);            
            Transform _transform = Transform.get_Translation(faceCenterNormalMultiplied);

            for (double u = min.U; u < max.U; u += (max.U - min.U) / 20)
            {
                for (double v = min.V; v < max.V; v += (max.V - min.V) / 20)
                {
                    UV uv = new UV(u, v);
                    if (_face.IsInside(uv))
                    {                        
                        XYZ pointToCalc = _face.Evaluate(uv);
                        double resultFC = CalcPointsFC(roomSpace, lightfixtures, pointToCalc);
                        uvPts.Add(uv);
                        doubleList.Add(resultFC);
                        s += "\n" + resultFC.ToString();
                        valList.Add(new ValueAtPoint(doubleList));
                        doubleList.Clear();
                    }
                }
            }

            //Visualization framwork
            FieldDomainPointsByUV pnts = new FieldDomainPointsByUV(uvPts);
            FieldValues vals = new FieldValues(valList);            
            AnalysisResultSchema resultSchema = new AnalysisResultSchema("Point by Point", "Illumination Point-By-Pont");

            //Units
            IList<string> names = new List<string> { "FC" };
            IList<double> multipliers = new List<double> { 1 };
            resultSchema.SetUnits(names, multipliers);

            //Add field primitive to view
            int idx = sfm.AddSpatialFieldPrimitive(_face,_transform);            
            int idx1 = sfm.RegisterResult(resultSchema);
            sfm.UpdateSpatialFieldPrimitive(idx, pnts, vals, idx1);

            
            //Test method to get values
            TaskDialog.Show("Res", s);

            using (Transaction tx = new Transaction(_doc))
            {
                tx.Start("DisplayAnalysisResults");
                bool avs = GetOrCreateAVS();
                tx.Commit();
            }

        }
        */

        private void SpaceInfoWriter(ElementId spaceID)
        {
            const double MAX_ROUNDING_PRECISION = 0.000000000001;
            
            // INFO NEEDED FOR ROOM/SPACE
            /*
             * [ROOM.R1]
             * Name=*SPACENAME*
             * Height=*SPACEHEIGHT* (in meters)
             * WorkingPlane=*WORKINGPLANE* (in meters)
             * NrPoints= *NUMBER OF VERTEX POITNS* (ex. 4 for square room)
             * Point1=X Y
             * Point2=X Y
             * Point3=X Y
             * Point4=X Y
             * R_Ceiling=*CEILINGREFLECTANCE* (double)
             * --Check if any fixtures in space, if none, skip. --
             * Lum1=LUMINAIRE.L1 (foreach luminaire in space new type)
             * Lum1.Pos=X Y Z (in meters)
             * Lum1.Rot=0 0 0 (need to figure out transform i guess)
             * Lum2=LUMINAIRE.L1
             * Lum2.Pos=
             * Lum2.Rot=
             * NrStruct=0
             * NrLums=*TOTALFIXTURESINSPACE*
             * NrFurns=0
             * [ROOM.R2]
             * 
             */
            
            //Get info from Space
            Space roomSpace = _doc.GetElement(spaceID) as Space;
            /* Not needed
            SpatialElementGeometryCalculator calc = new SpatialElementGeometryCalculator(_doc);
            SpatialElementGeometryResults results = calc.CalculateSpatialElementGeometry((SpatialElement)roomSpace);
            Solid geometry = results.GetGeometry();            
            */

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

                    int lumNum = count + 1;
                    string lumName = "Lum" + lumNum.ToString();
                    LocationPoint locpt = fi.Location as LocationPoint;
                    XYZ fixtureloc = locpt.Point;
                    double X = fixtureloc.X * meterMultiplier;
                    double Y = fixtureloc.Y * meterMultiplier;
                    double Z = fixtureloc.Z * meterMultiplier;

                    double rotation = locpt.Rotation;
                    writer += lumName + "=" + fs.Name +"\n";
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
            foreach (Autodesk.Revit.DB.BoundarySegment bs in bsa[0])
            {
                var X = bs.Curve.get_EndPoint(0).X * meterMultiplier;
                var Y = bs.Curve.get_EndPoint(0).Y * meterMultiplier;
                verticies.Add(X.ToString() + " " + Y.ToString());
            }

            return verticies;
        }

        private int getVertexPointNums(Space roomSpace)
        {
            SpatialElementBoundaryOptions opts = new SpatialElementBoundaryOptions();
            IList<IList<Autodesk.Revit.DB.BoundarySegment>> bsa = roomSpace.GetBoundarySegments(opts);

            return bsa[0].Count;

        }

        private double CalcPointsFC(Space roomSpace, List<FamilyInstance> lightfixtures, XYZ point)
        {
            XYZ pointToCalc = new XYZ(point.X, point.Y, 2.5);            
            double roomArea = roomSpace.Area;
            double roomPerim = roomSpace.Perimeter;
            double valueatPoint = 0;
            string valuestring = "";
            foreach (FamilyInstance fi in lightfixtures)
            {
                ElementId eid = fi.GetTypeId();
                Element e = _doc.GetElement(eid);
                PhotometricWebLightDistribution wl = new PhotometricWebLightDistribution(e.get_Parameter(BuiltInParameter.FBX_LIGHT_PHOTOMETRIC_FILE).AsString(), -90.0);
                double candelaPower = e.get_Parameter(BuiltInParameter.FBX_LIGHT_LIMUNOUS_INTENSITY).AsDouble();
                double lumens = e.get_Parameter(BuiltInParameter.FBX_LIGHT_LIMUNOUS_FLUX).AsDouble();
                double efficacy = e.get_Parameter(BuiltInParameter.FBX_LIGHT_EFFICACY).AsDouble();
                double llf = e.get_Parameter(BuiltInParameter.FBX_LIGHT_TOTAL_LIGHT_LOSS).AsDouble();
                double cu = fi.get_Parameter(BuiltInParameter.RBS_ROOM_COEFFICIENT_UTILIZATION).AsDouble();
                double workPlane = roomSpace.LightingCalculationWorkplane;
                double lightElev = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble();


                //Calc distance to point
                LocationPoint lightLocPt = fi.Location as LocationPoint;
                XYZ lightLoc = lightLocPt.Point;                
                double distToPoint = lightLoc.DistanceTo(pointToCalc);
                double angle = lightLoc.AngleTo(pointToCalc);
                double cosangle = Math.Cos(angle);

                // Calculation formula = 
                // Lumens = lamp lumens defined in type parameters
                // angle = angle to point (already expressed in cos needed) no need to adjust further I have found out
                // distToPoint = trig distance calced from light source to point
                // must square ^2 the result because resultant is expressed in lm/ft^2 so need to square in order to achieve
                //  true result of footcandles.
                double res = Math.Pow((lumens * angle) / Math.Pow(distToPoint, 2), 2);

                valueatPoint += res;
                valuestring += "\nLightLoc: " + lightLoc.ToString() + "\nDistToPt: " + distToPoint.ToString()
                    + "\nAngle: " + angle.ToString()
                    + "\nCosAngle: " + cosangle.ToString()
                    + "\nRes: " + res.ToString();
            }

            //TaskDialog.Show("result", "Result = " + valuestring);
            return valueatPoint;
        }
    }
}
