STF Exporter plugin for Revit to DIALux
========================================

## Add-in for Revit that exports Spaces along with Light fixtures within those spaces out to the .STF file format used by DIALux software for photometry calculations.

### To use: 
	1. Download sourcecode and build solution.
	2. Copy both .dll and .addin file to your Revit Add-Ins folder (C:\Users\USERNAME\AppData\Roaming\Autodesk\Revit\Addins\VERSION)
	3. Start Revit instance, command located in Add-Ins Tab, External Tools

Example video of workflow: http://www.youtube.com/watch?v=XWg7VH4ChZM

#### Known Issues:

- Have to re-map light fixture types in DIALux after first import, but Locations and elevations are correct (this is DIALux problem, not Revit)
- If using Imperial measurements: Measurements/distances will be off slightly. (This is due to DIALux conversions of units from meters to Feet/Inches. Even when getting exact from Revit out to STF, DIALux still converts back to feet/inches not exact)