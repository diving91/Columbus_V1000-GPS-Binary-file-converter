# Columbus_V1000-GPS-Binary-file-converter
 Convert Columbus V-1000 GPS Binary files (*.GPS) into Google Earth .kml file 

  Created by Diving91
  Version 0.1 - 2016-09-02

  - Parse the GPS original folder structure and copy Columbus Binary files in basedir with name format yyyy-mm-dd hh-mm-ss.GPS
  - Convert Binary files to csv files then into gpx format
  - Extract Waypoints (POIs) from csv files into a gpx file
  - Generate a kml file including all tracks and waypoints (POIs)

GPS is from Columbus: http://cbgps.com/v1000/index.html

This PowerShell Script uses other tools:
- http://www.gpsbabel.org/download.html
- http://static.routeconverter.com/download/RouteConverterCmdLine.jar
- Java JRE
- Google Earth

Photo of the GPS data logger: http://cbgps.com/v1000/front_home.jpg
