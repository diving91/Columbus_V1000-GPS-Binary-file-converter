<#
  Created by Diving91
  Version 0.1 - 2016-09-02

  - Parse the GPS original folder structure and copy Columbus Binary files in basedir with name format yyyy-mm-dd hh-mm-ss.GPS
  - Convert Binary files to csv files then into gpx format
  - Extract Waypoints from csv files into a gpx file
  - Generate a kml file including all tracks and waypoints

#>

# Current path
$CurrentDir = $(get-location).Path;
# Name for waypoint files that are extracted from Columbus file format
$wptFileName = 'myWaypoints.gpx'
# Name of the kml file to generate
$kmlFileName = 'All.kml'
# Name for tracks as POSIX strftime in the kml file
$trkName = "%Y-%m-%d %H%M%S" 

#----------------------------------------------------------
#region Path for programs
# Path for Columbus V-1000 conversion program for .GPS file
$rcconv = "F:\Simple Exe\GPS\RouteConverterCmdLine.jar"
if (!(Test-Path $rcconv)){
    Write-Host ("Cannot find RouteConverterCmdLine.jar at '{0}'" -f $rcconv)
    Exit(1)
}
# Path for java jre
$jre = "C:\Program Files\Java\jre1.8.0_73\bin\javaw.exe"
if (!(Test-Path $jre)){
    Write-Host ("Cannot find java.exe at '{0}'" -f $jre)
    Exit(1)
}
# Path for gpsbabel program
$gpsbabel = "F:\Program Files (x86)\GPSBabel\gpsbabel.exe" # Tested with Version 1.5.1
if (!(Test-Path $gpsbabel)){
    Write-Host ("Cannot find NVK.EXE at '{0}'" -f $gpsbabel)
    Exit(1)
}
else { & $gpsbabel -V }
# Path for Googel Earth program
$GEarth = "C:\Program Files (x86)\Google\Google Earth\client\googleearth.exe"
if (!(Test-Path $GEarth)){
    Write-Host ("Cannot find NVK.EXE at '{0}'" -f $GEarth)
    Exit(1)
}
#endregion

#----------------------------------------------------------
#region Move Columbus Binary files to root directory with complete datetime filename 
Write-Host 'Move & Rename Columbus Binary files' -ForegroundColor Yellow
$nbFiles = 0
Get-ChildItem -Directory | Sort-Object -Property Name | % {
    if ($_ -cmatch '(19|20)[0-9]{2}[-](0[1-9]|1[012])') {
        $dirname = $_.BaseName
        Write-Host $dirname -ForegroundColor Cyan
        Get-ChildItem $_ -File | Sort-Object -Property Name | % {
            if ($_ -cmatch '\.GPS$') {
                $filename = $dirname + '-' + $_.Name.Substring(0,2) + '_' + $_.Name.Substring(2,6) + '.GPS'
                Write-Host $_ ' -> ' $filename
                Move-Item $_.FullName $(Join-Path $CurrentDir $filename) -Force
                $nbFiles++
            }

        }
    }

}
Write-Host $nbFiles 'files moved' -ForegroundColor Yellow
#endregion

#----------------------------------------------------------
#region Convert columbus Binary file to Columbus type 2 cvs format
Write-Host "`nConvert files to Columbus Type 2 csv format" -ForegroundColor Yellow
$nbFiles = 0
Get-ChildItem -File -Path '*.GPS' | Sort-Object -Property Name | % {
    Write-Host $_.Name
    # Example: D:\Users\Hervé\Desktop\GPS\>javaw -jar RouteConverterCmdLine.jar "..\15111755.GPS" "ColumbusGpsType2Format" foo  
    $args = @('-jar',$rcconv,$_.FullName,"ColumbusGpsType2Format",$(Join-Path $_.DirectoryName $_.BaseName))        
    & $jre $args | Out-Null
    $nbFiles++
}
Write-Host $nbFiles 'files converted' -ForegroundColor Yellow
#endregion

#----------------------------------------------------------
#region Convert Columbus csv into gpx format
Write-Host "`nConvert files to GPX format" -ForegroundColor Yellow
$nbFiles = 0
Get-ChildItem -File -Path '*.csv' | Sort-Object -Property Name | % {
    $item = $_.BaseName
    Write-Host $_.Name
    # Example: D:\Users\Hervé\Desktop\GPS\>javaw -jar RouteConverterCmdLine.jar "..\15111755.GPS" "Gpx10Format" foo  
    $args = @('-jar',$rcconv,$_.FullName,"Gpx10Format",$(Join-Path $_.DirectoryName $_.BaseName))        
    & $jre $args | Out-Null
    $nbFiles++
    # rename track with file name
    (Get-Content $_.Name.Replace('csv','gpx')) | % {
        if ($_ -match '<name>Trackpoint [0-9]{1,9} to Trackpoint [0-9]{1,9}</name>') {
            $_ -replace '<name>Trackpoint [0-9]{1,9} to Trackpoint [0-9]{1,9}</name>',"<name>$item</name>"
        }
        else {
            $_
        }
    } | Out-File $_.Name.Replace('csv','gpx') -Encoding utf8

}

Write-Host $nbFiles 'files converted' -ForegroundColor Yellow
#endregion

#----------------------------------------------------------
#region Extract waypoints
$wptText = ',C,'
$nbWpt = 0
$tmpwptFile = 'tmp.csv'

Write-Host "`nExtract Waypoints" -ForegroundColor Yellow -NoNewline
if (Test-Path $tmpwptFile) { Remove-Item $tmpwptFile }

# Stop if no csv files are present
if ((Get-ChildItem -File -Path '*.csv').count -eq 0) {
    Write-Host "`nError: No csv file found" -ForegroundColor Red
    Exit(1)
}

# Header of csv file
$buffer = Get-Content (Get-ChildItem -Path '*.csv' | select -first 1) -Head 1

# Select waypoints from all csv files
Get-ChildItem -File -Path '*.csv' | Sort-Object -Property Name | % {
    Select-String -path $_ -pattern $wptText | % {
        $buffer += "`r`n" + $_.Line
        $nbWpt++
    }
}

# Only if waypoints have been found
if (!($nbWpt -eq 0)) {
    # write result into $tmpwptFile file
    Out-File -FilePath $tmpwptFile -InputObject $buffer -Encoding ascii -Force
    Write-Host ':' $nbWpt 'found' -ForegroundColor Yellow

    # Convert $tmpwptFile into gpx format. This creates a track file
    $args = @('-jar',$rcconv,$(Join-Path $CurrentDir $tmpwptFile),"Gpx10Format",$(Join-Path $CurrentDir $tmpwptFile.Replace('.csv','.gpx')))
    & $jre $args | Out-Null

    # Convert gpx track into gpx waypoint file 
    # See http://www.gpsbabel.org/htmldoc-development/filter_nuketypes.html
    #gpsbabel -i gpx -f track.gpx -x transform,wpt=trk -o gpx -F wpt.gpx
    $args = '-i','gpx','-f',$tmpwptFile.Replace('.csv','.gpx')
    $args += '-x','transform,wpt=trk,del'
    $args += '-o','gpx','-F',$wptFileName
    
    & $gpsbabel $args

    # remove timestamp and rename trackpoint
    $i=1
    (Get-Content $wptFileName) | % {
        if ($_ -match '<name>.*</name>') {
            $_ -replace '<name>.*</name>',"<name>wpt$i</name>"
            $i++
        }
        elseif ($_ -match '<cmt>.*</cmt>') { }
        elseif ($_ -match '<desc>.*</desc>') { }
        elseif ($_ -match '<time>.*</time>') { }
        else {
            $_
        }
    } | Out-File $wptFileName -Encoding utf8
    
    # delete tmp files
    Remove-Item $tmpwptFile
    Remove-Item $tmpwptFile.Replace('.csv','.gpx')
    Write-Host '-> ' $wptFileName 
}
#endregion

#----------------------------------------------------------
#region convert gpx files into kml file
Write-Host "`nGenerate Google Earth kml file" -ForegroundColor Yellow
$nbFiles = 0
$babelargs = @()
$babelargs += '-i','gpx'

Get-ChildItem -File -Path '*.gpx' | Sort-Object -Property Name | % {
    $babelargs += '-f'
    $babelargs += '"'+$_.Name+'"'
    $nbFiles++
}
# See http://www.gpsbabel.org/htmldoc-development/filter_nuketypes.html
$babelargs += '-x','nuketypes,routes'
# See http://www.gpsbabel.org/htmldoc-development/filter_simplify.html
$babelargs += '-x','simplify,crosstalk,error=0.001k'

# See http://www.gpsbabel.org/htmldoc-development/fmt_kml.html
$babelargs += '-o','kml,line_width=2,line_color=ff0000ff,points=0,trackdata=0'
$babelargs += '-F',"$kmlFileName"


# if some files have to be processed
if (!($nbFiles -eq 0)) {
    # Start GpsBabel with $batch command line
    # $batch format is "-i gpx -f f1 -f f2 -i nmea -f f3 -f f4 -x filter -o kml -F f5
    & $gpsbabel $babelargs
    Write-Host '-> ' $kmlFileName
}
#endregion

#----------------------------------------------------------
#region Manipulate kml file
# By default gpsbabel define a name for the track folder, but not for the Path itself
# This portion of the script renames the path to math the folder name
#     <Folder>
#     <name>20101002-121918</name> -> is context.PreContext[0]
#     <Placemark>
#     <name>Path</name> -> <name>20101002-121918</name>
# if some files have to be processed
if (!($nbFiles -eq 0)) {
    Write-Host 'Rename ' -ForegroundColor Green -NoNewline
    $ss = Select-String -Path $kmlFileName -Pattern '<name>Path</name>' -Context 2
    Write-Host $ss.count -ForegroundColor Green -NoNewline

    # Parse the $kml file and replace Path name by its PreContext[0] ie the string 2 lines above
    $i=0
    (Get-Content $kmlFileName) | % {
        if ($_ -match '<name>Path</name>') {
            $_ -replace '<name>Path</name>',$ss[$i++].context.PreContext[0]
        }
        else {
            $_
        }
    } | Out-File $kmlFileName -Encoding utf8
    Write-Host ' Tracks in KML file' -ForegroundColor Green
}
#endregion

#----------------------------------------------------------
#region cleanup files
Remove-Item -Path '*.csv'
Remove-Item -Path '*.gpx' -Exclude $wptFileName
#endregion

#----------------------------------------------------------
#region Open Google Earth
# Start Google Earth
if ((Test-Path $kmlFileName) -and ((Get-Item $kmlFileName).Length -gt 0)) {
    & $GEarth $(Get-ChildItem $kmlFileName).FullName
}
#endregion
