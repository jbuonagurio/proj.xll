proj.xll
========

Transform coordinates between various map projections directly within Excel.

This program uses the [PROJ.4 Cartographic Projections Library](http://proj.osgeo.org/).

Installation
------------

Copy proj.xll to `%USERPROFILE%\AppData\Roaming\Microsoft\AddIns\`.

Go to Excel Options, Add-ins, Manage Excel Add-ins. Click Browse, and select proj.xll.

Usage
-----

To display the PROJ.4 library version:

```
=PROJ.VERSION()
```

To transform coordinates, for example from latitude/longitude to UTM zone 17:
```
=PROJ.TRANSFORM("+proj=latlong +datum=WGS84","+proj=utm +zone=17 +datum=WGS84",RADIANS(-80),RADIANS(40),1)
=PROJ.TRANSFORM("+proj=latlong +datum=WGS84","+proj=utm +zone=17 +datum=WGS84",RADIANS(-80),RADIANS(40),2)
```

Arguments:
 
1. Source coordinate system
2. Destination coordinate system
3. X value
4. Y value
5. Output: 1=X, 2=Y

Latitude and longitude values must be in radians.

EPSG codes are also supported using the `EPSG()` helper function. For example, to transform from latitude/longitude to Web Mercator:
```
=PROJ.TRANSFORM(EPSG(4326),EPSG(3857),RADIANS(-89.83152841),RADIANS(40.91627447),1)
=PROJ.TRANSFORM(EPSG(4326),EPSG(3857),RADIANS(-89.83152841),RADIANS(40.91627447),2)
```

Development
-----------

Download the [PROJ.4 source](http://proj4.org/download.html), tested with v4.9.3 (2016-09-02). Extract to C:\PROJ.

Edit the PROJ.4 OPTFLAGS in nmake.opt as follows, to use the static multithreaded version of the CRT library:

```
!IFNDEF OPTFLAGS
!IFNDEF DEBUG
OPTFLAGS=	/nologo /Ox /Op /MT
!ELSE
OPTFLAGS=	/nologo /Zi /MTd /Fdproj.pdb
!ENDIF
!ENDIF
```

Build the PROJ.4 static library from a Visual Studio command prompt:

```
C:\PROJ>nmake /f makefile.vc
C:\PROJ>nmake /f makefile.vc install-all
```

Download and install the [Excel 2013 SDK](https://www.microsoft.com/en-us/download/details.aspx?id=35567) to the default location. Build the FRAMEWORK library from a Visual Studio command prompt:

```
C:\2013 Office System Developer Resources\Excel2013XLLSDK\SAMPLES>MAKE DEBUG
```

You should now be able to build the XLL using either the Visual Studio project or nmake makefile.
