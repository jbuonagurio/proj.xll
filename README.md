proj.xll
========

Excel add-in for PROJ.4 routines.

Usage
-----

To display the PROJ.4 library version:

```
=PROJ.VERSION()
```

To transform coordinates:
```
=PROJ.TRANSFORM("+proj=latlong +datum=WGS84","+proj=utm +zone=17 +datum=NAD83",RADIANS(-80),RADIANS(40),1)
```

Arguments:
 
1. Source coordinate system
2. Destination coordinate system
3. X value
4. Y value
5. Output: 1=X, 2=Y

Latitude and longitude values must be in radians.

Development
-----------

[Download the PROJ.4 source](http://proj4.org/download.html), tested with v4.9.3 (2016-09-02).

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



