#include <ctype.h>
#include <windows.h>
#include "XLCALL.H"
#include "FRAMEWRK.H"
#include "proj_api.h"
#include "util.h"
#include "epsg.h"

#define rgWorksheetFuncsRows 3
#define rgWorksheetFuncsCols 14

static LPWSTR rgWorksheetFuncs[rgWorksheetFuncsRows][rgWorksheetFuncsCols] =
{
    {
        L"projVersion",                         // LPXLOPER12 pxProcedure
        L"UU",                                  // LPXLOPER12 pxTypeText
        L"PROJ.VERSION",                        // LPXLOPER12 pxFunctionText
        L"",                                    // LPXLOPER12 pxArgumentText
        L"1",                                   // LPXLOPER12 pxMacroType
        L"PROJ.4",                              // LPXLOPER12 pxCategory
        L"",                                    // LPXLOPER12 pxShortcutText
        L"",                                    // LPXLOPER12 pxHelpTopic
        L"Returns the PROJ.4 library version.", // LPXLOPER12 pxFunctionHelp
        L"",                                    // LPXLOPER12 pxArgumentHelp1
        L"",                                    // LPXLOPER12 pxArgumentHelp2
        L"",                                    // LPXLOPER12 pxArgumentHelp3
        L"",                                    // LPXLOPER12 pxArgumentHelp4
        L""                                     // LPXLOPER12 pxArgumentHelp5
    },
	{
		L"projTransform",
		L"UCCBBH",
		L"PROJ.TRANSFORM",
		L"",
		L"1",
		L"PROJ.4",
		L"",
		L"",
		L"Transform X / Y points from source coordinate system to destination coordinate system.", 
		L"Source coordinate system",
		L"Destination coordinate system",
		L"X coordinate",
		L"Y coordinate",
		L"Output flag: 1 = X, 2 = Y"
	},
	{
		L"projEPSG",
		L"UJ",
		L"EPSG",
		L"",
		L"1",
		L"PROJ.4",
		L"",
		L"",
		L"Returns the PROJ.4 string associated with an EPSG code.",
		L"EPSG code",
		L"",
		L"",
		L"",
		L""
	}
};

/*
** Standard XLL functions:
** - xlAutoOpen
** - xlAutoClose
** - xlAutoRegister12
** - xlAutoAdd
** - xlAutoRemove
** - xlAddInManagerInfo12
**
** UDFs:
** - projTransform
** - projVersion
*/

// Excel calls xlAutoOpen when it loads the XLL.
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    static XLOPER12 xDLL;   /* name of this DLL */
    int i;                  /* Loop index */

    /*
    ** In the following block of code the name of the XLL is obtained by
    ** calling xlGetName. This name is used as the first argument to the
    ** REGISTER function to specify the name of the XLL. Next, the XLL loops
    ** through the rgFuncs[] table, registering each function in the table using
    ** xlfRegister. Functions must be registered before you can add a menu
    ** item.
    */

    Excel12f(xlGetName, &xDLL, 0);

    for (i=0; i < rgWorksheetFuncsRows; i++) 
    {
        Excel12f(xlfRegister, 0, 1 + rgWorksheetFuncsCols,
            (LPXLOPER12)&xDLL,
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][0]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][1]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][2]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][3]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][4]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][5]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][6]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][7]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][8]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][9]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][10]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][11]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][12]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][13]));
    }

    /* Free the XLL filename */
    Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);
    
    return 1;
}

// Excel calls xlAutoClose when it unloads the XLL.
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    int i;

     // Delete all names added by xlAutoOpen or xlAutoRegister.
    for (i = 0; i < rgWorksheetFuncsRows; i++)
    {
        Excel12f(xlfSetName, 0, 1, TempStr12(rgWorksheetFuncs[i][2]));
    }

    return 1;
}


// Excel calls xlAutoRegister12 if a macro sheet tries to register
// a function without specifying the type_text argument.
__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
    static XLOPER12 xDLL, xRegId;
    int i;

    /*
    ** This block initializes xRegId to a #VALUE! error first. This is done in
    ** case a function is not found to register. Next, the code loops through the
    ** functions in rgFuncs[] and uses lpstricmp to determine if the current
    ** row in rgFuncs[] represents the function that needs to be registered.
    ** When it finds the proper row, the function is registered and the
    ** register ID is returned to Microsoft Excel. If no matching function is
    ** found, an xRegId is returned containing a #VALUE! error.
    */

    xRegId.xltype = xltypeErr;
    xRegId.val.err = xlerrValue;

    for (i = 0; i < rgWorksheetFuncsRows; i++) 
    {
        if (!lpwstricmp(rgWorksheetFuncs[i][0], pxName->val.str)) 
        {
            Excel12f(xlGetName, &xDLL, 0);

            Excel12f(xlfRegister, 0, 4,
                (LPXLOPER12)&xDLL,
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][0]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][1]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][2]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][3]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][4]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][5]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][6]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][7]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][8]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][9]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][10]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][11]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][12]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][13]));
                
            /* Free the XLL filename */
            Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

            return (LPXLOPER12)&xRegId;
        }
    }

    //Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
    //for UDFs declared as thread safe, use alternate memory allocation mechanisms

    return (LPXLOPER12)&xRegId;
}

// When you add an XLL to the list of active add-ins, the Add-in
// Manager calls xlAutoAdd() and then opens the XLL, which in turn
// calls xlAutoOpen.
__declspec(dllexport) int WINAPI xlAutoAdd(void)
{
    return 1;
}

// When you remove an XLL from the list of active add-ins, the
// Add-in Manager calls xlAutoRemove() and then
// UNREGISTER("SAMPLE.XLL").
__declspec(dllexport) int WINAPI xlAutoRemove(void)
{
    return 1;
}

// ----------------------------------------------------------------------------
// PROJ.VERSION
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projVersion(LPXLOPER12 x)
{
    static XLOPER12 xResult;

    xResult.xltype = xltypeStr;
    xResult.val.str = new_xl12string(pj_get_release());
    
    return (LPXLOPER12)&xResult;
}

// ----------------------------------------------------------------------------
// PROJ.TRANSFORM
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projTransform(const char* src, const char* dst, const double x, const double y, const WORD type)
{
    static XLOPER12 xResult;

    projPJ proj_src, proj_dst;
	proj_src = pj_init_plus(src);
	proj_dst = pj_init_plus(dst);

    if (!proj_src || !proj_dst)
    {
		xResult.xltype = xltypeErr;
        xResult.val.err = xlerrValue;
        return &xResult;
    }

    double x1 = x;
    double y1 = y;

    if (pj_transform(proj_src, proj_dst, 1, 1, &x1, &y1, NULL) == 0)
    {
		if (type == 1) {
			xResult.xltype = xltypeNum;
			xResult.val.num = x1;
		}
		else if (type == 2) {
			xResult.xltype = xltypeNum;
			xResult.val.num = y1;
		}
		else {
			xResult.xltype = xltypeErr;
			xResult.val.err = xlerrValue; // Invalid argument
		}
    }
    else
    {
		xResult.xltype = xltypeErr;
        xResult.val.err = xlerrNum; // Error in pj_transform
    }
    
    if (proj_src != NULL)
        pj_free(proj_src);
    if (proj_dst != NULL)
        pj_free(proj_dst);

    return (LPXLOPER12)&xResult;
}

// ----------------------------------------------------------------------------
// EPSG
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projEPSG(const int code)
{
	static XLOPER12 xResult;
	
	wchar_t *projStr = epsgLookup(code);

	if (projStr != NULL) {
		xResult.xltype = xltypeStr;
		xResult.val.str = projStr;
	}
	else {
		xResult.xltype = xltypeErr;
		xResult.val.err = xlerrNA; // No value available
	}

	return (LPXLOPER12)&xResult;
}
