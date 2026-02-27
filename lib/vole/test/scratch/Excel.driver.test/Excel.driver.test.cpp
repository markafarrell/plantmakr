// Excel.driver.test.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#include <vole/vole.hpp>

#include <comstl/util/initialisers.hpp>
#include <comstl/util/variant.hpp>
#include <winstl/error/error_desc.hpp>

#include <iostream>

#include <stdlib.h>

using std::cout;
using std::cerr;
using std::endl;

int main()
{
    try
    {
        comstl::ole_init    coinit;

        using vole::object;
        using vole::collection;
        using vole::of_type;

        // Start server ...
        object  xlApp   =   object::create(L"Excel.Application", CLSCTX_LOCAL_SERVER, vole::coercion_level::naturalPromotion);

        // Make it visible (i.e. app.visible = 1)
        xlApp.put_property(L"Visible", 1);

#ifdef VOLE_METHOD_TEMPLATE_RETURN_SUPPORT

        // Get Workbooks collection
        collection  xlBooks =   xlApp.get_property<collection>(L"Workbooks");

        // Call Workbooks.Add() to get a new workbook...
        object      xlBook  =   xlBooks.get_property<object>(L"Add");

        // Create a 15x15 safearray of variants...
        comstl::variant arr;
        arr.vt = VT_ARRAY | VT_VARIANT;
        {
        SAFEARRAYBOUND sab[2];
        sab[0].lLbound = 1; sab[0].cElements = 15;
        sab[1].lLbound = 1; sab[1].cElements = 15;
        arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
        }

        // Fill safearray with some values...
        for(int i=1; i<=15; i++) {
        for(int j=1; j<=15; j++) {
            // Create entry value for (i,j)
            VARIANT tmp;
            tmp.vt = VT_I4;
            tmp.lVal = i*j;
            // Add to safearray...
            long indices[] = {i,j};
            SafeArrayPutElement(arr.parray, indices, (void *)&tmp);
        }
        }

        // Get ActiveSheet object
        object  xlSheet =   xlApp.get_property<object>(L"ActiveSheet");

        // Get Range object for the Range A1:O15...
        object  xlRange =   xlSheet.get_property<object>(L"Range", "A1:O15");

#else /* ? VOLE_METHOD_TEMPLATE_RETURN_SUPPORT */

        // Get Workbooks collection
        collection  xlBooks =   xlApp.get_property(of_type<collection>(), L"Workbooks");

        // Call Workbooks.Add() to get a new workbook...
        object      xlBook  =   xlBooks.get_property(of_type<object>(), L"Add");

        // Create a 15x15 safearray of variants...
        comstl::variant arr;
        arr.vt = VT_ARRAY | VT_VARIANT;
        {
        SAFEARRAYBOUND sab[2];
        sab[0].lLbound = 1; sab[0].cElements = 15;
        sab[1].lLbound = 1; sab[1].cElements = 15;
        arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
        }

        // Fill safearray with some values...
        for(int i=1; i<=15; i++) {
        for(int j=1; j<=15; j++) {
            // Create entry value for (i,j)
            VARIANT tmp;
            tmp.vt = VT_I4;
            tmp.lVal = i*j;
            // Add to safearray...
            long indices[] = {i,j};
            SafeArrayPutElement(arr.parray, indices, (void *)&tmp);
        }
        }

        // Get ActiveSheet object
        object  xlSheet =   xlApp.get_property(of_type<object>(), L"ActiveSheet");

        // Get Range object for the Range A1:O15...
        object  xlRange =   xlSheet.get_property(of_type<object>(), L"Range", "A1:O15");

#endif /* VOLE_METHOD_TEMPLATE_RETURN_SUPPORT */

        // Set range with our safearray...
        xlRange.put_property(L"Value", arr);

        // Wait for user...
        ::MessageBox(NULL, "All done.", "Notice", 0x10000);

        // Set .Saved property of workbook to TRUE so we aren't prompted
        // to save when we tell Excel to quit...
        xlBook.put_property(L"Saved", 1);

        // Tell Excel to quit (i.e. App.Quit)
        xlApp.invoke_method_v(L"Quit");

        return EXIT_SUCCESS;
    }
    catch(vole::vole_exception &x)
    {
        cout << "error: " << x.what() << ": " << winstl::error_desc(x.hr()).c_str() << endl;

        return EXIT_FAILURE;
    }
    catch(std::exception &x)
    {
        cout << "error: " << x.what() << endl;

        return EXIT_FAILURE;
    }
}

/* ///////////////////////////// end of file //////////////////////////// */
