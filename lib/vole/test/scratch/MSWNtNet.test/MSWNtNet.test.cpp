/* /////////////////////////////////////////////////////////////////////////
 * File:        MSWNtNet.test.cpp
 *
 * Purpose:     Implementation file for the MSWNtNet.test project.
 *
 * Created:     4th January 2007
 * Updated:     1st February 2011
 *
 * Status:      Wizard-generated
 *
 * License:     (Licensed under the Synesis Software Open License)
 *
 *              Copyright (c) 2007-2011, Synesis Software Pty Ltd.
 *              All rights reserved.
 *
 *              www:        http://www.synesis.com.au/software
 *
 * ////////////////////////////////////////////////////////////////////// */


/* VOLE Header Files */
#include <vole/vole.hpp>

/* STLSoft Header Files */
#include <stlsoft/smartptr/ref_ptr.hpp>
#include <comstl/error/errorinfo_desc.hpp>
#include <comstl/string/bstr.hpp>
#include <comstl/util/creation_functions.hpp>
#include <comstl/util/initialisers.hpp>

/* Standard C++ Header Files */
#include <exception>
#include <iostream>

using std::cerr;
using std::cin;
using std::cout;
using std::wcout;
using std::endl;

/* Standard C Header Files */
#include <stdlib.h>

#if defined(_MSC_VER) && \
    defined(_DEBUG)
# include <crtdbg.h>
#endif /* _MSC_VER) && _DEBUG */

/* /////////////////////////////////////////////////////////////////////////
 * main()
 */

static int main_(int /* argc */, char ** /*argv*/)
{
    using vole::object;
    using vole::collection;
    using vole::of_type;

    object  ServiceManager  =   object::create("SynesisSoftware.WindowsNt.ServiceManager");

    long    timeout         =   ServiceManager.get_property(of_type<long>(), L"Timeout");

    object  Services        =   ServiceManager.get_property(of_type<object>(), L"Services");

    long    numServices     =   Services.get_property(of_type<long>(), L"Count");

    { for(long i = 0; i < numServices; ++i)
    {
        object          Service                 =   Services.get_property(of_type<object>(), L"Item", i);
        comstl::bstr    ServiceName             =   Service.get_property(of_type<comstl::bstr>(), L"ServiceName");
        comstl::bstr    DisplayName             =   Service.get_property(of_type<comstl::bstr>(), L"DisplayName");
        collection      DependentServices       =   Service.get_property(of_type<object>(), L"DependentServices");
        long            numDependentServices    =   DependentServices.get_Count();

        wcout << L"[" << ServiceName << L"] " << DisplayName << endl;

        if(0 != numDependentServices)
        {
            wcout   << L"\t"
                    << numDependentServices
                    << L" dependent service"
                    << ((1 != numDependentServices) ? L"s:" : L":")
                    << endl;

#if 1
            { for(collection::iterator b = DependentServices.begin(); b != DependentServices.end(); ++b)
            {
                object          DependentService    =   object(*b);
#else /* ? 0 */
            { for(long i = 0; i < numDependentServices; ++i)
            {
                object          DependentService    =   DependentServices[i];
#endif /* 0 */
                comstl::bstr    ServiceName         =   DependentService.get_property(of_type<comstl::bstr>(), L"ServiceName");
                comstl::bstr    DisplayName         =   DependentService.get_property(of_type<comstl::bstr>(), L"DisplayName");

                wcout << L"\t\t[" << ServiceName << L"] " << DisplayName << endl;
            }}
        }

        wcout << endl;

#if 0
        enumerator      en                  =   DependepentServices.get_NewEnum();

        { for(enumerator::iterator b = en.begin(); b != en.end(); ++b)
        {
            object          DependentService    =   *b;
            comstl::bstr    ServiceName         =   Service.get_property(of_type<comstl::bstr>(), L"ServiceName");
            comstl::bstr    DisplayName         =   Service.get_property(of_type<comstl::bstr>(), L"DisplayName");

            wcout << L"\t[" << ServiceName << L"] " << DisplayName << endl << endl;
        }}
#endif /* 0 */
    }}

    STLSOFT_SUPPRESS_UNUSED(timeout);

    return EXIT_SUCCESS;
}

int main(int argc, char *argv[])
{
    int             res;

#if defined(_MSC_VER) && \
    defined(_DEBUG)
    _CrtMemState    memState;
#endif /* _MSC_VER && _MSC_VER */

#if defined(_MSC_VER) && \
    defined(_DEBUG)
    _CrtMemCheckpoint(&memState);
#endif /* _MSC_VER && _MSC_VER */

#if 0
    { for(size_t i = 0; i < 0xffffffff; ++i){} }
#endif /* 0 */

    try
    {
#if defined(_DEBUG) || \
    defined(__SYNSOFT_DBS_DEBUG)
        cout << "MSWNtNet.test: " << STLSOFT_COMPILER_LABEL_STRING << endl;
#endif /* debug */

        comstl::com_initialiser coinit;

        res = main_(argc, argv);
    }
    catch(std::exception &x)
    {
        std::cerr << "Unexpected general error: " << x.what() << ". Application terminating" << std::endl;

        res = EXIT_FAILURE;
    }
    catch(...)
    {
        std::cerr << "Unhandled unknown error" << std::endl;

        res = EXIT_FAILURE;
    }

#if defined(_MSC_VER) && \
    defined(_DEBUG)
    _CrtMemDumpAllObjectsSince(&memState);
#endif /* _MSC_VER) && _DEBUG */

    return res;
}

/* ///////////////////////////// end of file //////////////////////////// */
