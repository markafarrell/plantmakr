/* /////////////////////////////////////////////////////////////////////////
 * File:        pantheios.com.test.cpp
 *
 * Purpose:     Implementation file for the pantheios_com_test project.
 *
 * Created:     3rd January 2007
 * Updated:     9th August 2012
 *
 * Status:      Wizard-generated
 *
 * License:     (Licensed under the Synesis Software Open License)
 *
 *              Copyright (c) 2007-2012, Synesis Software Pty Ltd.
 *              All rights reserved.
 *
 *              www:        http://www.synesis.com.au/software
 *
/* /////////////////////////////////////////////////////////////////////////

//#define _COMSTL_NO_UUIDOF


/* VOLE Header Files */
#include <vole/vole.hpp>

/* STLSoft Header Files */
#include <stlsoft/smartptr/ref_ptr.hpp>
#include <comstl/collections/enumerator_sequence.hpp>
#include <comstl/conversion/interface_cast.hpp>
#include <comstl/error/errorinfo_desc.hpp>
#include <comstl/util/creation_functions.hpp>
#include <comstl/util/initialisers.hpp>
#include <comstl/util/value_policies.hpp>
#include <winstl/error/error_desc.hpp>

/* Standard C++ Header Files */
#include <exception>
#include <iostream>
#include <vector>

using std::cerr;
using std::cin;
using std::cout;
using std::endl;

/* Standard C Header Files */
#include <stdlib.h>

#if defined(_MSC_VER) && \
    defined(_DEBUG)
# include <crtdbg.h>
#endif /* _MSC_VER) && _DEBUG */

/* /////////////////////////////////////////////////////////////////////////
 * Constants & definitions
 */

static const wchar_t PROCESS_IDENTITY[] = L"VOLE-pantheios.COM-test";

/* /////////////////////////////////////////////////////////////////////////
 * Forward declarations
 */

static void pantheios_COM_test_manually();
static void pantheios_COM_test_with_VOLE();

/* ////////////////////////////////////////////////////////////////////// */

static int main_(int /* argc */, char ** /*argv*/)
{
    {
        using vole::object;

        object obj;

        try
        {
            object().swap(obj);

            obj.get();

            obj.is_nothing();

            obj.lcid();

            obj.invoke_method_v(L"Xxx");

        }
        catch(vole::vole_exception& x)
        {
            std::string message(x.what());

            message = "";
        }
    }

    pantheios_COM_test_manually();
    pantheios_COM_test_with_VOLE();

    return EXIT_SUCCESS;
}

int main(int argc, char *argv[])
{
    int             res;

#if defined(_MSC_VER) && \
    defined(_DEBUG)
    _CrtMemState    memState;
#endif /* _MSC_VER && _MSC_VER */

    cout << "";

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
        cout << "pantheios_com_test: " << STLSOFT_COMPILER_LABEL_STRING << endl;
#endif /* debug */

        comstl::com_initialiser coinit;

        res = main_(argc, argv);
    }
    catch(std::exception &x)
    {
        cerr << "Unhandled error: " << x.what() << endl;

        res = EXIT_FAILURE;
    }
    catch(...)
    {
        cerr << "Unhandled unknown error" << endl;

        res = EXIT_FAILURE;
    }

#if defined(_MSC_VER) && \
    defined(_DEBUG)
    _CrtMemDumpAllObjectsSince(&memState);
#endif /* _MSC_VER) && _DEBUG */

    return res;
}

/* ////////////////////////////////////////////////////////////////////// */

static void pantheios_COM_test_manually()
{
    CLSID   clsid;
    HRESULT hr  =   ::CLSIDFromProgID(L"pantheios.COM.LoggerManager", &clsid);

    if(FAILED(hr))
    {
        cerr << "Could not find class id of pantheios.COM.LoggerManager" << endl;
    }
    else
    {
        IDispatch   *pdisp;

        hr = ::CoCreateInstance(clsid, NULL, CLSCTX_ALL, IID_IDispatch, reinterpret_cast<void**>(&pdisp));

        if(FAILED(hr))
        {
            cerr << "Could not create an instance of the pantheios.COM.LoggerManager class" << endl;
        }
        else
        {
            LPOLESTR    names1[1] = { L"GetLogger" };
            DISPID      ids1[1];

            hr = pdisp->GetIDsOfNames(IID_NULL, &names1[0], 1, LOCALE_SYSTEM_DEFAULT, &ids1[0]);

            if(FAILED(hr))
            {
                cerr << "GetLogger() method not available on logger manager" << endl;
            }
            else
            {
                VARIANT     args1[3];
                DISPPARAMS  params1 =   { args1, NULL, 3, 0 };
                VARIANT     result1;
                UINT        argErr1;

                ::VariantInit(&args1[0]);
                ::VariantInit(&args1[1]);
                ::VariantInit(&args1[2]);

                // Arguments are backwards
                args1[2].vt         =   VT_BSTR;
                args1[2].bstrVal    =   ::SysAllocString(L"Console");
                args1[1].vt         =   VT_BSTR;
                args1[1].bstrVal    =   ::SysAllocString(PROCESS_IDENTITY);

                hr = pdisp->Invoke(ids1[0], IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params1, &result1, NULL, &argErr1);

                if(FAILED(hr))
                {
                    cerr << "Could not invoke GetLogger() method: " << winstl::error_desc(hr).c_str() << ": " << comstl::errorinfo_desc() << endl;
                }
                else
                {
                    if( VT_DISPATCH != result1.vt ||
                        NULL == result1.pdispVal)
                    {
                        cerr << "GetLogger() returned an invalid result" << endl;
                    }
                    else
                    {
                        if(SUCCEEDED(hr))
                        {
                            LPOLESTR    names2[1] = { L"Log" };
                            DISPID      ids2[1];

                            hr = result1.pdispVal->GetIDsOfNames(IID_NULL, &names2[0], 1, LOCALE_SYSTEM_DEFAULT, &ids2[0]);

                            if(FAILED(hr))
                            {
                                cerr << "Log() method not available on logger manager" << endl;
                            }
                            else
                            {
                                VARIANT     args2[3];
                                DISPPARAMS  params2 =   { args2, NULL, 3, 0 };
                                VARIANT     result2;
                                UINT        argErr2;

                                ::VariantInit(&args2[0]);
                                ::VariantInit(&args2[1]);

                                // Arguments are backwards
                                args2[2].vt         =   VT_I4;
                                args2[2].lVal       =   5;
                                args2[1].vt         =   VT_BSTR;
                                args2[1].bstrVal    =   ::SysAllocString(L"The answer is ");
                                args2[0].vt         =   VT_I4;
                                args2[0].lVal       =   42;

                                hr = result1.pdispVal->Invoke(ids2[0], IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params2, &result2, NULL, &argErr2);

                                if(FAILED(hr))
                                {
                                    cerr << "Could not invoke Log() method: " << winstl::error_desc(hr).c_str() << ": " << comstl::errorinfo_desc() << endl;
                                }
                                else
                                {
                                    cout << "Log() operation completed successfully" << endl;

                                    ::VariantClear(&result2);
                                }
                            }
                        }
                    }

                    ::VariantClear(&result1);
                }

                ::VariantClear(&args1[0]);
                ::VariantClear(&args1[1]);
                ::VariantClear(&args1[2]);
            }

            pdisp->Release();
        }
    }
}

static void pantheios_COM_test_with_VOLE()
{
    try
    {
        using vole::collection;
        using vole::object;
        using vole::of_type;

        object  loggerManager   =   object::create("pantheios.COM.LoggerManager");
        object  logger          =   loggerManager.invoke_method(of_type<object>(), L"GetLogger", L"Console", PROCESS_IDENTITY);

#if 1
        logger.invoke_method_v(L"Log", 3, L"The answer is: ", 43);
#else /* ? 0 */
        logger.invoke_method<void>(L"Log", 3, L"The answer is: ", 43);
#endif /* 0 */

        std::string processId       =   logger.get_property(of_type<std::string>(), L"ProcessIdentity");

        long        backEndId       =   logger.get_property(of_type<long>(), L"BackEndId");

        logger.invoke_method_v(L"Log", 4, "abc(", L"DEF", "): ", 2319, " - ", 19.19, L" - ", std::string("yada!").c_str());

        // Aliases
        typedef comstl::enumerator_sequence<IEnumVARIANT
                                        ,   VARIANT
                                        ,   comstl::VARIANT_policy
                                        ,   VARIANT const &
                                        ,   comstl::input_cloning_policy<IEnumVARIANT>
                                        >           enumerator_t;

        collection      aliases =   loggerManager.get_property(of_type<object>(), L"KnownLoggerAliases");

        enumerator_t    en(aliases.get__NewEnum().get(), true);

        { for(enumerator_t::iterator b = en.begin(); b != en.end(); ++b)
        {
            cout << "\t" << stlsoft::c_str_ptr(*b) << endl;
        }}

        STLSOFT_SUPPRESS_UNUSED(backEndId);
    }
    catch(comstl::comstl_exception &x)
    {
        cerr << "COM error: " << x.what() << ": " << winstl::error_desc(x.hr()).c_str() << endl;
    }
    catch(std::exception &x)
    {
        cerr << x.what() << endl;
    }
}

/* ///////////////////////////// end of file //////////////////////////// */
