/* /////////////////////////////////////////////////////////////////////////
 * File:        WordDriver.cpp
 *
 * Purpose:     Implementation file for the WordDriver project.
 *
 * Created:     12th January 2007
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
#include <comstl/util/creation_functions.hpp>
#include <comstl/util/initialisers.hpp>
#include <winstl/error/error_desc.hpp>

/* Standard C++ Header Files */
#include <exception>
#include <iostream>

/* Standard C Header Files */
#include <stdlib.h>

#if defined(_MSC_VER) && \
    defined(_DEBUG)
# include <crtdbg.h>
#endif /* _MSC_VER) && _DEBUG */

/* /////////////////////////////////////////////////////////////////////////
 * Forward declarations
 */

static void usage(int bExit, char const *reason, int invalidArg, int argc, char **argv);

/* /////////////////////////////////////////////////////////////////////////
 * main()
 */

static int main_(int argc, char *argv[])
{
    int bVerbose = true;

    { for(int i = 1; i < argc; ++i)
    {
        char const  *const  arg =   argv[i];

        if(arg[0] == '-')
        {
            if(arg[1] == '-')
            {
                /* -- arguments */
                usage(1, "Invalid argument(s) specified", i, argc, argv);
            }
            else
            {
                /* - arguments */
                switch(arg[1])
                {
                    case    '?':
                        usage(1, NULL, -1, argc, argv);
                        break;
                    case    's':
                        bVerbose    =   false;
                        break;
                    case    'v':
                        bVerbose    =   true;
                        break;
                    default:
                        usage(1, "Invalid argument(s) specified", i, argc, argv);
                        break;
                }
            }
        }
        else
        {
            /* other arguments */
            usage(1, "Invalid argument(s) specified", i, argc, argv);
        }
    }}

    // This is to be compared with Microsoft's Q238393 (http://support.microsoft.com/kb/238393)
    {
        using vole::object;
        using vole::collection;
        using vole::of_type;


        // Q238393: Get CLSID for Word.Application ...

        // VOLE: Don't need to do that



        // Q238393: Start Word and get IDispatch...

        // VOLE: Just create in a vole::object. No need to get raw IDispatch pointer.
        //       Note that we must explicitly specify CLSCTX_LOCAL_SERVER, just as
        //       in Q238393
        object  word    =   object::create("Word.Application", CLSCTX_LOCAL_SERVER, vole::coercion_level::naturalPromotion);

        try
        {
            // Q238393: Make Word visible

            word.put_property(L"Visible", true);



            // Q238393: Get Documents collection

            // VOLE: Rather than dealing with raw pointers, we just return an instance
            //       of vole::collection.

            collection  documents   =   word.get_property(of_type<collection>(), L"Documents");



            // Q238393: Call Documents.Open() to open C:\Doc1.doc

            // VOLE: Invoke the method, returning an instance of vole::object that holds
            //       the document object.

            object      document    =   documents.invoke_method(of_type<object>(), L"Open", L"C:\\Doc1.doc");


            try
            {

                // Q238393: Get BuiltinDocumentProperties collection

                collection  builtInProps =  document.get_property(of_type<collection>(), L"BuiltinDocumentProperties");


#ifdef _DEBUG
                int     numProps1   =   builtInProps.get_Count();
                double  numProps2   =   builtInProps.get_property(of_type<double>(), L"Count");

                STLSOFT_SUPPRESS_UNUSED(numProps1);
                STLSOFT_SUPPRESS_UNUSED(numProps2);
#endif /* _DEBUG */


                // Q238393: Get "Subject" from BuiltInDocumentProperties.Item("Subject")

                object      propSubject =   builtInProps.get_property(of_type<object>(), L"Item", L"Subject");



                // Q238393: Get the Value of the Subject property and display it

                std::string subject     =   propSubject.get_property(of_type<std::string>(), L"Value");

                std::cout << "Subject property: \"" << subject << '"' << std::endl;



                // Q238393: Set the Value of the Subject DocumentProperty

                propSubject.put_property(L"Value", "This is my subject");



                // Q238393: Get CustomDocumentProperties collection

                collection  customProps =   document.get_property(of_type<collection>(), L"CustomDocumentProperties");



                // Q238393: Add a new property named "CurrentYear"

                customProps.invoke_method_v(L"Add", L"CurrentYear", false, 1, 1999);



                // Q238393: Get the custom property "CurrentYear" and delete it

                object      propCurrYear =  customProps.get_property(of_type<object>(), L"Item", L"CurrentYear");

                propCurrYear.invoke_method_v(L"Delete");



                // Q238393: Close the document without saving changes and quit Word

                document.invoke_method_v(L"Close", false);

                word.invoke_method_v(L"Quit");
            }
            catch(comstl::comstl_exception& )
            {
                try
                {
                    document.invoke_method_v(L"Close", false);
                }
                catch(...)
                {}

                throw;
            }
        }
        catch(comstl::comstl_exception& )
        {
            word.invoke_method_v(L"Quit");

            throw;
        }
    }

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
        std::cout << "WordDriver: " << STLSOFT_COMPILER_LABEL_STRING << std::endl;
#endif /* debug */

//        comstl::com_initialiser coinit;
        comstl::ole_initialiser oleinit;

        res = main_(argc, argv);
    }
    catch(comstl::comstl_exception &x)
    {
        std::cerr << "COM error: " << x.what() << ": " << winstl::error_desc(x.hr()) << std::endl;

        res = EXIT_FAILURE;
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

/* ////////////////////////////////////////////////////////////////////// */

static void usage(int bExit, char const *reason, int invalidArg, int argc, char **argv)
{
    std::cerr << std::endl;
    std::cerr << "  WordDriver" << std::endl << std::endl;
    std::cerr << std::endl;

    if(NULL != reason)
    {
        std::cerr << "  Error: " << reason << std::endl;
        std::cerr << std::endl;
    }

    if(0 < invalidArg)
    {
        std::cerr << "  First invalid argument #" << invalidArg << ": " << argv[invalidArg] << std::endl;
        std::cerr << "  Arguments were (first bad marked *):" << std::endl << std::endl;
        { for(int i = 1; i < argc; ++i)
        {
            std::cerr << "  " << ((i == invalidArg) ? "* " : "  ") << argv[i] << std::endl;
        }}
        std::cerr << "" << std::endl;
    }

#if 0
    std::cerr << "  WordDriver" << std::endl << std::endl;
    std::cerr << "  USAGE: WordDriver [{-w | -p<root-paths> | -h}] [-u] [-d] [{-v | -s}] <search-spec>" << std::endl;
    std::cerr << "    where:" << std::endl << std::endl;
    std::cerr << "    -w             - searches from the current working directory. The default" << std::endl;
    std::cerr << "    -p<root-paths> - searches from the given root path(s), separated by \';\'," << std::endl;
    std::cerr << "                     eg." << std::endl;
    std::cerr << "                       -p\"c:\\windows;x:\\bin\"" << std::endl;
    std::cerr << "    -r             - deletes readonly files" << std::endl;
    std::cerr << "    -h             - searches from the roots of all drives on the system" << std::endl;
    std::cerr << "    -d             - displays the path(s) searched" << std::endl;
    std::cerr << "    -u             - do not act recursively" << std::endl;
    std::cerr << "    -v             - verbose output. Prints time, attributes, size and path. (default)" << std::endl;
    std::cerr << "    -s             - succinct output. Prints path only" << std::endl;
    std::cerr << "    <search-spec>  - one or more file search specifications, separated by \';\'," << std::endl;
    std::cerr << "                     eg." << std::endl;
    std::cerr << "                       \"*.exe\"" << std::endl;
    std::cerr << "                       \"myfile.ext\"" << std::endl;
    std::cerr << "                       \"*.exe;*.dll\"" << std::endl;
    std::cerr << "                       \"*.xl?;report.*\"" << std::endl;
    std::cerr << std::endl;
    std::cerr << "  Contact " << _nameSynesisSoftware << std::endl;
    std::cerr << "    at \"" << _wwwSynesisSoftware_SystemTools << "\"," << std::endl;
    std::cerr << "    or, via email, at \"" << _emailSynesisSoftware_SystemTools << "\"" << std::endl;
    std::cerr << std::endl;
#endif /* 0 */

    if(bExit)
    {
        exit(EXIT_FAILURE);
    }
}

/* ///////////////////////////// end of file //////////////////////////// */
