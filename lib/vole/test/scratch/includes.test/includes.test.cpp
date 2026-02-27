/* /////////////////////////////////////////////////////////////////////////
 * File:        includes.test.cpp
 *
 * Purpose:     Implementation file for the includes_test project.
 *
 * Created:     3rd January 2007
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
/* /////////////////////////////////////////////////////////////////////////


/* VOLE Header Files */
//#include <vole/object.hpp>
#include <vole/return_traits/std_types.hpp>

/* STLSoft Header Files */
#include <stlsoft/stlsoft.h>

/* Standard C++ Header Files */
#include <exception>

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
    /* . */
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
        res = main_(argc, argv);
    }
    catch(std::exception &x)
    {
        fprintf(stderr, "Unhandled error: %s\n", x.what());

        res = EXIT_FAILURE;
    }
    catch(...)
    {
        fprintf(stderr, "Unhandled unknown error\n");

        res = EXIT_FAILURE;
    }

#if defined(_MSC_VER) && \
    defined(_DEBUG)
    _CrtMemDumpAllObjectsSince(&memState);
#endif /* _MSC_VER) && _DEBUG */

    return res;
}

/* ///////////////////////////// end of file //////////////////////////// */
