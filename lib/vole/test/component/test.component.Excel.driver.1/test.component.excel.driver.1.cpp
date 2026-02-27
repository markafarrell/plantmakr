/* /////////////////////////////////////////////////////////////////////////
 * File:        test.component.excel.driver.1.cpp
 *
 * Purpose:     Implementation file for the test.component.excel.driver.1 project.
 *
 * Created:     21st January 2010
 * Updated:     28th April 2010
 *
 * Status:      Wizard-generated
 *
 * License:     (Licensed under the Synesis Software Open License)
 *
 *              Copyright (c) 2010, Synesis Software Pty Ltd.
 *              All rights reserved.
 *
 *              www:        http://www.synesis.com.au/software
 *
 * ////////////////////////////////////////////////////////////////////// */


/* /////////////////////////////////////////////////////////////////////////
 * Test component header file include(s)
 */

#include <vole/vole.hpp>

/* /////////////////////////////////////////////////////////////////////////
 * Includes
 */

/* xTests Header Files */
#include <xtests/xtests.h>

/* STLSoft Header Files */
#include <stlsoft/stlsoft.h>
//#include <comstl/util/initialisers.hpp>

/* Standard C Header Files */
#include <stdlib.h>

/* /////////////////////////////////////////////////////////////////////////
 * Forward declarations
 */

namespace
{
    static void test_1_0(void);
    static void test_1_1(void);
    static void test_1_2(void);
    static void test_1_3(void);
    static void test_1_4(void);
    static void test_1_5(void);
    static void test_1_6(void);
    static void test_1_7(void);
    static void test_1_8(void);
    static void test_1_9(void);
    static void test_1_10(void);
    static void test_1_11(void);
    static void test_1_12(void);
    static void test_1_13(void);
    static void test_1_14(void);
    static void test_1_15(void);
    static void test_1_16(void);
    static void test_1_17(void);
    static void test_1_18(void);
    static void test_1_19(void);

} // anonymous namespace

/* /////////////////////////////////////////////////////////////////////////
 * Main
 */

int main(int argc, char **argv)
{
    int retCode = EXIT_SUCCESS;
    int verbosity = 2;

    ::OleInitialize(NULL);

    XTESTS_COMMANDLINE_PARSEVERBOSITY(argc, argv, &verbosity);

    if(XTESTS_START_RUNNER("test.component.excel.driver.1", verbosity))
    {
        XTESTS_RUN_CASE(test_1_0);
        XTESTS_RUN_CASE(test_1_1);
        XTESTS_RUN_CASE(test_1_2);
        XTESTS_RUN_CASE(test_1_3);
        XTESTS_RUN_CASE(test_1_4);
        XTESTS_RUN_CASE(test_1_5);
        XTESTS_RUN_CASE(test_1_6);
        XTESTS_RUN_CASE(test_1_7);
        XTESTS_RUN_CASE(test_1_8);
        XTESTS_RUN_CASE(test_1_9);
        XTESTS_RUN_CASE(test_1_10);
        XTESTS_RUN_CASE(test_1_11);
        XTESTS_RUN_CASE(test_1_12);
        XTESTS_RUN_CASE(test_1_13);
        XTESTS_RUN_CASE(test_1_14);
        XTESTS_RUN_CASE(test_1_15);
        XTESTS_RUN_CASE(test_1_16);
        XTESTS_RUN_CASE(test_1_17);
        XTESTS_RUN_CASE(test_1_18);
        XTESTS_RUN_CASE(test_1_19);

        XTESTS_PRINT_RESULTS();

        XTESTS_END_RUNNER_UPDATE_EXITCODE(&retCode);
    }

    return retCode;
}

/* /////////////////////////////////////////////////////////////////////////
 * Test function implementations
 */

namespace
{

    using vole::collection;
    using vole::object;
    using vole::coercion_level::valueCoercion;
    using vole::vole_exception;
    using vole::invocation_exception;

static void test_1_0()
{
    object instance;

    STLSOFT_SUPPRESS_UNUSED(instance);

    XTESTS_TEST_PASSED();
}

static void test_1_1()
{
    try
    {
        object Excel = object::create("{DDAA4054-D7F7-44f9-981B-3115D53299B2}", CLSCTX_ALL, valueCoercion);

        XTESTS_TEST_FAIL("should not get here");
    }
    catch(vole_exception& x)
    {
        HRESULT valid_options[] = { CO_E_CLASSSTRING, REGDB_E_CLASSNOTREG };

        XTESTS_TEST_INTEGER_EQUAL_ANY_IN_RANGE(valid_options, XTESTS_ARRAY_END_POST(valid_options), x.hr());
    }
}

static void test_1_2()
{
    try
    {
        object      Excel   =   object::create("Excel.Application", CLSCTX_LOCAL_SERVER, valueCoercion);

        XTESTS_TEST_POINTER_NOT_EQUAL(NULL, Excel.get());

        DISPID      id      =   Excel.get_memberId(L"{92086058-BCA4-47ed-A5EA-DFB856EF1BC0}");

        XTESTS_TEST_FAIL("should not get here");

        STLSOFT_SUPPRESS_UNUSED(id);
    }
    catch(invocation_exception& x)
    {
        XTESTS_TEST_INTEGER_EQUAL(DISP_E_UNKNOWNNAME, x.hr());
    }
}

static void test_1_3()
{
}

static void test_1_4()
{
}

static void test_1_5()
{
}

static void test_1_6()
{
}

static void test_1_7()
{
}

static void test_1_8()
{
}

static void test_1_9()
{
}

static void test_1_10()
{
}

static void test_1_11()
{
}

static void test_1_12()
{
    using vole::of_type;

    try
    {
        object      Excel = object::create("Excel.Application", CLSCTX_LOCAL_SERVER, valueCoercion);

        XTESTS_TEST_POINTER_NOT_EQUAL(NULL, Excel.get());

        try
        {

#ifdef _DEBUG
            Excel.put_property(L"Visible", true);
#endif /* _DEBUG */

            collection  Workbooks   =   Excel.get_property(of_type<collection>(), L"Workbooks");

            object      Workbook    =   Workbooks.invoke_method(of_type<object>(), L"Add");

            try
            {
//              collection  Worksheets  =   Excel.get_property(of_type<collection>(), L"Worksheets");

                object      Worksheet   =   Workbook.get_property(of_type<object>(), L"ActiveSheet");

                object      Cell1       =   Worksheet.get_property(of_type<object>(), L"Range", "A1");
                object      Cell2       =   Worksheet.get_property(of_type<object>(), L"Range", "B1");
                object      Cell3       =   Worksheet.get_property(of_type<object>(), L"Range", "C1");
                object      Cell4       =   Worksheet.get_property(of_type<object>(), L"Range", "D1");
                object      Cell5       =   Worksheet.get_property(of_type<object>(), L"Range", "E1");

                Cell1.put_property(L"Value", -12);
                Cell2.put_property(L"Value", -3400);
                Cell3.put_property(L"Value", -560000);
                Cell4.put_property(L"Value", -78000000);

                // Do cell #5, but this time using a cached dispatch identifier

                DISPID      idFormula   =   Cell5.get_memberId(L"Formula");

                Cell5.put_property(idFormula, "=SUM(A1:D1)");
                Cell5.invoke_method_v(L"Calculate");

                //fprintf(stdout, "Cell5=%s\n", stlsoft::c_str_ptr_a(stlsoft::c_str_ptr_a(Cell5.get_property(of_type<comstl::variant>(), L"Value"))));

                int         Value       =   Cell5.get_property(of_type<int>(), L"Value");

                XTESTS_TEST_INTEGER_EQUAL(-78563412, Value);

                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_property(L"Saved", 1);
            }
            catch(...)
            {
                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_property(L"Saved", 1);

                throw;
            }

            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");
        }
        catch(...)
        {
            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");

            throw;
        }
    }
    catch(vole_exception& x)
    {
        HRESULT valid_options[] = { CO_E_CLASSSTRING, REGDB_E_CLASSNOTREG };

        if(!XTESTS_TEST_INTEGER_EQUAL_ANY_IN_RANGE(valid_options, XTESTS_ARRAY_END_POST(valid_options), x.hr()))
        {
            fprintf(stderr, "unexpected failure: %s (0x%08x)\n", x.what(), x.hr());
        }
    }
}

static void test_1_13()
{
    using vole::of_type;

    try
    {
        object      Excel = object::create("Excel.Application", CLSCTX_LOCAL_SERVER, valueCoercion);

        XTESTS_TEST_POINTER_NOT_EQUAL(NULL, Excel.get());

        try
        {

#ifdef _DEBUG
            Excel.put_property(L"Visible", true);
#endif /* _DEBUG */

            collection  Workbooks   =   Excel.get_property(of_type<collection>(), L"Workbooks");

            object      Workbook    =   Workbooks.invoke_method(of_type<object>(), L"Add");

            try
            {
//              collection  Worksheets  =   Excel.get_property(of_type<collection>(), L"Worksheets");

                object      Worksheet   =   Workbook.get_property(of_type<object>(), L"ActiveSheet");

                object      Cell1       =   Worksheet.get_property(of_type<object>(), L"Range", "A1");
                object      Cell2       =   Worksheet.get_property(of_type<object>(), L"Range", "B1");
                object      Cell3       =   Worksheet.get_property(of_type<object>(), L"Range", "C1");
                object      Cell4       =   Worksheet.get_property(of_type<object>(), L"Range", "D1");
                object      Cell5       =   Worksheet.get_property(of_type<object>(), L"Range", "E1");

                Cell1.put_property(L"Value", -12);
                Cell2.put_property(L"Value", -3400);
                Cell3.put_property(L"Value", -560000);
                Cell4.put_property(L"Value", -78000000);

                DISPID      idFormula   =   Cell5.get_memberId(L"Formula");

                Cell5.put_property(idFormula, "=SUM(A1:D1)");
                Cell5.invoke_method_v(L"Calculate");

                //fprintf(stdout, "Cell5=%s\n", stlsoft::c_str_ptr_a(stlsoft::c_str_ptr_a(Cell5.get_property(of_type<comstl::variant>(), L"Value"))));

                std::string Value       =   Cell5.get_property(of_type<std::string>(), L"Value");

                XTESTS_TEST_MULTIBYTE_STRING_EQUAL("-78563412", Value);

                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_property(L"Saved", 1);
            }
            catch(...)
            {
                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_property(L"Saved", 1);

                throw;
            }

            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");
        }
        catch(...)
        {
            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");

            throw;
        }
    }
    catch(vole_exception& x)
    {
        HRESULT valid_options[] = { CO_E_CLASSSTRING, REGDB_E_CLASSNOTREG };

        if(!XTESTS_TEST_INTEGER_EQUAL_ANY_IN_RANGE(valid_options, XTESTS_ARRAY_END_POST(valid_options), x.hr()))
        {
            fprintf(stderr, "unexpected failure: %s (0x%08x)\n", x.what(), x.hr());
        }
    }
}

static void test_1_14()
{
#ifdef VOLE_METHOD_TEMPLATE_RETURN_SUPPORT
    try
    {
        object      Excel = object::create("Excel.Application", CLSCTX_LOCAL_SERVER, valueCoercion);

        XTESTS_TEST_POINTER_NOT_EQUAL(NULL, Excel.get());

        try
        {

# ifdef _DEBUG
            Excel.put_prop(L"Visible", true);
# endif /* _DEBUG */

            collection  Workbooks   =   Excel.get_property<collection>(L"Workbooks");

            object      Workbook    =   Workbooks.invoke_method<object>(L"Add");

            try
            {
//              collection  Worksheets  =   Excel.get_property<cection>(), L"Worksheets");

                object      Worksheet   =   Workbook.get_property<object>(L"ActiveSheet");

                object      Cell1       =   Worksheet.get_property<object>(L"Range", "A1");
                object      Cell2       =   Worksheet.get_property<object>(L"Range", "B1");
                object      Cell3       =   Worksheet.get_property<object>(L"Range", "C1");
                object      Cell4       =   Worksheet.get_property<object>(L"Range", "D1");
                object      Cell5       =   Worksheet.get_property<object>(L"Range", "E1");

                Cell1.put_prop(L"Value", -12);
                Cell2.put_prop(L"Value", -3400);
                Cell3.put_prop(L"Value", -560000);
                Cell4.put_prop(L"Value", -78000000);

                // Do cell #5, but this time using a cached dispatch identifier

                DISPID      idFormula   =   Cell5.get_memberId(L"Formula");

                Cell5.put_prop(idFormula, "=SUM(A1:D1)");
                Cell5.invoke_method_v(L"Calculate");

                //fprintf(stdout, "Cell5=%s\n", stlsoft::c_str_ptr_a(stlsoft::c_str_ptr_a(Cell5.get_property<ctl::variant>(), L"Value"))));

                int         Value       =   Cell5.get_property<int>(L"Value");

                XTESTS_TEST_INTEGER_EQUAL(-78563412, Value);

                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_prop(L"Saved", 1);
            }
            catch(...)
            {
                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_property(L"Saved", 1);

                throw;
            }

            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");
        }
        catch(...)
        {
            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");

            throw;
        }
    }
    catch(vole_exception& x)
    {
        HRESULT valid_options[] = { CO_E_CLASSSTRING, REGDB_E_CLASSNOTREG };

        if(!XTESTS_TEST_INTEGER_EQUAL_ANY_IN_RANGE(valid_options, XTESTS_ARRAY_END_POST(valid_options), x.hr()))
        {
            fprintf(stderr, "unexpected failure: %s (0x%08x)\n", x.what(), x.hr());
        }
    }
#endif /* VOLE_METHOD_TEMPLATE_RETURN_SUPPORT */
}

static void test_1_15()
{
#ifdef VOLE_METHOD_TEMPLATE_RETURN_SUPPORT
    try
    {
        object      Excel = object::create("Excel.Application", CLSCTX_LOCAL_SERVER, valueCoercion);

        XTESTS_TEST_POINTER_NOT_EQUAL(NULL, Excel.get());

        try
        {

# ifdef _DEBUG
            Excel.put_property(L"Visible", true);
# endif /* _DEBUG */

            collection  Workbooks   =   Excel.get_property<collection>(L"Workbooks");

            object      Workbook    =   Workbooks.invoke_method<object>(L"Add");

            try
            {
//              collection  Worksheets  =   Excel.get_property<cection>(), L"Worksheets");

                object      Worksheet   =   Workbook.get_property<object>(L"ActiveSheet");

                object      Cell1       =   Worksheet.get_property<object>(L"Range", "A1");
                object      Cell2       =   Worksheet.get_property<object>(L"Range", "B1");
                object      Cell3       =   Worksheet.get_property<object>(L"Range", "C1");
                object      Cell4       =   Worksheet.get_property<object>(L"Range", "D1");
                object      Cell5       =   Worksheet.get_property<object>(L"Range", "E1");

                Cell1.put_property(L"Value", -12);
                Cell2.put_property(L"Value", -3400);
                Cell3.put_property(L"Value", -560000);
                Cell4.put_property(L"Value", -78000000);

                DISPID      idFormula   =   Cell5.get_memberId(L"Formula");

                Cell5.put_property(idFormula, "=SUM(A1:D1)");
                Cell5.invoke_method_v(L"Calculate");

                //fprintf(stdout, "Cell5=%s\n", stlsoft::c_str_ptr_a(stlsoft::c_str_ptr_a(Cell5.get_property<ctl::variant>(), L"Value"))));

                std::string Value       =   Cell5.get_property<std::string>(L"Value");

                XTESTS_TEST_MULTIBYTE_STRING_EQUAL("-78563412", Value);

                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_property(L"Saved", 1);
            }
            catch(...)
            {
                // Set .Saved property of workbook to TRUE so we aren't prompted
                // to save when we tell Excel to quit...
                Workbook.put_property(L"Saved", 1);

                throw;
            }

            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");
        }
        catch(...)
        {
            // Tell Excel to quit (i.e. App.Quit)
            Excel.invoke_method_v(L"Quit");

            throw;
        }
    }
    catch(vole_exception& x)
    {
        HRESULT valid_options[] = { CO_E_CLASSSTRING, REGDB_E_CLASSNOTREG };

        if(!XTESTS_TEST_INTEGER_EQUAL_ANY_IN_RANGE(valid_options, XTESTS_ARRAY_END_POST(valid_options), x.hr()))
        {
            fprintf(stderr, "unexpected failure: %s (0x%08x)\n", x.what(), x.hr());
        }
    }
#endif /* VOLE_METHOD_TEMPLATE_RETURN_SUPPORT */
}

static void test_1_16()
{
}

static void test_1_17()
{
}

static void test_1_18()
{
}

static void test_1_19()
{
}


} // anonymous namespace

/* ///////////////////////////// end of file //////////////////////////// */
