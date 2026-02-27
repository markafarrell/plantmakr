/* /////////////////////////////////////////////////////////////////////////
 * File:        test.unit.return_traits.std_types.cpp
 *
 * Purpose:     Implementation file for the test.unit.return_traits.std_types project.
 *
 * Created:     23rd August 2008
 * Updated:     23rd August 2008
 *
 * Status:      Wizard-generated
 *
 * License:     (Licensed under the Synesis Software Open License)
 *
 *              Copyright (c) 2008, Synesis Software Pty Ltd.
 *              All rights reserved.
 *
 *              www:        http://www.synesis.com.au/software
 *
 * ////////////////////////////////////////////////////////////////////// */


/* /////////////////////////////////////////////////////////////////////////
 * Test component header file include(s)
 */

#include <vole/return_traits/std_types.hpp>

/* /////////////////////////////////////////////////////////////////////////
 * Includes
 */

/* xTests Header Files */
#include <xtests/xtests.h>

/* STLSoft Header Files */
#include <stlsoft/stlsoft.h>
#include <comstl/util/variant.hpp>
#include <winstl/error/error_desc.hpp>

/* Standard C Header Files */
#include <stdlib.h>

/* /////////////////////////////////////////////////////////////////////////
 * Forward declarations
 */

namespace
{

    static void test_string_from_int_noCoercion(void);
    static void test_string_from_int_naturalPromotion(void);
    static void test_string_from_int_nonDestructiveCoercion(void);
    static void test_string_from_int_valueCoercion(void);

    static void test_wstring_from_int_noCoercion(void);
    static void test_wstring_from_int_naturalPromotion(void);
    static void test_wstring_from_int_nonDestructiveCoercion(void);
    static void test_wstring_from_int_valueCoercion(void);

    static void test_string_from_HRESULT_noCoercion(void);
    static void test_string_from_HRESULT_naturalPromotion(void);
    static void test_string_from_HRESULT_nonDestructiveCoercion(void);
    static void test_string_from_HRESULT_valueCoercion(void);

    static void test_wstring_from_HRESULT_noCoercion(void);
    static void test_wstring_from_HRESULT_naturalPromotion(void);
    static void test_wstring_from_HRESULT_nonDestructiveCoercion(void);
    static void test_wstring_from_HRESULT_valueCoercion(void);

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

    XTESTS_COMMANDLINE_PARSEVERBOSITY(argc, argv, &verbosity);

    if(XTESTS_START_RUNNER("test.unit.return_traits.std_types", verbosity))
    {
        XTESTS_RUN_CASE_THAT_THROWS(test_string_from_int_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_string_from_int_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_string_from_int_nonDestructiveCoercion);
        XTESTS_RUN_CASE(test_string_from_int_valueCoercion);

        XTESTS_RUN_CASE_THAT_THROWS(test_wstring_from_int_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_wstring_from_int_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_wstring_from_int_nonDestructiveCoercion);
        XTESTS_RUN_CASE(test_wstring_from_int_valueCoercion);

        XTESTS_RUN_CASE_THAT_THROWS(test_string_from_HRESULT_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_string_from_HRESULT_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_string_from_HRESULT_nonDestructiveCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_string_from_HRESULT_valueCoercion);

        XTESTS_RUN_CASE_THAT_THROWS(test_wstring_from_HRESULT_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_wstring_from_HRESULT_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_wstring_from_HRESULT_nonDestructiveCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_wstring_from_HRESULT_valueCoercion);


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
    typedef vole::com_return_traits<std::string>        string_traits_t;
    typedef vole::com_return_traits<std::wstring>       wstring_traits_t;

    using vole::coercion_level::noCoercion;
    using vole::coercion_level::naturalPromotion;
    using vole::coercion_level::nonDestructiveCoercion;
    using vole::coercion_level::valueCoercion;


static void test_string_from_int_noCoercion()
{
    comstl::variant     i1(10);

    string_traits_t::convert(i1, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted int=>string with no coercion");
}

static void test_string_from_int_naturalPromotion()
{
    comstl::variant     i1(10);

    string_traits_t::convert(i1, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted int=>string with natural promotion");
}

static void test_string_from_int_nonDestructiveCoercion()
{
    comstl::variant     i1(10);

    XTESTS_TEST_MULTIBYTE_STRING_EQUAL("10", string_traits_t::convert(i1, valueCoercion));
}

static void test_string_from_int_valueCoercion()
{
    comstl::variant     i1(10);

    XTESTS_TEST_MULTIBYTE_STRING_EQUAL("10", string_traits_t::convert(i1, valueCoercion));
}


static void test_wstring_from_int_noCoercion()
{
    comstl::variant     i1(10);

    wstring_traits_t::convert(i1, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted int=>wstring with no coercion");
}

static void test_wstring_from_int_naturalPromotion()
{
    comstl::variant     i1(10);

    wstring_traits_t::convert(i1, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted int=>wstring with natural promotion");
}

static void test_wstring_from_int_nonDestructiveCoercion()
{
    comstl::variant     i1(10);

    XTESTS_TEST_WIDE_STRING_EQUAL(L"10", wstring_traits_t::convert(i1, valueCoercion));
}

static void test_wstring_from_int_valueCoercion()
{
    comstl::variant     i1(10);

    XTESTS_TEST_WIDE_STRING_EQUAL(L"10", wstring_traits_t::convert(i1, valueCoercion));
}


static void test_string_from_HRESULT_noCoercion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    string_traits_t::convert(e1, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::string with no coercion");

    string_traits_t::convert(e2, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::string with no coercion");
}

static void test_string_from_HRESULT_naturalPromotion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    string_traits_t::convert(e1, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::string with natural promotion");

    string_traits_t::convert(e2, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::string with natural promotion");
}

static void test_string_from_HRESULT_nonDestructiveCoercion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    string_traits_t::convert(e1, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::string with non-destructive coercion");

    string_traits_t::convert(e2, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::string with non-destructive coercion");
}

static void test_string_from_HRESULT_valueCoercion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    XTESTS_TEST_MULTIBYTE_STRING_EQUAL(winstl::error_desc_a(e1.scode), string_traits_t::convert(e1, valueCoercion));
    XTESTS_TEST_MULTIBYTE_STRING_EQUAL(winstl::error_desc_a(e2.scode), string_traits_t::convert(e2, valueCoercion));
}



static void test_wstring_from_HRESULT_noCoercion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    wstring_traits_t::convert(e1, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::wstring with no coercion");

    wstring_traits_t::convert(e2, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::wstring with no coercion");
}

static void test_wstring_from_HRESULT_naturalPromotion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    wstring_traits_t::convert(e1, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::wstring with natural promotion");

    wstring_traits_t::convert(e2, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::wstring with natural promotion");
}

static void test_wstring_from_HRESULT_nonDestructiveCoercion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    wstring_traits_t::convert(e1, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::wstring with non-destructive coercion");

    wstring_traits_t::convert(e2, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>std::wstring with non-destructive coercion");
}

static void test_wstring_from_HRESULT_valueCoercion()
{
    VARIANT e1;
    VARIANT e2;

    ::VariantInit(&e1);
    ::VariantInit(&e2);

    e1.vt       =   VT_ERROR;
    e1.scode    =   S_OK;

    e2.vt       =   VT_ERROR;
    e2.scode    =   E_UNEXPECTED;

    XTESTS_TEST_WIDE_STRING_EQUAL(winstl::error_desc_w(e1.scode), wstring_traits_t::convert(e1, valueCoercion));
    XTESTS_TEST_WIDE_STRING_EQUAL(winstl::error_desc_w(e2.scode), wstring_traits_t::convert(e2, valueCoercion));
}



static void test_1_2()
{
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
}

static void test_1_13()
{
}

static void test_1_14()
{
}

static void test_1_15()
{
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
