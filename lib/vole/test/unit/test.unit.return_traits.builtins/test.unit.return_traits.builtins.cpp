/* /////////////////////////////////////////////////////////////////////////
 * File:        test.unit.return_traits.builtins.cpp
 *
 * Purpose:     Implementation file for the test.unit.return_traits.builtins project.
 *
 * Created:     23rd August 2008
 * Updated:     18th March 2010
 *
 * Status:      Wizard-generated
 *
 * License:     (Licensed under the Synesis Software Open License)
 *
 *              Copyright (c) 2008-2010, Synesis Software Pty Ltd.
 *              All rights reserved.
 *
 *              www:        http://www.synesis.com.au/software
 *
 * ////////////////////////////////////////////////////////////////////// */


/* /////////////////////////////////////////////////////////////////////////
 * Test component header file include(s)
 */

#include <vole/return_traits/builtins.hpp>

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
#include <float.h>
#include <stdlib.h>

/* /////////////////////////////////////////////////////////////////////////
 * Forward declarations
 */

namespace
{

    static void test_bool_from_bool_noCoercion(void);
    static void test_bool_from_bool_naturalPromotion(void);
    static void test_bool_from_bool_nonDestructiveCoercion(void);
    static void test_bool_from_bool_valueCoercion(void);

    static void test_bool_from_int_noCoercion(void);
    static void test_bool_from_int_naturalPromotion(void);
    static void test_bool_from_int_nonDestructiveCoercion(void);
    static void test_bool_from_int_valueCoercion(void);

    static void test_short_from_long_noCoercion(void);
    static void test_short_from_long_naturalPromotion(void);
    static void test_short_from_long_nonDestructiveCoercion_in_range(void);
    static void test_short_from_long_nonDestructiveCoercion_out_of_range(void);
    static void test_short_from_long_valueCoercion_in_range(void);
    static void test_short_from_long_valueCoercion_out_of_range(void);

    static void test_ulong_from_short_noCoercion(void);
    static void test_ulong_from_short_naturalPromotion(void);
    static void test_ulong_from_short_nonDestructiveCoercion_in_range(void);
    static void test_ulong_from_short_nonDestructiveCoercion_out_of_range(void);
    static void test_ulong_from_short_valueCoercion_in_range(void);
    static void test_ulong_from_short_valueCoercion_out_of_range(void);

    static void test_1_5(void);
    static void test_1_6(void);
    static void test_1_7(void);
    static void test_1_8(void);
    static void test_1_9(void);

    static void test_bool_from_HRESULT_noCoercion(void);
    static void test_bool_from_HRESULT_naturalPromotion(void);
    static void test_bool_from_HRESULT_nonDestructiveCoercion(void);
    static void test_bool_from_HRESULT_valueCoercion(void);

    static void test_float_from_DECIMAL_noCoercion(void);
    static void test_float_from_DECIMAL_naturalPromotion(void);
    static void test_float_from_DECIMAL_nonDestructiveCoercion(void);
    static void test_float_from_DECIMAL_valueCoercion(void);

    static void test_bool_from_int_natural1(void);
    static void test_bool_from_int_natural2(void);
    static void test_bool_from_int_natural3(void);
    static void test_bool_from_int_natural4(void);
    static void test_bool_from_int_natural5(void);
    static void test_bool_from_int_natural6(void);
    static void test_bool_from_int_natural7(void);
    static void test_bool_from_int_natural8(void);
    static void test_bool_from_int_natural9(void);

} // anonymous namespace

/* /////////////////////////////////////////////////////////////////////////
 * Main
 */

int main(int argc, char **argv)
{
    int retCode = EXIT_SUCCESS;
    int verbosity = 2;

    XTESTS_COMMANDLINE_PARSEVERBOSITY(argc, argv, &verbosity);

    if(XTESTS_START_RUNNER("test.unit.return_traits.builtins", verbosity))
    {
        XTESTS_RUN_CASE(test_bool_from_bool_noCoercion);
        XTESTS_RUN_CASE(test_bool_from_bool_naturalPromotion);
        XTESTS_RUN_CASE(test_bool_from_bool_nonDestructiveCoercion);
        XTESTS_RUN_CASE(test_bool_from_bool_valueCoercion);

        XTESTS_RUN_CASE_THAT_THROWS(test_bool_from_int_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_bool_from_int_naturalPromotion);
        XTESTS_RUN_CASE(test_bool_from_int_nonDestructiveCoercion);
        XTESTS_RUN_CASE(test_bool_from_int_valueCoercion);

        XTESTS_RUN_CASE_THAT_THROWS(test_short_from_long_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_short_from_long_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_short_from_long_nonDestructiveCoercion_in_range);
        XTESTS_RUN_CASE_THAT_THROWS(test_short_from_long_nonDestructiveCoercion_out_of_range, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_short_from_long_valueCoercion_in_range);
        XTESTS_RUN_CASE_THAT_THROWS(test_short_from_long_valueCoercion_out_of_range, vole::type_conversion_exception);

        XTESTS_RUN_CASE_THAT_THROWS(test_ulong_from_short_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_ulong_from_short_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_ulong_from_short_nonDestructiveCoercion_in_range);
        XTESTS_RUN_CASE_THAT_THROWS(test_ulong_from_short_nonDestructiveCoercion_out_of_range, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_ulong_from_short_valueCoercion_in_range);
        XTESTS_RUN_CASE_THAT_THROWS(test_ulong_from_short_valueCoercion_out_of_range, vole::type_conversion_exception);


        XTESTS_RUN_CASE(test_1_5);
        XTESTS_RUN_CASE(test_1_6);
        XTESTS_RUN_CASE(test_1_7);
        XTESTS_RUN_CASE(test_1_8);
        XTESTS_RUN_CASE(test_1_9);

        XTESTS_RUN_CASE_THAT_THROWS(test_bool_from_HRESULT_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_bool_from_HRESULT_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_bool_from_HRESULT_nonDestructiveCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_bool_from_HRESULT_valueCoercion);

        XTESTS_RUN_CASE_THAT_THROWS(test_float_from_DECIMAL_noCoercion, vole::type_conversion_exception);
        XTESTS_RUN_CASE_THAT_THROWS(test_float_from_DECIMAL_naturalPromotion, vole::type_conversion_exception);
        XTESTS_RUN_CASE(test_float_from_DECIMAL_nonDestructiveCoercion);
        XTESTS_RUN_CASE(test_float_from_DECIMAL_valueCoercion);

        XTESTS_RUN_CASE(test_bool_from_int_natural1);
        XTESTS_RUN_CASE(test_bool_from_int_natural2);
        XTESTS_RUN_CASE(test_bool_from_int_natural3);
        XTESTS_RUN_CASE(test_bool_from_int_natural4);
        XTESTS_RUN_CASE(test_bool_from_int_natural5);
        XTESTS_RUN_CASE(test_bool_from_int_natural6);
        XTESTS_RUN_CASE(test_bool_from_int_natural7);
        XTESTS_RUN_CASE(test_bool_from_int_natural8);
        XTESTS_RUN_CASE(test_bool_from_int_natural9);

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
    typedef vole::com_return_traits<bool>       bool_traits_t;
    typedef vole::com_return_traits<short>      short_traits_t;
    typedef vole::com_return_traits<int>        int_traits_t;
    typedef vole::com_return_traits<long>       long_traits_t;
    typedef vole::com_return_traits<float>      float_traits_t;
    typedef vole::com_return_traits<double>     double_traits_t;

    using vole::coercion_level::noCoercion;
    using vole::coercion_level::naturalPromotion;
    using vole::coercion_level::nonDestructiveCoercion;
    using vole::coercion_level::valueCoercion;

    static void get_DECIMALs(DECIMAL* decZero, DECIMAL* decMin, DECIMAL* decMax)
    {
        ::memset(decZero, 0, sizeof(*decZero));
        ::memset(decMin, 0, sizeof(*decMin));
        ::memset(decMax, 0, sizeof(*decMax));

        decMin->sign    =   DECIMAL_NEG;
        decMin->Hi32    =   stlsoft::limit_traits<ULONG>::maximum();
        decMin->Mid32   =   stlsoft::limit_traits<ULONG>::maximum();
        decMin->Lo32    =   stlsoft::limit_traits<ULONG>::maximum();

        decMax->sign    =   0;
        decMax->Hi32    =   stlsoft::limit_traits<ULONG>::maximum();
        decMax->Mid32   =   stlsoft::limit_traits<ULONG>::maximum();
        decMax->Lo32    =   stlsoft::limit_traits<ULONG>::maximum();
    }


static void test_bool_from_bool_noCoercion()
{
    comstl::variant     b1(true);
    comstl::variant     b2(false);

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(b1, noCoercion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(b2, noCoercion));
}

static void test_bool_from_bool_naturalPromotion()
{
    comstl::variant     b1(true);
    comstl::variant     b2(false);

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(b1, naturalPromotion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(b2, naturalPromotion));
}

static void test_bool_from_bool_nonDestructiveCoercion()
{
    comstl::variant     b1(true);
    comstl::variant     b2(false);

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(b1, nonDestructiveCoercion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(b2, nonDestructiveCoercion));
}

static void test_bool_from_bool_valueCoercion()
{
    comstl::variant     b1(true);
    comstl::variant     b2(false);

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(b1, valueCoercion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(b2, valueCoercion));
}


static void test_bool_from_int_noCoercion()
{
    comstl::variant     i1(0);

    bool_traits_t::convert(i1, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted int=>bool with no coercion");
}

static void test_bool_from_int_naturalPromotion()
{
    comstl::variant     i1(-10);
    comstl::variant     i2(0);

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(i1, naturalPromotion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(i2, naturalPromotion));
}


static void test_bool_from_int_nonDestructiveCoercion()
{
    comstl::variant     i1(-10);
    comstl::variant     i2(0);

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(i1, nonDestructiveCoercion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(i2, nonDestructiveCoercion));
}

static void test_bool_from_int_valueCoercion()
{
    comstl::variant     i1(-10);
    comstl::variant     i2(0);

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(i1, valueCoercion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(i2, valueCoercion));
}


static void test_short_from_long_noCoercion()
{
    comstl::variant     l1(0L);

    short_traits_t::convert(l1, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with no coercion");
}

static void test_short_from_long_naturalPromotion()
{
    comstl::variant     l1(0L);

    short_traits_t::convert(l1, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with natural promotion");
}

static void test_short_from_long_nonDestructiveCoercion_in_range()
{
    comstl::variant     l1(0L);
    comstl::variant     l2(32767L);
    comstl::variant     l3(-32768L);

    XTESTS_TEST_INTEGER_EQUAL(short(0), short_traits_t::convert(l1, nonDestructiveCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(32767), short_traits_t::convert(l2, nonDestructiveCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(-32768), short_traits_t::convert(l3, nonDestructiveCoercion));
}

static void test_short_from_long_nonDestructiveCoercion_out_of_range()
{
    comstl::variant     l1(32768L);
    comstl::variant     l2(-32769L);

    short_traits_t::convert(l1, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with non-destructive truncation");

    short_traits_t::convert(l2, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with non-destructive truncation");
}

static void test_short_from_long_valueCoercion_in_range()
{
    comstl::variant     l1(0L);
    comstl::variant     l2(32767L);
    comstl::variant     l3(-32768L);

    XTESTS_TEST_INTEGER_EQUAL(short(0), short_traits_t::convert(l1, valueCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(32767), short_traits_t::convert(l2, valueCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(-32768), short_traits_t::convert(l3, valueCoercion));
}

static void test_short_from_long_valueCoercion_out_of_range()
{
    comstl::variant     l1(32768L);
    comstl::variant     l2(-32769L);

    short_traits_t::convert(l1, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with value coercion");

    short_traits_t::convert(l2, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with value coercion");
}



static void test_ulong_from_short_noCoercion()
{
    comstl::variant     v(static_cast<unsigned long>(0));

    short_traits_t::convert(v, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with no coercion");
}

static void test_ulong_from_short_naturalPromotion()
{
    comstl::variant     v(static_cast<unsigned long>(0));

    short_traits_t::convert(v, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with natural promotion");
}

static void test_ulong_from_short_nonDestructiveCoercion_in_range()
{
    comstl::variant     s1(0L);
    comstl::variant     l2(32767L);
    comstl::variant     l3(-32768L);

    XTESTS_TEST_INTEGER_EQUAL(short(0), short_traits_t::convert(s1, nonDestructiveCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(32767), short_traits_t::convert(l2, nonDestructiveCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(-32768), short_traits_t::convert(l3, nonDestructiveCoercion));
}

static void test_ulong_from_short_nonDestructiveCoercion_out_of_range()
{
    comstl::variant     s1(32768L);
    comstl::variant     l2(-32769L);

    short_traits_t::convert(s1, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with non-destructive truncation");

    short_traits_t::convert(l2, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with non-destructive truncation");
}

static void test_ulong_from_short_valueCoercion_in_range()
{
    comstl::variant     s1(0L);
    comstl::variant     l2(32767L);
    comstl::variant     l3(-32768L);

    XTESTS_TEST_INTEGER_EQUAL(short(0), short_traits_t::convert(s1, valueCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(32767), short_traits_t::convert(l2, valueCoercion));
    XTESTS_TEST_INTEGER_EQUAL(short(-32768), short_traits_t::convert(l3, valueCoercion));
}

static void test_ulong_from_short_valueCoercion_out_of_range()
{
    comstl::variant     s1(32768L);
    comstl::variant     l2(-32769L);

    short_traits_t::convert(s1, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with value coercion");

    short_traits_t::convert(l2, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted long=>short with value coercion");
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


static void test_bool_from_HRESULT_noCoercion()
{
    VARIANT v1;
    VARIANT v2;

    ::VariantInit(&v1);
    ::VariantInit(&v2);

    v1.vt       =   VT_ERROR;
    v1.scode    =   S_OK;

    v2.vt       =   VT_ERROR;
    v2.scode    =   E_UNEXPECTED;

    bool_traits_t::convert(v1, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>bool with no coercion");

    bool_traits_t::convert(v2, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>bool with no coercion");
}

static void test_bool_from_HRESULT_naturalPromotion()
{
    VARIANT v1;
    VARIANT v2;

    ::VariantInit(&v1);
    ::VariantInit(&v2);

    v1.vt       =   VT_ERROR;
    v1.scode    =   S_OK;

    v2.vt       =   VT_ERROR;
    v2.scode    =   E_UNEXPECTED;

    bool_traits_t::convert(v1, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>bool with natural promotion");

    bool_traits_t::convert(v2, naturalPromotion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>bool with natural promotion");
}

static void test_bool_from_HRESULT_nonDestructiveCoercion()
{
    VARIANT v1;
    VARIANT v2;

    ::VariantInit(&v1);
    ::VariantInit(&v2);

    v1.vt       =   VT_ERROR;
    v1.scode    =   S_OK;

    v2.vt       =   VT_ERROR;
    v2.scode    =   E_UNEXPECTED;

    bool_traits_t::convert(v1, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>bool with non-destructive coercion");

    bool_traits_t::convert(v2, nonDestructiveCoercion);

    XTESTS_TEST_FAIL("incorrectly converted HRESULT=>bool with non-destructive coercion");
}

static void test_bool_from_HRESULT_valueCoercion()
{
    VARIANT v1;
    VARIANT v2;

    ::VariantInit(&v1);
    ::VariantInit(&v2);

    v1.vt       =   VT_ERROR;
    v1.scode    =   S_OK;

    v2.vt       =   VT_ERROR;
    v2.scode    =   E_UNEXPECTED;

    XTESTS_TEST_BOOLEAN_EQUAL(true, bool_traits_t::convert(v1, valueCoercion));
    XTESTS_TEST_BOOLEAN_EQUAL(false, bool_traits_t::convert(v2, valueCoercion));
}



static void test_float_from_DECIMAL_noCoercion()
{
    DECIMAL     decZero;
    DECIMAL     decMin;
    DECIMAL     decMax;

    get_DECIMALs(&decZero, &decMin, &decMax);

    comstl::variant vdecZero(decZero);
    comstl::variant vdecMin(decMin);
    comstl::variant vdecMax(decMax);

    float_traits_t::convert(vdecZero, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted DECIMAL=>float with no coercion");

    float_traits_t::convert(vdecMin, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted DECIMAL=>float with no coercion");

    float_traits_t::convert(vdecMax, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted DECIMAL=>float with no coercion");
}

static void test_float_from_DECIMAL_naturalPromotion()
{
    DECIMAL     decZero;
    DECIMAL     decMin;
    DECIMAL     decMax;

    get_DECIMALs(&decZero, &decMin, &decMax);

    comstl::variant vdecZero(decZero);
    comstl::variant vdecMin(decMin);
    comstl::variant vdecMax(decMax);

    float_traits_t::convert(vdecZero, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted DECIMAL=>float with natural promotion");

    float_traits_t::convert(vdecMin, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted DECIMAL=>float with natural promotion");

    float_traits_t::convert(vdecMax, noCoercion);

    XTESTS_TEST_FAIL("incorrectly converted DECIMAL=>float with natural promotion");
}

static void test_float_from_DECIMAL_nonDestructiveCoercion()
{
    DECIMAL     decZero;
    DECIMAL     decMin;
    DECIMAL     decMax;

    get_DECIMALs(&decZero, &decMin, &decMax);

    comstl::variant vdecZero(decZero);
    comstl::variant vdecMin(decMin);
    comstl::variant vdecMax(decMax);

    XTESTS_TEST_FLOATINGPOINT_EQUAL_APPROX(0.0, float_traits_t::convert(vdecZero, nonDestructiveCoercion));
    XTESTS_TEST_FLOATINGPOINT_EQUAL_APPROX(-7.9228162458924105385300197375e+28, float_traits_t::convert(vdecMin, nonDestructiveCoercion));
    XTESTS_TEST_FLOATINGPOINT_EQUAL_APPROX(7.9228162458924105385300197375e+28, float_traits_t::convert(vdecMax, nonDestructiveCoercion));
}

static void test_float_from_DECIMAL_valueCoercion()
{
    DECIMAL     decZero;
    DECIMAL     decMin;
    DECIMAL     decMax;

    get_DECIMALs(&decZero, &decMin, &decMax);

    comstl::variant vdecZero(decZero);
    comstl::variant vdecMin(decMin);
    comstl::variant vdecMax(decMax);

    XTESTS_TEST_FLOATINGPOINT_EQUAL_APPROX(0.0, float_traits_t::convert(vdecZero, valueCoercion));
    XTESTS_TEST_FLOATINGPOINT_EQUAL_APPROX(-7.9228162458924105385300197375e+28, float_traits_t::convert(vdecMin, valueCoercion));
    XTESTS_TEST_FLOATINGPOINT_EQUAL_APPROX(7.9228162458924105385300197375e+28, float_traits_t::convert(vdecMax, valueCoercion));
}

static void test_bool_from_int_natural1()
{
}

static void test_bool_from_int_natural2()
{
}

static void test_bool_from_int_natural3()
{
}

static void test_bool_from_int_natural4()
{
}

static void test_bool_from_int_natural5()
{
}

static void test_bool_from_int_natural6()
{
}

static void test_bool_from_int_natural7()
{
}

static void test_bool_from_int_natural8()
{
}

static void test_bool_from_int_natural9()
{
}


} // anonymous namespace

/* ///////////////////////////// end of file //////////////////////////// */
