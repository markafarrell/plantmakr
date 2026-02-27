// openrj.com.test.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#include <vole/vole.hpp>

#include <comstl/error/errorinfo_desc.hpp>
#include <comstl/util/initialisers.hpp>

int _tmain(int argc, _TCHAR* argv[])
{
    using std::cout;
    using std::cerr;
    using std::endl;

    try
    {
        comstl::com_init    init;

        if(argc < 2)
        {
            cout << "USAGE: openrj.com.text.exe <db-file>" << endl;
        }
        else
        {
            using vole::object;
            using vole::collection;
            using vole::of_type;

            char const  *dbFile     =   argv[1];
            collection  db          =   object::create("openrj.COM.database");

#ifdef VOLE_METHOD_TEMPLATE_RETURN_SUPPORT

            db.invoke_method<void>(L"OpenFile", dbFile);

            long        numRecords  =   db.Count;
            std::string fileName    =   db.get_property<std::string>(L"File");

            cout << "Database in file '" << fileName << "' contains " << numRecords << " record(s)" << endl;

            { for(collection::iterator b = db.begin(); b != db.end(); ++b)
            {
                collection  record(*b);
                long        numFields   =   record.Count;

                cout << "  Record has " << numFields << " field(s)" << endl;

                { for(collection::iterator b2 = record.begin(); b2 != record.end(); ++b2)
                {
                    object          field(*b2);
                    std::string     name    =   field.get_property<std::string>(L"Name");
                    std::string     value   =   field.get_property<std::string>(L"Value");

                    cout << "    " << name << "=" << value << endl;
                }}
            }}

#else /* ? VOLE_METHOD_TEMPLATE_RETURN_SUPPORT */

            db.invoke_method_v(L"OpenFile", dbFile);

            long        numRecords  =   db.get_Count();
            std::string fileName    =   db.get_property(of_type<std::string>(), L"File");

            cout << "Database in file '" << fileName << "' contains " << numRecords << " record(s)" << endl;

            { for(collection::iterator b = db.begin(); b != db.end(); ++b)
            {
                collection  record(*b);
                long        numFields   =   record.get_Count();

                cout << "  Record has " << numFields << " field(s)" << endl;

                { for(collection::iterator b2 = record.begin(); b2 != record.end(); ++b2)
                {
                    object          field(*b2);
                    std::string     name    =   field.get_property(of_type<std::string>(), L"Name");
                    std::string     value   =   field.get_property(of_type<std::string>(), L"Value");

                    cout << "    " << name << "=" << value << endl;
                }}
            }}

#endif /* VOLE_METHOD_TEMPLATE_RETURN_SUPPORT */
        }

    }
    catch(::vole::vole_exception &x)
    {
        cerr << x.what() << ": " << comstl::errorinfo_desc() << endl;
    }
    catch(std::exception &x)
    {
        cerr << x.what() << endl;
    }

    return 0;
}

