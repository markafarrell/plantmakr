/* /////////////////////////////////////////////////////////////////////////
 * File:        recls.com.test.cpp
 *
 * Purpose:     Implementation file for the recls.com.test project.
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
 * ////////////////////////////////////////////////////////////////////// */

//#define VOLE_NO_NAMESPACE
/* VOLE Header Files */
#include <vole/vole.hpp>

/* STLSoft Header Files */
#include <stlsoft/smartptr/ref_ptr.hpp>
#include <comstl/collections/enumerator_sequence.hpp>
#include <comstl/conversion/interface_cast.hpp>
#include <comstl/error/errorinfo_desc.hpp>
#include <comstl/string/bstr.hpp>
#include <comstl/util/creation_functions.hpp>
#include <comstl/util/initialisers.hpp>
#include <comstl/util/value_policies.hpp>
#include <winstl/conversion/char_conversions.hpp>
#include <winstl/error/error_desc.hpp>
#include <winstl/shims/access/string/time.hpp>
#include <winstl/shims/conversion/to_SYSTEMTIME.hpp>

/* Standard C++ Header Files */
#include <exception>
#include <iostream>
#include <iomanip>

using std::cout;
using std::cerr;
using std::endl;

/* Standard C Header Files */
#include <stdio.h>
#include <stdlib.h>

#if defined(_MSC_VER) && \
    defined(_DEBUG)
# include <crtdbg.h>
#endif /* _MSC_VER) && _DEBUG */

/* /////////////////////////////////////////////////////////////////////////
 * Namespace
 */

#ifndef VOLE_NO_NAMESPACE
using vole::collection;
using vole::vole_exception;
using vole::object;
using vole::of_type;
#endif /* !VOLE_NO_NAMESPACE */

/* /////////////////////////////////////////////////////////////////////////
 * Forward declarations
 */

static void usage(int bExit, char const* reason, int invalidArg, int argc, char **argv);
static int  processRoot(object search, char const* rootDir, char const* pattern, long flags, int bSuccinct/* , IFileSearchProgress *fsp */);
static int  processCollection(collection files, long flags, int bSuccinct, int bFTP);

/* ////////////////////////////////////////////////////////////////////// */

enum RECLS_FLAG
{
        RECLS_F_FILES                       =   0x00000001  /*!< Include files in search. Included by default if none specified */
    ,   RECLS_F_DIRECTORIES                 =   0x00000002  /*!< Include directories in search */
    ,   RECLS_F_LINKS                       =   0x00000004  /*!< Include links in search. Ignored in Win32 */
    ,   RECLS_F_DEVICES                     =   0x00000008  /*!< Include devices in search. Not currently supported */
    ,   RECLS_F_TYPEMASK                    =   0x00000FFF
    ,   RECLS_F_DIR_PROGRESS                =   0x00001000  /*!< Reports each traversed directory to the callback function supplied to Recls_SearchFeedback() */
    ,   RECLS_F_RECURSIVE                   =   0x00010000  /*!< Searches given directory and all sub-directories */
    ,   RECLS_F_NO_FOLLOW_LINKS             =   0x00020000  /*!< Does not expand links */
    ,   RECLS_F_DIRECTORY_PARTS             =   0x00040000  /*!< Fills out the directory parts. Supported from version 1.1.1 onwards */
    ,   RECLS_F_DETAILS_LATER               =   0x00080000  /*!< Does not fill out anything other than the path. When specified to Recls_Stat(), and no types (RECLS_F_TYPEMASK) are specified, this will allow some details for a non-existant path to be elicited. */
    ,   RECLS_F_PASSIVE_FTP                 =   0x00100000  /*!< Passive mode in FTP. Supported from version 1.5.1 onwards */
    ,   RECLS_F_MARK_DIRS                   =   0x00200000  /*!< Marks the directories with a trailing slash. */
    ,   RECLS_F_ALLOW_REPARSE_DIRS          =   0x00400000  /*!< Allow Win32 reparse point directories to be examined (which can cause infinite loops). */
#if 0
    ,   RECLS_F_CALC_CHECKSUM               =   0x00800000
#endif /* 0 */
    ,   RECLS_F_CALLBACKS_STDCALL_ON_WIN32  =   0x01000000  /*!< Treats the callbacks as <b>stdcall</b> calling convention on Win32. Default is <b>cdecl</b>. This is useful for interfacing as delegates with .NET */
    ,   RECLS_F_USE_TILDE_ON_NO_SEARCHROOT  =   0x04000000  /*!< Interprets a NULL or empty searchRoot as the home directory, rather than the current directory. Supported from version 1.8.1 onwards. */
#if 0
    ,   RECLS_F_DIR_SIZE_IS_NUM_FILES       =   0x02000000  /*!< This causes the size of the directory to be the number of files contained within it, rather than being 0. */
#endif /* 0 */
    ,   RECLS_F_IGNORE_HIDDEN_ENTRIES_ON_WIN32   =   0x08000000  /*!< This causes hidden files to be ignored. Currently supported on Win32 only. */

#if !defined(FILES)
    ,   FILES = RECLS_F_FILES /*!< Equivalent to \link recls::RECLS_F_FILES RECLS_F_FILES\endlink. */
#endif /* !FILES */
#if !defined(DIRECTORIES)
    ,   DIRECTORIES = RECLS_F_DIRECTORIES /*!< Equivalent to \link recls::RECLS_F_DIRECTORIES RECLS_F_DIRECTORIES\endlink. */
#endif /* !DIRECTORIES */
#if !defined(LINKS)
    ,   LINKS = RECLS_F_LINKS /*!< Equivalent to \link recls::RECLS_F_LINKS RECLS_F_LINKS\endlink. */
#endif /* !LINKS */
#if !defined(DEVICES)
    ,   DEVICES = RECLS_F_DEVICES /*!< Equivalent to \link recls::RECLS_F_DEVICES RECLS_F_DEVICES\endlink. */
#endif /* !DEVICES */
#if !defined(TYPEMASK)
    ,   TYPEMASK = RECLS_F_TYPEMASK /*!< Equivalent to \link recls::RECLS_F_TYPEMASK RECLS_F_TYPEMASK\endlink. */
#endif /* !TYPEMASK */
#if !defined(DIR_PROGRESS)
    ,   DIR_PROGRESS = RECLS_F_DIR_PROGRESS /*!< Equivalent to \link recls::RECLS_F_DIR_PROGRESS RECLS_F_DIR_PROGRESS\endlink. */
#endif /* !DIR_PROGRESS */
#if !defined(RECURSIVE)
    ,   RECURSIVE = RECLS_F_RECURSIVE /*!< Equivalent to \link recls::RECLS_F_RECURSIVE RECLS_F_RECURSIVE\endlink. */
#endif /* !RECURSIVE */
#if !defined(NO_FOLLOW_LINKS)
    ,   NO_FOLLOW_LINKS = RECLS_F_NO_FOLLOW_LINKS /*!< Equivalent to \link recls::RECLS_F_NO_FOLLOW_LINKS RECLS_F_NO_FOLLOW_LINKS\endlink. */
#endif /* !NO_FOLLOW_LINKS */
#if !defined(DIRECTORY_PARTS)
    ,   DIRECTORY_PARTS = RECLS_F_DIRECTORY_PARTS /*!< Equivalent to \link recls::RECLS_F_DIRECTORY_PARTS RECLS_F_DIRECTORY_PARTS\endlink. */
#endif /* !DIRECTORY_PARTS */
#if !defined(DETAILS_LATER)
    ,   DETAILS_LATER = RECLS_F_DETAILS_LATER /*!< Equivalent to \link recls::RECLS_F_DETAILS_LATER RECLS_F_DETAILS_LATER\endlink. */
#endif /* !DETAILS_LATER */
#if !defined(PASSIVE_FTP)
    ,   PASSIVE_FTP = RECLS_F_PASSIVE_FTP /*!< Equivalent to \link recls::RECLS_F_PASSIVE_FTP RECLS_F_PASSIVE_FTP\endlink. */
#endif /* !PASSIVE_FTP */
#if !defined(MARK_DIRS)
    ,   MARK_DIRS = RECLS_F_MARK_DIRS /*!< Equivalent to \link recls::RECLS_F_MARK_DIRS RECLS_F_MARK_DIRS\endlink. */
#endif /* !MARK_DIRS */
#if !defined(ALLOW_REPARSE_DIRS)
    ,   ALLOW_REPARSE_DIRS = RECLS_F_ALLOW_REPARSE_DIRS /*!< Equivalent to \link recls::RECLS_F_ALLOW_REPARSE_DIRS RECLS_F_ALLOW_REPARSE_DIRS\endlink. */
#endif /* !ALLOW_REPARSE_DIRS */
#if 0
#if !defined(CALC_CHECKSUM)
    ,   CALC_CHECKSUM = RECLS_F_CALC_CHECKSUM /*!< Equivalent to \link recls::RECLS_F_CALC_CHECKSUM RECLS_F_CALC_CHECKSUM\endlink. */
#endif /* !CALC_CHECKSUM */
#endif /* 0 */
#if !defined(CALLBACKS_STDCALL_ON_WIN32)
    ,   CALLBACKS_STDCALL_ON_WIN32 = RECLS_F_CALLBACKS_STDCALL_ON_WIN32 /*!< Equivalent to \link recls::RECLS_F_CALLBACKS_STDCALL_ON_WIN32 RECLS_F_CALLBACKS_STDCALL_ON_WIN32\endlink. */
#endif /* !CALLBACKS_STDCALL_ON_WIN32 */
/*  ,   DIR_SIZE_IS_NUM_CHILDREN    =   RECLS_F_DIR_SIZE_IS_NUM_CHILDREN    */
#if !defined(F_USE_TILDE_ON_NO_SEARCHROOT)
    ,   USE_TILDE_ON_NO_SEARCHROOT = RECLS_F_USE_TILDE_ON_NO_SEARCHROOT /*!< Equivalent to \link recls::RECLS_F_USE_TILDE_ON_NO_SEARCHROOT RECLS_F_USE_TILDE_ON_NO_SEARCHROOT\endlink. */
#endif /* !RECLS_F_USE_TILDE_ON_NO_SEARCHROOT */
    ,   IGNORE_HIDDEN_ENTRIES_ON_WIN32 = RECLS_F_IGNORE_HIDDEN_ENTRIES_ON_WIN32 /*!< Equivalent to \link recls::RECLS_F_IGNORE_HIDDEN_ENTRIES_ON_WIN32 RECLS_F_IGNORE_HIDDEN_ENTRIES_ON_WIN32\endlink. */
#if !defined(IGNORE_HIDDEN_ENTRIES_ON_WIN32)
#endif /* !IGNORE_HIDDEN_ENTRIES_ON_WIN32 */
#if 0
#if !defined(RECLS_F_DIR_SIZE_IS_NUM_FILES)
    ,   DIR_SIZE_IS_NUM_FILES = RECLS_F_DIR_SIZE_IS_NUM_FILES /*!< Equivalent to \link recls::RECLS_F_DIR_SIZE_IS_NUM_FILES RECLS_F_DIR_SIZE_IS_NUM_FILES\endlink. */
#endif /* !RECLS_F_DIR_SIZE_IS_NUM_FILES */
#endif /* 0 */
};

/* /////////////////////////////////////////////////////////////////////////
 * main()
 */

static int main_(int argc, char *argv[])
{
    int             iRet            =   EXIT_SUCCESS;
    int             totalFound      =   0;
    int             bAllHardDrives  =   false;
    char const      *host           =   NULL;
    char const      *username       =   NULL;
    char const      *password       =   NULL;
    char const      *pattern        =   NULL;
    char const      *rootDir        =   NULL;
    long            flags           =   RECLS_F_RECURSIVE;
    int             bSuccinct       =   false;

    { for(int i = 1; i < argc; ++i)
    {
        char const  *const  arg =   argv[i];

        if(arg[0] == '-')
        {
            if(arg[1] == '-')
            {
                /* -- arguments */
                if(0 == ::strcmp(arg, "--home"))
                {
                    rootDir =   "";
                    flags   |=  RECLS_F_USE_TILDE_ON_NO_SEARCHROOT;
                }
                else
                {
                    /* -- arguments */
                    usage(1, "Invalid argument(s) specified", i, argc, argv);
                }
            }
            else
            {
                /* - arguments */
                switch(arg[1])
                {
                    case    '?':
                        usage(1, NULL, -1, argc, argv);
                        break;
                    case    'R':    /* Do not recurse */
                        flags &= ~(RECLS_F_RECURSIVE);
                        break;
                    case    'p':    /* Show directory parts */
                        flags |= RECLS_F_DIRECTORY_PARTS;
                        break;
                    case    'f':    /* Find files */
                        flags |= RECLS_F_FILES;
                        break;
                    case    'd':    /* Find directories */
                        flags |= RECLS_F_DIRECTORIES;
                        break;
                    case    'h':    /* Searches from hard drives */
                        bAllHardDrives = true;
                        break;
                    case    'm':    /* Mark directories with a trailing slash */
                        flags |= RECLS_F_MARK_DIRS;
                        break;
                    case    'r':    /* Allow reparse points */
                        flags |= RECLS_F_ALLOW_REPARSE_DIRS;
                        break;
                    case    's':    /* Show only the full path; WHEREIS functionality */
                        bSuccinct = true;
                        break;
                    case    'H':    /* FTP host */
                        host = arg + 2;
                        break;
                    case    'U':    /* FTP username */
                        username = arg + 2;
                        break;
                    case    'P':    /* FTP password */
                        password = arg + 2;
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
            if(NULL == pattern)
            {
                pattern = arg;
            }
            else if(NULL == rootDir)
            {
                rootDir = arg;
            }
            else
            {
                usage(1, "Invalid argument(s) specified", i, argc, argv);
            }
        }
    }}

    if( NULL != password &&
        (   NULL == host ||
            NULL == username))
    {
        usage(1, "Must specify host and username if specifying password", -1, argc, argv);
    }

    if( NULL != username &&
        NULL == host)
    {
        usage(1, "Must specify host if specifying username", -1, argc, argv);
    }

    if( NULL != host &&
        bAllHardDrives)
    {
        usage(1, "-h flag meaningless to FTP searches", -1, argc, argv);
    }

    if( NULL != host &&
        NULL == rootDir)
    {
        usage(1, "Must specify a root directory for FTP searches", -1, argc, argv);
    }

    // Search for files if neither files or directories specified
    //
    // Even though this is not necessary, because the recls API provides the
    // same interpretation, it's best to be explicit.
    if(0 == (flags & (RECLS_F_FILES | RECLS_F_DIRECTORIES)))
    {
        flags |= RECLS_F_FILES;
    }

    if(NULL == rootDir)
    {
        rootDir = ".";
    }

    if(NULL != host)
    {
    }
    else
    {
        object search = object::create("recls.COM.FileSearch");

        if(bAllHardDrives)
        {
        }
        else
        {
            totalFound = processRoot(search, rootDir, pattern, flags, bSuccinct/* , &s_fsp */);
        }
    }

    return iRet;
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
        puts("recls.com.test: " STLSOFT_COMPILER_LABEL_STRING);
#endif /* debug */

        comstl::com_initialiser coinit;

        res = main_(argc, argv);
    }
    catch(vole_exception &x)
    {
        fprintf(stderr, "Unexpected exception: %s: 0x%08x, %s\n", x.what(), x.get_hr(), winstl::error_desc(x.get_hr()).c_str());

        res = EXIT_FAILURE;
    }
    catch(std::exception &x)
    {
        fprintf(stderr, "Unexpected exception: %s\n", x.what());

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

/* ////////////////////////////////////////////////////////////////////// */

static void usage(int bExit, char const* reason, int invalidArg, int argc, char **argv)
{
    FILE    *stm    =   (NULL == reason) ? stdout : stderr;


    fprintf(stm, "\n");
    fprintf(stm, "  recls.com.test\n\n");
    fprintf(stm, "\n");

    if(NULL != reason)
    {
        fprintf(stm, "  Error: %s\n", reason);
        fprintf(stm, "\n");
    }

    if(0 < invalidArg)
    {
        fprintf(stm, "  First invalid argument #%d: %s\n", invalidArg, argv[invalidArg]);
        fprintf(stm, "  Arguments were (first bad marked *):\n\n");
        { int i; for(i = 1; i < argc; ++i)
        {
            fprintf(stm, "  %s%s\n", (i == invalidArg) ? "* " : "  ", argv[i]);
        }}
        fprintf(stm, "\n");
    }

#if 0
    fprintf(stm, "  USAGE: recls.com.test [{-w | -p<root-paths> | -h}] [-u] [-d] [{-v | -s}] <search-spec>\n");
    fprintf(stm, "    where:\n\n");
    fprintf(stm, "    -w             - searches from the current working directory. The default\n");
    fprintf(stm, "    -p<root-paths> - searches from the given root path(s), separated by \';\',\n");
    fprintf(stm, "                     eg.\n");
    fprintf(stm, "                       -p\"c:\\windows;x:\\bin\"\n");
    fprintf(stm, "    -r             - deletes readonly files\n");
    fprintf(stm, "    -h             - searches from the roots of all drives on the system\n");
    fprintf(stm, "    -d             - displays the path(s) searched\n");
    fprintf(stm, "    -u             - do not act recursively\n");
    fprintf(stm, "    -v             - verbose output. Prints time, attributes, size and path. (default)\n");
    fprintf(stm, "    -s             - succinct output. Prints path only\n");
    fprintf(stm, "    <search-spec>  - one or more file search specifications, separated by \';\',\n");
    fprintf(stm, "                     eg.\n");
    fprintf(stm, "                       \"*.exe\"\n");
    fprintf(stm, "                       \"myfile.ext\"\n");
    fprintf(stm, "                       \"*.exe;*.dll\"\n");
    fprintf(stm, "                       \"*.xl?;report.*\"\n");
    fprintf(stm, "\n");
    fprintf(stm, "  Contact %s\n", _nameSynesisSoftware);
    fprintf(stm, "    at \"%s\",\n", _wwwSynesisSoftware_SystemTools);
    fprintf(stm, "    or, via email, at \"%s\"\n", _emailSynesisSoftware_SystemTools);
    fprintf(stm, "\n");
#endif /* 0 */

    if(bExit)
    {
        exit(EXIT_FAILURE);
    }
}

static int processRoot(object search, char const* rootDir, char const* pattern, long flags, int bSuccinct/* , IFileSearchProgress *fsp */)
{
    collection  files       =   search.invoke_method(of_type<collection>(), L"Search", rootDir, pattern, comstl::variant(flags));
    int         totalFound  =   processCollection(files, flags, bSuccinct, false);

    return totalFound;
}

static int processCollection(
    collection  files
,   long        /* flags */
,   int         bSuccinct
,   int         /* bFTP */
)
{
    typedef std::string string_t;

    int totalFound = 0;

    { for(collection ::iterator b = files.begin(); b != files.end(); ++b)
    {
        ++totalFound;

#if 0
        s_fsp.Reset();
#endif /* 0 */

        object      fe(*b);
        string_t    path                    =   fe.get_prop(of_type<string_t>(), L"Path");
        char        drive                   =   char(fe.get_prop(of_type<comstl::variant>(), L"Drive").convert(VT_I2).iVal);
        string_t    directory               =   fe.get_prop(of_type<string_t>(), L"Directory");
        string_t    directoryPath           =   fe.get_prop(of_type<string_t>(), L"DirectoryPath");
        string_t    searchRelativePath      =   fe.get_prop(of_type<string_t>(), L"SearchRelativePath");
        string_t    uncDrive                =   fe.get_prop(of_type<string_t>(), L"UNCDrive");
        collection  directoryParts          =   fe.get_prop(of_type<collection>(), L"DirectoryParts");
        string_t    file                    =   fe.get_prop(of_type<string_t>(), L"File");
        string_t    shortFile               =   fe.get_prop(of_type<string_t>(), L"ShortFile");
        string_t    fileName                =   fe.get_prop(of_type<string_t>(), L"FileName");
        string_t    fileExt                 =   fe.get_prop(of_type<string_t>(), L"FileExt");
        SYSTEMTIME  creationTime            =   fe.get_prop(of_type<SYSTEMTIME>(), L"CreationTime");
        SYSTEMTIME  modificationTime        =   fe.get_prop(of_type<SYSTEMTIME>(), L"ModificationTime");
        SYSTEMTIME  lastAccessTime          =   fe.get_prop(of_type<SYSTEMTIME>(), L"LastAccessTime");
        SYSTEMTIME  lastStatusChangeTime    =   fe.get_prop(of_type<SYSTEMTIME>(), L"LastStatusChangeTime");
        long        size                    =   fe.get_prop(of_type<long>(), L"Size");
        bool        isReadOnly              =   fe.get_prop(of_type<bool>(), L"IsReadOnly");
        bool        isDirectory             =   fe.get_prop(of_type<bool>(), L"IsDirectory");
        bool        isLink                  =   fe.get_prop(of_type<bool>(), L"IsLink");
        bool        isUNC                   =   fe.get_prop(of_type<bool>(), L"IsUNC");

        cout << "  " << path << endl;

        if(!bSuccinct)
        {
            int off = 0;

            if(!searchRelativePath.empty())
            {
                cout << "  " << std::setw(path.size()) << searchRelativePath << endl;
            }

            cout << "  " << directoryPath << endl;

            if(isUNC)
            {
                cout << "  " << uncDrive << endl;
                off += int(uncDrive.size());
            }
            else
            {
                cout << "  " << drive << endl;
                off += 2;
            }
            cout << "  " << std::setw(off + directory.size()) << directory << std::setw(1) << endl;

#if 0
            { for(collection::iterator pi = directoryParts.begin(); pi != directoryParts.end(); ++pi)
            {
                object      part(*pi);
                string_t    dirPart(vPart.bstrVal);

                off += dirPart.Length();
                cout << "  " << std::setw(off) << dirPart << std::setw(1) << endl;
            }}
#endif /* 0 */

            off += int(file.size());
            cout << "  " << std::setw(off) << file << std::setw(1) << endl;
            if(fileExt.size() > 0)
            {
                cout << "  " << std::setw(off - (1 + fileExt.size())) << fileName << std::setw(1) << endl;
                cout << "  " << std::setw(off) << fileExt << std::setw(1) << endl;
            }
            else
            {
                cout << "  " << std::setw(off) << fileName << std::setw(1) << endl;
            }

            cout << "  size: " << size << endl;

            cout << "  created: " << stlsoft::c_str_ptr_a(creationTime) << endl;
            cout << "  modified: " << stlsoft::c_str_ptr_a(modificationTime) << endl;
            cout << "  accessed: " << stlsoft::c_str_ptr_a(lastAccessTime) << endl;
            cout << "  status: " << stlsoft::c_str_ptr_a(lastStatusChangeTime) << endl;

            if(isReadOnly)
            {
                cout << "    - Read-only" << endl;
            }
            if(isDirectory)
            {
                cout << "    - Directory" << endl;
            }
            else
            {
                cout << "    - File" << endl;
            }
            if(isLink)
            {
                cout << "    - Link" << endl;
            }
            if(isUNC)
            {
                cout << "    - UNC" << endl;
            }
        }
    }}

    return totalFound;
}

/* ///////////////////////////// end of file //////////////////////////// */
