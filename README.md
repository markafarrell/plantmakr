plantmakr
=========

Requirements:
---------
  1. ObjectARX (http://usa.autodesk.com/adsk/servlet/item?siteID=123112&id=785550)
  2. stlsoft (http://www.stlsoft.org/)
  3. vole (http://vole.sourceforge.net/)
  4. boost (http://www.boost.org/)
  5. Visual Studio (https://www.visualstudio.com/post-download-vs?sku=community&clcid=0x409)
  
  Note Use Visual Studio Community Update 3 for Autocad 2018
  
Build Steps:
---------
  1. Download and install the latest versions of the required libraries into the correct directiories.
  2. Update lib/stlsoft/include/stlsoft.h to support your version of MSVC
  e.g.
  ~~~
  # elif (_MSC_VER == 1914)
  #  define STLSOFT_COMPILER_VERSION_STRING       "Visual C++ 14.1"
  ~~~
  3. Build boost libraries (x64 and w32 versions). 
  ```
    .\bootstrap.bat --with-libraries=regex
    .\b2 --toolset=msvc-14.2 address-model=64 --with-regex --build-type=complete --stagedir=lib\x64 stage
  ```
  4. Open the project in Visual Studio
  5. Build the project for the desired architecture.
