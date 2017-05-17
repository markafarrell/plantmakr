plantmakr
=========

Requirements:
---------
  1. ObjectARX (http://usa.autodesk.com/adsk/servlet/item?siteID=123112&id=785550)
  2. stlsoft (http://www.stlsoft.org/)
  3. vole (http://vole.sourceforge.net/)
  4. boost (http://www.boost.org/)
  5. Visual Studio (https://www.visualstudio.com/post-download-vs?sku=community&clcid=0x409)
  
Build Steps:
---------
  1. Download and install the latest versions of the required libraries into the correct directiories.
  2. Build boost libraries (x64 and w32 versions). 
    - .\bootstrap.bat
    - .\b2 --toolset=msvc-14.1 address-model=64 --build-type=complete --stagedir=lib\x64 stage
    - .\b2 --toolset=msvc-14.1 --build-type=complete --stagedir=lib\w32 stage
  3. Open the project in Visual Studio
  4. Build the project for the desired architecture.
