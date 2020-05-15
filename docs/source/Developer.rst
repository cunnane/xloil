=======================
xlOil Development Guide
=======================

Getting started as developer
----------------------------

- You need Visual Studio 2017 or newer
- All xlOil_Core dependencies are already in the `external` folder. 
  Some of them are compressed, so unpack them.
- To build xlOil_Python you need to set the correct paths in 
  `PySettings.props`, `PySettings36.props` and `PySettings37.props` to point to your python distribution.
- For debugging, set xlOil_Loader as the target project, with 
  command=`<Path-to-Excel.exe>` args=`$(OutDir)\xloil.xll`

Why use the xlOil library?
--------------------------

Interfacing with Excel is tricky. You have a choice of poisons:

- C-API - is C and hence unsafe, the API is also old, has some quirks 
  and is missing many features of the other APIs
- COM - more fully-featured but slower and missing some features of 
  C-API. Has some unexpected behaviour and may fail to respond. 
  Requires COM binding support in your language or a great deal of 
  pain will be endured.
- .Net API - actually sits on top of COM, good but limited to .Net 
  languages

xlOil tries to give you the first two APIs blended in a more friendly 
fashion and adds:

- Solution to the "how to register a worksheet function without a static DLL entry point" problem
- Object caching
- A framework for converting excel variant types to another language and back
- A convenient way of creating worksheet-scope functions
- A loader stub
- Goodwill to all men

Of course, there are already existing solutions to for Excel add-in
development and if you want to use .Net languages, I suggest using
one of them as xlOil does not currently support them.  

Most available packages are commercial (and quite expensive) so I 
cannot test relative performance with xlOil. In addtion they are 
focussed on a single language rather than trying to provide a framework
for language bindings as xlOil does.

The following, likely incomplete, list includes some of the most  
prominent Excel addin software. 

- Addin express: (commercial) fully-featured with nice GUI. Limited to 
  .Net languages
- ExcelDNA: (free) mature, well-regarded and widely used, covers 
  almost all Excel API features, but only for .Net languages
- XLL Plus: (commercial) seems to be fully-featured with GUI wizards
  to help developers, but only for C++ and the most expensive sofware 
  here.
- PyXLL: (commercial) Python-only but supports the full range of Excel
  API features and some functionality to run on remote servers.
- XlWings: (mostly free) Python-only. More mature sofware, but considerably
  slower (2000x in my test case) than xlOil due to use of slower APIs.
  Can only create 'local' functions backed by VBA, not viable for addin
  deployment. Supports Macs and Python 2.7 (but licence fee required 
  for this).
