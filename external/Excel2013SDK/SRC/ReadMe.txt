Changes to XLCALL.CPP in the Microsoft Excel 2013 XLL Software Development Kit (SDK)
====================================================================================

There are no changes to XLCALL.CPP.

What is XLCALL.CPP?
===================

XLCALL.CPP was introduced in the Excel 2007 XLL SDK.
XLCALL.CPP contains the definitions of two Excel entry points -
Excel12 and Excel12v.

Any XLL built starting with the Excel 2007 XLL SDK that uses the Excel12 or Excel12v
entry points must include the code from XLCALL.CPP for successful compilation.

The code from XLCALL.CPP can be included in the following three ways:

1. If you are using Microsoft Visual Studio for development, add XLCALL.CPP to the project.

2. If you are using makefiles to build the XLL, add XLCALL.CPP to the list of .CPP files
to compile.

3. Include XLCALL.CPP in a C or C++ file that contains code for your XLL.
XLCALL.CPP can be included using the following statement:

#include "XLCALL.CPP"

Make sure that XLCALL.CPP is present in your include path.

The first two methods are the preferred way of including the code from XLCALL.CPP.
The third method is illustrated by the framework sample included in the Excel 2010 XLL SDK.

NOTE - It is not necessary to explicitly include the code from XLCALL.CPP if
the XLL links to frmwrk32.lib. frmwrk32.lib already contains this code.