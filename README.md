# xlOil

xlOil is a framework for building Excel language bindings. More succinently, a way to write functions in a language of your choice and have them appear in Excel.

xlOil supports different languages via plugins. The languages currently supported are
- C++
- Python

You can use xlOil as an end-user of these plugins or you can use it to write you own language bindings and contribute!

### Why write xlOil?

This is a section for people thinking about writing language bindings. If you want to write worksheet functions in a nice language, skip to the plugin documentation. 

Interfacing with Excel is tricky for a general language. You have a choice of poisons:

- C-API - is C and hence unsafe, the API is also old, has some quirks and is missing many features
- COM - more fully-featured but slower and missing some features of C-API. Requires COM binding support in your language or a great deal of pain will be endured
- .Net API - actually sits on top of COM, good but limited to .Net languages


xlOil tries to give you the first two blended in a more friendly fashion and adds:

- Solution to the "how to register a worksheet function without a static DLL entry point" problem
- Object caching
- A framework for converting excel variant types to another language and back
- A loader stub
- Goodwill to all men
