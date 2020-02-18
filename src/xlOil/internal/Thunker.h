#pragma once

namespace xloil
{
  /// <summary>
  /// Builds a thunk like the following
  ///   xloper12* func(xloper12* arg1, xloper12* arg2, ...)
  ///   {
  ///     xloper12* args[numArgs];
  ///     args[0] = arg1; args[1] = arg2; ...
  ///     return callback(contextData, args);
  ///   }
  /// 
  /// Or for async functions:
  ///   xloper12* func(xloper12* handle, xloper12* arg1, xloper12* arg2, ...)
  ///   {
  ///     xloper12* args[numArgs];
  ///     args[0] = arg1; args[1] = arg2; ...
  ///     asyncCallback(contextData, handle, args);
  ///   }
  /// 
  /// In x86_64 assembly, this looks like:
  ///   sub         rsp,38h  
  ///   mov         qword ptr[rsp + 20h], rcx
  ///   mov         qword ptr[rsp + 28h], rdx
  ///   mov         qword ptr[rsp + 30h], r8
  ///   mov         rcx, 23250C10570h
  ///   lea         rdx, [rsp + 20h]
  ///   call        xloil::callback(07FFD47DAA89Ch)
  ///   add         rsp, 38h
  ///   ret
  /// </summary>
  template <class TCallback>
  void* buildThunk(
    TCallback callback,
    const void* contextData,
    const size_t numArgs,
    char* codeBuffer,
    size_t bufferSize,
    size_t& codeSize);

  /// <summary>
  /// Patches the context data object in a given thunk to a new location.
  /// @see buildThunk.
  /// </summary>
  bool patchThunkData(
    char* thunk,
    size_t thunkSize,
    const void* fromData,
    const void* toData);
}