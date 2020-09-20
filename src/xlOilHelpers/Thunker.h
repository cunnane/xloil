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
  /// 
  /// In x86 it will be like:
  ///   sub         esp,0Ch  
  ///   mov         eax,dword ptr [esp+10h]  
  ///   mov         dword ptr [esp+8],eax  
  ///   lea         eax,[esp+8]  
  ///   mov         dword ptr [esp],1C565A70h  
  ///   mov         dword ptr [esp+4],eax  
  ///   call        xloil::callback(210FF8E9h)
  ///   add         esp,0Ch  
  ///   ret         4  
  /// </summary>
  template <class TCallback>
  void* buildThunk(
    TCallback callback,
    const void* contextData,
    const size_t numArgs,
    char* codeBuffer,
    size_t bufferSize,
    size_t& codeSize);

#ifdef _WIN64
  /// <summary>
  /// A hand optimised version of buildThunk which runs faster and reduces final binary
  /// size by ~80kb by avoiding the use of asmjit's compiler. The resulting ASM code is
  /// faster by avoiding spills, particuarly for async functions and when number of args
  /// exceeds 5.
  /// (Currently only available under 64-bit)
  /// </summary>
  template <class TCallback>
  void* buildThunkLite(
    TCallback callback,
    const void* data,
    const size_t numArgs,
    char* codeBuffer,
    size_t bufferSize,
    size_t& codeSize);
#endif

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