#pragma once

namespace asmjit { class CodeHolder; }
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
  class ThunkWriter
  {
  public:
    /// <summary>
    /// In x64, a hand optimised version of buildThunk is used which runs faster
    /// and reduces  final binary size by ~80kb by avoiding the use of asmjit's 
    /// compiler. The resulting ASM code is faster by avoiding spills, 
    /// particuarly for async functions and when number of args exceeds 5.
    /// This can be disabled with the SLOW_BUILD parameter
    /// </summary>
    /// <param name="callback"></param>
    /// <param name="contextData"></param>
    /// <param name="numArgs"></param>
    /// <param name="hasReturnVal"></param>
    ThunkWriter(
      const void* callback,
      const void* contextData,
      const size_t numArgs,
      const bool hasReturnVal);

    enum SlowBuild {SLOW_BUILD};
    ThunkWriter(
      const void* callback,
      const void* contextData,
      const size_t numArgs,
      const bool hasReturnVal,
      SlowBuild);

    ~ThunkWriter();
    /// <summary>
    /// Writes code to provided buffer, returning number of bytes written.
    /// If <param ref="buffer"> is null, returns size of buffer required.
    /// </summary>
    /// <param name="buffer"></param>
    /// <param name="bufSize"></param>
    /// <returns></returns>
    size_t writeCode(char* buffer, size_t bufSize);
    size_t codeSize() const;
    asmjit::CodeHolder* _holder;

  private:
    ThunkWriter(ThunkWriter&) = delete;
  };
  
  /// <summary>
  /// Patches the context data object in a given thunk to a new location.
  /// <see ref="ThunkWriter">
  /// </summary>
  bool patchThunkData(
    char* thunk,
    size_t thunkSize,
    const void* fromData,
    const void* toData) noexcept;
}