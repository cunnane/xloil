#pragma once

namespace xloil
{
  /// <summary>
  /// Builds a thunk like this:
  /// 53                            push    rbx  
  /// 48 83 EC 40                   sub     rsp, 40h
  /// 48 8D 5C 24 20                lea     rbx, [rsp + 20h]
  /// 48 89 0B                      mov     qword ptr[rbx], rcx
  /// 48 89 53 08                   mov     qword ptr[rbx + 8], rdx
  /// 4C 89 43 10                   mov     qword ptr[rbx + 10h], r8
  /// 48 B8 80 61 01 10 DE 01 00 00 mov     rax, 1DE10016180h
  /// 48 8B C8                      mov     rcx, rax
  /// 48 8B D3                      mov     rdx, rbx
  /// 40 E8 52 E6 8C F3             call    xloil::Python::pythonCallback(07FFD297052F5h)
  /// 48 83 C4 40                   add     rsp, 40h
  /// 5B                            pop     rbx
  /// C3                            ret
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
  /// Patches the data object in a provided thunk to a new location.
  /// </summary>
  bool patchThunkData(
    char* thunk,
    size_t thunkSize,
    const void* fromData,
    const void* toData);
}