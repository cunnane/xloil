#pragma once

namespace xloil
{
  template <class TCallback>
  void* buildThunk(
    TCallback callback,
    const void* contextData,
    const size_t numArgs,
    char* codeBuffer,
    size_t bufferSize,
    size_t& codeSize);

  bool patchThunkData(
    char* thunk,
    size_t thunkSize,
    const void* fromData,
    const void* toData);
}