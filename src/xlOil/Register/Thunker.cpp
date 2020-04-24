#include "Thunker.h"
#include <xlOil/Register.h>
#include <xlOil/Throw.h>
#include <xlOil/Log.h>
#include <asmjit/src/asmjit/asmjit.h>
#include <string>
#include <algorithm>
using std::string;

class xloper12;

namespace xloil {
  class ExcelObj;
}

namespace
{
  using namespace asmjit;

  static JitRuntime theRunTime;

  // This is jitRuntime.add() but with in-place allocation
  Error asmJitWriteCode(uint8_t* dst, CodeHolder* code, size_t& codeSize) noexcept
  {
    ASMJIT_PROPAGATE(code->flatten());
    ASMJIT_PROPAGATE(code->resolveUnresolvedLinks());

    size_t estimatedCodeSize = code->codeSize();
    if (estimatedCodeSize == 0)
      return DebugUtils::errored(kErrorNoCodeGenerated);

    // Relocate the code.
    Error err = code->relocateToBase(uintptr_t((void*)dst));
    if (err)
      return err;

    // Recalculate the final code size and shrink the memory we allocated for it
    // in case that some relocations didn't require records in an address table.
    codeSize = code->codeSize();

    for (Section* section : code->_sections) {
      size_t offset = size_t(section->offset());
      size_t bufferSize = size_t(section->bufferSize());
      size_t virtualSize = size_t(section->virtualSize());

      ASMJIT_ASSERT(offset + bufferSize <= codeSize);
      memcpy(dst + offset, section->data(), bufferSize);

      if (virtualSize > bufferSize) {
        ASMJIT_ASSERT(offset + virtualSize <= codeSize);
        memset(dst + offset + bufferSize, 0, virtualSize - bufferSize);
      }
    }

    return kErrorOk;
  }

  class SimpleErrorHandler : public asmjit::ErrorHandler {
  public:
    SimpleErrorHandler() : _err(asmjit::kErrorOk) {}
    virtual void handleError(asmjit::Error err, const char* message, asmjit::BaseEmitter* origin) {
      ASMJIT_UNUSED(origin);
      _err = err;
      _message = message;
    }

    asmjit::Error _err;
    string _message;
  };

}
namespace xloil
{
  void createArrayOfArgsOnStack(asmjit::x86::Compiler& cc, x86::Mem& stackPtr, size_t startArg, size_t endArg)
  {
    const size_t numArgs = endArg - startArg;
    if (numArgs == 0)
      return;

    const auto ptrSize = (int32_t)sizeof(void*);

    // Get some space on the stack
    stackPtr = cc.newStack((unsigned)numArgs * ptrSize, alignof(void*));

    // Copy function arguments to array on the stack
    for (auto i = (int32_t)startArg; i < (int32_t)endArg; i++)
    {
      const auto offset = (ptrSize * (i - (uint32_t)startArg));
      x86::Mem stackPos = stackPtr.cloneAdjusted(offset);
      x86::Gp arg = cc.newUIntPtr("arg");
      cc.setArg(i, arg);
      cc.mov(stackPos, arg);
    }
  }

  void writeFunctionBody(
    asmjit::x86::Compiler& cc,
    RegisterCallback callback,
    const void* data,
    size_t numArgs)
  {
    // Take args passed to thunk and put them into an array on the stack
    // This should give us an xloper12** which we load into argsPtr
    x86::Mem argsPtr;
    createArrayOfArgsOnStack(cc, argsPtr, 0, numArgs);
    
    // For an x64 call, there are only two arguments and we explictly place 
    // them in rcx and rdx as per the calling convention.
    //
    // We do this manually as asmjit's register allocator is sub-optimal
    // and will generate additional movs and spills.
    //
#ifdef _WIN64
    cc.mov(x86::rcx, imm(data));
    cc.lea(x86::rdx, argsPtr);
#endif

    // Setup the signature to call the target callback, this is not a stdcall
    // that was only for the call from Excel to xlOil
    FuncCallNode* call(
      cc.call(imm((void*)callback), 
        FuncSignatureT<ExcelObj*, const void*, const ExcelObj**>(CallConv::kIdHost)));

    // For x86 we rely on asmjit to handle the calling convention and
    // register allocation.
    //
#ifndef _WIN64 
    x86::Gp arg1 = cc.newUIntPtr("arg1");
    x86::Gp arg2 = cc.newUIntPtr("arg2");
    cc.mov(arg1, imm(data));
    cc.lea(arg2, argsPtr);
    call->setArg(0, arg1);
    call->setArg(1, arg2);
#endif

    // Allocate a register for the return value, in fact 
    // this is a no-op as we just return the callback value
    x86::Gp ret = cc.newUIntPtr("ret");

    // TODO: any chance of setting xl-free bit here?
    call->setRet(0, ret);

    // Pass callback return as our return
    cc.ret(ret);
  }

  void writeFunctionBody(
    asmjit::x86::Compiler& cc,
    AsyncCallback callback,
    const void* data,
    size_t numArgs)
  {
    // Take args passed to thunk and put them into an array on the stack
    // This should give us an xloper12** which we load into argsPtr.
    // We separate out the first argument as this will contain the async handle
    // which needs to be returned to Excel.
    x86::Mem argsPtr;
    x86::Gp handle = cc.newUIntPtr("handle");

    cc.setArg(0, handle); // Will be rcx on x64
    createArrayOfArgsOnStack(cc, argsPtr, 1, numArgs + 1);

    // For an x64 call, there are only three arguments and we explictly  
    // place them in rcx, rdx and r8 as per the calling convention.
    //
    // We do this explictly as asmjit's register allocator seems to get
    // confused and spill and mov things all over.
    //
#ifdef _WIN64
    cc.mov(x86::rdx, handle);
    cc.mov(x86::rcx, imm(data));
    cc.lea(x86::r8, argsPtr);
#endif

    // Setup the signature to call the target callback. Note the void return type
    // as the function will return it's value by invoking xlAsyncReturn.
    FuncCallNode* call(
      cc.call(imm((void*)callback),
        FuncSignatureT<void, const void*, const ExcelObj*, const ExcelObj**>(CallConv::kIdHost)));

    // For x86 we rely on asmjit to handle the calling convention and
    // register allocation.
    //
#ifndef _WIN64
    call->setArg(0, handle);
    call->setArg(1, imm(data));
    x86::Gp args = cc.newUIntPtr("args");
    cc.lea(args, argsPtr);
    call->setArg(2, args);
#endif

    // No return from async, this is a no-op.
    cc.ret();
  }


  template <class TCallback>
  void* buildThunk(
    TCallback callback,
    const void* data,
    const size_t numArgs,
    char* codeBuffer,
    size_t bufferSize,
    size_t& codeSize)
  {
    using namespace asmjit;

    XLO_DEBUG("Building thunk with {0} arguments", numArgs);

    // Initialise place to hold code before compilation
    CodeHolder codeHolder;
    codeHolder.init(theRunTime.codeInfo());

    // What does this do?
    SimpleErrorHandler errorHandler;
    codeHolder.setErrorHandler(&errorHandler);

    // Initialise JIT compiler
    x86::Compiler cc(&codeHolder);

    // Begin code

    // Declare thunk function signature: need stdcall for Excel functions.
    // We assume all arguments are xloper12*.
    auto ptrType = Type::IdOfT<xloper12*>::kTypeId;
    FuncSignatureBuilder signature(CallConv::kIdHostStdCall);
    for (size_t i = 0; i < numArgs; i++)
      signature.addArg(ptrType);

    // Normal callbacks should return, async ones will not, so set return
    // type appropriately.
    if (!std::is_same<TCallback, AsyncCallback>::value)
      signature.setRet(ptrType);
    else
    {
      signature.addArg(ptrType);
      signature.setRet(Type::kIdVoid);
    }

    cc.addFunc(signature);

    // Write the appropriate function body for the callback type
    writeFunctionBody(cc, callback, data, numArgs);

    cc.endFunc();

    auto err = cc.finalize();
    if (err)
      XLO_THROW("Thunk compilation failed: {0}", DebugUtils::errorAsString(err));

    // TODO: run compactThunks first?
    if (codeHolder.codeSize() > bufferSize)
      XLO_THROW("Cannot write thunk: buffer exhausted");

    err = asmJitWriteCode((uint8_t*)codeBuffer, &codeHolder, codeSize);
    if (err != kErrorOk)
      XLO_THROW("Thunk write failed: {0}", DebugUtils::errorAsString(err));

    // We need to get permissions to write to the code cave, since it's in the
    // executable part of the program it won't be writeable by default.
    DWORD dummy;
    if (!VirtualProtect(codeBuffer, codeSize, PAGE_EXECUTE_READWRITE, &dummy))
      XLO_THROW(writeWindowsError());
    return codeBuffer;
  }

  // Explicitly instantiate buildThunk;
  template void* buildThunk<RegisterCallback>(RegisterCallback, const void*, const size_t, char*, size_t, size_t&);
  template void* buildThunk<AsyncCallback>(AsyncCallback, const void*, const size_t, char*, size_t, size_t&);

  bool patchThunkData(char* thunk, size_t thunkSize, const void* fromData, const void* toData)
  {
    if (fromData == toData)
      return true;

    char bufferBefore[10], bufferAfter[10];
    auto bufsize = sizeof(bufferBefore);
    // TODO: This will only work in 64-bits as asmjit will load the data 
    // into another register, maybe ecx, but we can't be sure.
    // Probably better to just scan memory for the mov instruction
    {
      CodeHolder code;
      code.init(theRunTime.codeInfo());
      x86::Assembler as(&code);
      as.mov(x86::rcx, imm(fromData));
      asmJitWriteCode((uint8_t*)bufferBefore, &code, bufsize);
    }
    {
      CodeHolder code;
      code.init(theRunTime.codeInfo());
      x86::Assembler as(&code);
      as.mov(x86::rcx, imm(toData));
      asmJitWriteCode((uint8_t*)bufferAfter, &code, bufsize);
    }
   
    auto found = std::search(thunk, thunk + thunkSize, bufferBefore, bufferBefore + bufsize);
    if (found == thunk + thunkSize)
      return false;

    memcpy(found, bufferAfter, bufsize);
    return true;
  }
}