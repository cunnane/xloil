#include "Thunker.h"
#include <xlOil/Register.h>
#include <asmjit/src/asmjit/asmjit.h>
#include <string>
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
  void createArrayOfArgsOnStack(asmjit::x86::Compiler& cc, x86::Gp& stackPtr, size_t startArg, size_t endArg)
  {
    const size_t numArgs = endArg - startArg;
    if (numArgs == 0)
      return;

    const auto ptrSize = (int32_t)sizeof(void*);

    // Get some space on the stack
    x86::Mem args = cc.newStack((unsigned)numArgs * ptrSize, alignof(void*));
    cc.lea(stackPtr, args);

    // Copy function arguments to array on the stack
    for (auto i = (int32_t)startArg; i < (int32_t)endArg; i++)
    {
      x86::Gp arg = cc.newUIntPtr("arg");
      auto offset = (ptrSize * (i - (uint32_t)startArg));
      x86::Mem stackP = x86::ptr(stackPtr, offset);
      cc.setArg(i, arg);
      cc.mov(stackP, arg);
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
    x86::Gp argsPtr = cc.newUIntPtr("argsPtr");
    createArrayOfArgsOnStack(cc, argsPtr, 0, numArgs);

    // Setup the signature to call the target callback
    FuncCallNode* call(
      cc.call(imm((void*)callback), FuncSignatureT<ExcelObj*, const void*, const ExcelObj**>(CallConv::kIdHost)));

    call->setArg(0, imm(data));
    call->setArg(1, argsPtr);

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
    x86::Gp firstArg = cc.newUIntPtr("firstArg");
    x86::Gp argsPtr = cc.newUIntPtr("argsPtr");

    cc.setArg(0, firstArg);
    createArrayOfArgsOnStack(cc, argsPtr, 1, numArgs + 1);

    // Setup the signature to call the target callback. Note the void return type
    // as the function will return it's value by invoking xlAsyncReturn.
    FuncCallNode* call(
      cc.call(imm((void*)callback),
        FuncSignatureT<void, const void*, const ExcelObj*, const ExcelObj**>(CallConv::kIdHost)));

    call->setArg(0, imm(data));
    call->setArg(1, firstArg);
    call->setArg(2, argsPtr);

    // No return from async
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

    XLO_TRACE("Building thunk with {0} arguments", numArgs);

    // Initialise place to hold code before compilation
    CodeHolder codeHolder;
    codeHolder.init(theRunTime.codeInfo());

    // What does this do?
    SimpleErrorHandler errorHandler;
    codeHolder.setErrorHandler(&errorHandler);

    // Initialise JIT compiler
    x86::Compiler cc(&codeHolder);

    // Begin code

    // Declare thunk function signature. Need stdcall for Excel functions
    // Assuming all arguments are xloper12* for now
    auto ptrType = Type::IdOfT<xloper12*>::kTypeId;
    FuncSignatureBuilder signature(CallConv::kIdHostStdCall);
    for (auto i = 0; i < numArgs; i++)
      signature.addArg(ptrType);

    // Normal callbacks should return, async ones will not
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
    {
      CodeHolder code;
      code.init(theRunTime.codeInfo());
      x86::Assembler as(&code);
      as.mov(x86::rax, imm(fromData));
      asmJitWriteCode((uint8_t*)bufferBefore, &code, bufsize);
    }
    {
      CodeHolder code;
      code.init(theRunTime.codeInfo());
      x86::Assembler as(&code);
      as.mov(x86::rax, imm(toData));
      asmJitWriteCode((uint8_t*)bufferAfter, &code, bufsize);
    }
   
    auto found = std::search(thunk, thunk + thunkSize, bufferBefore, bufferBefore + bufsize);
    if (found == thunk + thunkSize)
      return false;

    memcpy(found, bufferAfter, bufsize);
    return true;
  }
}