#include "Thunker.h"
#include <xlOil/Register.h>
#include <xlOilHelpers/Exception.h>
#include <xlOil/Log.h>

// See asmjit/core/build.h
#define ASMJIT_STATIC
#define ASMJIT_NO_LOGGING
#define ASMJIT_NO_VALIDATION
#define ASMJIT_NO_JIT

// We could also define NO_BUILDER, NO_COMPILER and NO_INTROSPECTION for x64
// but it only saves 3kb in the compiled binary as they are optimised away

#include <asmjit/src/asmjit/asmjit.h>
#include <string>
#include <algorithm>
using std::string;
using xloil::Helpers::Exception;
using std::unique_ptr;

class xloper12;

namespace xloil {
  class ExcelObj;
}

namespace
{
  /// <summary>
  /// We create the asmjit::CodeInfo object ourselves rather than rely on the
  /// JitRuntime object as this saves about 10kb of optimised binary
  /// </summary>
  /// <returns></returns>
  auto createCodeInfo()
  {
    using namespace asmjit;
    CodeInfo info;
    info._archInfo = CpuInfo::host().archInfo();
#ifdef _WIN64
    info._stackAlignment = 16;
#else
    info._stackAlignment = uint8_t(sizeof(uintptr_t));
#endif
    info._cdeclCallConv = CallConv::kIdHostCDecl;
    info._stdCallConv = CallConv::kIdHostStdCall;
    info._fastCallConv = CallConv::kIdHostFastCall;
    return info;
  }

  static asmjit::CodeInfo theCodeInfo = createCodeInfo();

  using namespace asmjit;

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
}

namespace xloil
{
  // Saves 80kb in the release build :)
  void handRoll64(CodeHolder* code,
    void* callback,
    const void* data,
    size_t numArgs,
    bool hasReturnVal)
  {
    asmjit::x86::Assembler asmb(code);

    XLO_DEBUG("Building thunk with {0} arguments", numArgs);

    // Build the signature of the function we are creating
    FuncSignatureBuilder signature(CallConv::kIdHostStdCall);
    constexpr auto ptrType = Type::IdOfT<int*>::kTypeId;

    for (size_t i = 0; i < numArgs; i++)
      signature.addArg(ptrType);

    // Normal callbacks should return, async ones will not
    signature.setRet(hasReturnVal ? ptrType : Type::kIdVoid);
    
    FuncDetail func;
    func.init(signature);

    // The frame emits the correct function prolog & epilog
    FuncFrame frame;
    frame.init(func);

    // We will need some local stack to create the array of xloper* which 
    // is sent to the callback
    constexpr auto ptrSize = (int32_t)sizeof(void*);
    const auto stackSize = (unsigned)numArgs * ptrSize;
    frame.setLocalStackSize(stackSize);

    // Need to allocate some spill zone for the call to the callback
    frame.updateCallStackSize(func.callConv().spillZoneSize()); 

    frame.finalize();

    // Note we do not preserve the frame pointer as there is litte benefit
    // in debugging the thunk
    asmb.emitProlog(frame);
    
    // See the help for FuncFrame to understand why these stack offsets work.
    // There doesn't seem to be a clean way of getting to the first stack argument
    // without manually skipping the spill zone
    x86::Mem localStack(x86::rsp, frame.localStackOffset());
    x86::Mem stackArgs(x86::rsp, frame.saOffsetFromSP() + frame.spillZoneSize());

    const auto startArg = 0;
    
    // Under x64 Microsoft calling convention the args will be in rcx, rdx, r8, r9
    // with the remainder on the stack. 
    // We copy each of the 4 register arguments to a array in our stack, then copy
    // the remaining stack arguments from earlier in the stack via rax. (rax is 
    // considered volatile so we can clobber it)
    for (size_t i = startArg; i < numArgs; ++i)
    {
      const auto offset = (ptrSize * (i - (uint32_t)startArg));
      auto stackPos = localStack.cloneAdjusted(offset);
      // TODO: this would be nicer as a cascading switch, but it won't write faster code!
      switch (i + 1)
      {
      case 1:
        asmb.mov(stackPos, x86::rcx); break;
      case 2:
        asmb.mov(stackPos, x86::rdx); break;
      case 3:
        asmb.mov(stackPos, x86::r8); break;
      case 4:
        asmb.mov(stackPos, x86::r9); break;
      default:
        asmb.mov(x86::rax, stackArgs.cloneAdjusted((i - 4) * ptrSize));
        asmb.mov(stackPos, x86::rax);
      }
    }

    // Setup arguments for callback
    asmb.lea(x86::rdx, localStack);
    asmb.mov(x86::rcx, imm(data));

    asmb.call(imm((void*)callback));

    // We just pass on the value (an xloper*) returned by the callback
    asmb.emitEpilog(frame);
  }

  void createArrayOfArgsOnStack(
    asmjit::x86::Compiler& cc, x86::Mem& stackPtr, size_t startArg, size_t endArg)
  {
    const size_t numArgs = endArg - startArg;
    if (numArgs == 0)
      return;

    constexpr auto ptrSize = (int32_t)sizeof(void*);

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
    const void* callback,
    const void* data,
    size_t numArgs,
    bool retVal)
  {
    // Take args passed to thunk and put them into an array on the stack
    // This should give us an xloper12** which we load into argsPtr
    x86::Mem argsPtr;
    createArrayOfArgsOnStack(cc, argsPtr, 0, numArgs);
    
    x86::Gp args = cc.newUIntPtr("arg2");
    cc.lea(args, argsPtr);

    // Setup the signature to call the target callback, this is not a stdcall
    // that was only required for the call from Excel to xlOil
    auto sig = FuncSignatureT<void, const void*, const ExcelObj**>(CallConv::kIdHost);
    if (retVal)
      sig._ret = Type::IdOfT<ExcelObj*>::kTypeId;

    FuncCallNode* call(cc.call(imm(callback), sig));

    call->setArg(0, imm(data));
    call->setArg(1, args);

    if (retVal)
    {
      // Allocate a register for the return value, in fact 
      // this is a no-op as we just return the callback value
      x86::Gp ret = cc.newUIntPtr("ret");

      call->setRet(0, ret);

      // Pass callback return as our return
      cc.ret(ret);
    }
  }

  void buildThunk(
    const void* callback,
    const void* data,
    const size_t numArgs,
    const bool hasReturnVal,
    asmjit::CodeHolder& codeHolder)
  {
    using namespace asmjit;

    XLO_DEBUG("Building thunk with {0} arguments", numArgs);

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
    signature.setRet(hasReturnVal ? ptrType : Type::kIdVoid);

    cc.addFunc(signature);

    // Write the appropriate function body for the callback type
    writeFunctionBody(cc, callback, data, numArgs, hasReturnVal);

    cc.endFunc();

    auto err = cc.finalize();
    if (err)
      throw Exception("Thunk compilation failed: %s", DebugUtils::errorAsString(err));
  }

   ThunkWriter::ThunkWriter(
    const void* callback,
    const void* contextData,
    const size_t numArgs,
    const bool hasReturnVal,
    ThunkWriter::SlowBuild)
  {
    _holder = new CodeHolder();
    _holder->init(theCodeInfo);
    buildThunk(callback, contextData, numArgs, hasReturnVal, *_holder);
  }

  ThunkWriter::ThunkWriter(
    const void* callback,
    const void* contextData,
    const size_t numArgs,
    const bool hasReturnVal)
  {
    _holder = new CodeHolder();
    _holder->init(theCodeInfo);
#if _WIN64
    handRoll64(_holder, (void*)callback, contextData, numArgs, hasReturnVal);
#else
    buildThunk(callback, contextData, numArgs, hasReturnVal, *_holder);
#endif
  }

  ThunkWriter::~ThunkWriter()
  {
    delete _holder;
  }
  size_t ThunkWriter::codeSize() const
  {
    return _holder->codeSize();
  }
  size_t ThunkWriter::writeCode(char* buffer, size_t bufferSize)
  {
    if (!buffer || _holder->codeSize() > bufferSize)
      throw Exception("Cannot write thunk: no buffer or buffer too small");

    size_t codeSize;
    auto err = asmJitWriteCode((uint8_t*)buffer, _holder, codeSize);
    if (err != kErrorOk)
      throw Exception("Thunk write failed: %s", DebugUtils::errorAsString(err));

    return codeSize;
  }


  bool patchThunkData(char* thunk, size_t thunkSize, const void* fromData, const void* toData) noexcept
  {
    if (fromData == toData)
      return true;

    char bufferBefore[10], bufferAfter[10];
    auto bufsize = sizeof(bufferBefore);
    {
      CodeHolder code;
      code.init(theCodeInfo);
      x86::Assembler as(&code);
#ifdef _WIN64
      as.mov(x86::rcx, imm(fromData));
#else
      as.mov(x86::esp, imm(fromData));
#endif
      asmJitWriteCode((uint8_t*)bufferBefore, &code, bufsize);
    }
    {
      CodeHolder code;
      code.init(theCodeInfo);
      x86::Assembler as(&code);
#ifdef _WIN64
      as.mov(x86::rcx, imm(toData));
#else
      as.mov(x86::esp, imm(toData));
#endif
      asmJitWriteCode((uint8_t*)bufferAfter, &code, bufsize);
    }
   
    auto found = std::search(thunk, thunk + thunkSize, bufferBefore, bufferBefore + bufsize);
    if (found == thunk + thunkSize)
      return false;

    memcpy(found, bufferAfter, bufsize);

    return true;
  }
}