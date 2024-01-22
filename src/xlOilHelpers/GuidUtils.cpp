#include "GuidUtils.h"
#include "Exception.h"

#include <xloil/WindowsSlim.h>
#include <xloil/StringUtils.h>
#include <string>
#include <vector>
#include <memory>
#include <combaseapi.h> // StringFromGUID2
#include <bcrypt.h>
#include <wincrypt.h>

#pragma comment(lib, "bcrypt.lib")
//#pragma comment(lib, "crypt32.lib")

using std::string;
using std::wstring;
using std::make_shared;
using xloil::Helpers::Exception;

#define STATUS_UNSUCCESSFUL ((NTSTATUS)0xC0000001L)

namespace
{
  class HashAlgorithm
  {
    BCRYPT_ALG_HANDLE       hAlg = NULL;
    BCRYPT_HASH_HANDLE      hHash = NULL;
    PBYTE                   pbHashObject = NULL;
    PBYTE                   pbHash = NULL;
    DWORD                   cbHash;

  public:
    HashAlgorithm(const wchar_t* algorithm)
    {
      NTSTATUS status = STATUS_UNSUCCESSFUL;
      DWORD cbData = 0, cbHashObject = 0;

      //open an algorithm handle
      if (!BCRYPT_SUCCESS(status = BCryptOpenAlgorithmProvider(
        &hAlg,
        algorithm,
        NULL,
        0)))
      {
        throw Exception(L"Error 0x{x} returned by BCryptOpenAlgorithmProvider", status);
      }

      //calculate the size of the buffer to hold the hash object
      if (!BCRYPT_SUCCESS(status = BCryptGetProperty(
        hAlg,
        BCRYPT_OBJECT_LENGTH,
        (PBYTE)&cbHashObject,
        sizeof(DWORD),
        &cbData,
        0)))
      {
        throw Exception(L"Error 0x{x} returned by BCryptGetProperty", status);
      }

      //allocate the hash object on the heap
      pbHashObject = new BYTE[cbHashObject];

      //calculate the length of the hash
      if (!BCRYPT_SUCCESS(status = BCryptGetProperty(
        hAlg,
        BCRYPT_HASH_LENGTH,
        (PBYTE)&cbHash,
        sizeof(DWORD),
        &cbData,
        0)))
      {
        throw Exception(L"Error 0x{x} returned by BCryptGetProperty", status);
      }

      //allocate the hash buffer on the heap
      pbHash = new BYTE[cbHash];

      //create a hash
      if (!BCRYPT_SUCCESS(status = BCryptCreateHash(
        hAlg,
        &hHash,
        pbHashObject,
        cbHashObject,
        NULL,
        0,
        0)))
      {
        throw Exception(L"Error 0x{x} returned by BCryptCreateHash", status);
      }
    }

    void transformBlock(BYTE* data, DWORD dataBytes)
    {
      NTSTATUS status = STATUS_UNSUCCESSFUL;

      //hash some data
      if (!BCRYPT_SUCCESS(status = BCryptHashData(
        hHash,
        data,
        dataBytes,
        0)))
      {
        throw Exception(L"Error 0x{x} returned by BCryptHashData", status);
      }
    }

    auto computeHash()
    {
      NTSTATUS status = STATUS_UNSUCCESSFUL;

      //close the hash
      if (!BCRYPT_SUCCESS(status = BCryptFinishHash(
        hHash,
        pbHash,
        cbHash,
        0)))
      {
        throw Exception(L"Error 0x{x} returned by BCryptFinishHash", status);
      }

      return std::vector<BYTE>(pbHash, pbHash + cbHash);
    }

    ~HashAlgorithm()
    {

      if (hAlg)
        BCryptCloseAlgorithmProvider(hAlg, 0);

      if (hHash)
        BCryptDestroyHash(hHash);

      if (pbHashObject)
        delete[] pbHashObject;

      if (pbHash)
        delete[] pbHash;
    }
  };

  void SwapBytes(BYTE* guid, int left, int right)
  {
    auto temp = guid[left];
    guid[left] = guid[right];
    guid[right] = temp;
  }

  // Converts a GUID (expressed as a byte array) to/from network order (MSB-first).
  void SwapByteOrder(BYTE* guid)
  {
    SwapBytes(guid, 0, 3);
    SwapBytes(guid, 1, 2);
    SwapBytes(guid, 4, 5);
    SwapBytes(guid, 6, 7);
  }


  /// <summary>
  /// Creates a name-based UUID using the algorithm from RFC 4122 §4.3. (for SHA-1 hashing)
  /// </summary>
  /// <param name="namespaceId">The ID of the namespace.</param>
  /// <param name="name">The name (within that namespace).</param>
  /// <returns>A UUID derived from the namespace and name.</returns>
  /// <remarks>See <a href="https://faithlife.codes/blog/2011/04/generating_a_deterministic_guid/">Generating a deterministic GUID</a>.</remarks>
  void Create(GUID& result, const GUID& namespaceId, const string& name)
  {
    static_assert(sizeof(GUID) == 16);

    // convert the name to a sequence of octets (as defined by the standard or conventions of its namespace) (step 3)
    // ASSUME: UTF-8 encoding is always appropriate
    auto nameBytes = (BYTE*)name.data();

    // convert the namespace UUID to network order (step 3)
    BYTE namespaceBytes[sizeof(GUID)];
    memcpy(namespaceBytes, (BYTE*)&namespaceId, sizeof(GUID));
    SwapByteOrder(namespaceBytes);

    // compute the hash of the name space ID concatenated with the name (step 4)
    HashAlgorithm algorithm(BCRYPT_SHA1_ALGORITHM);
    algorithm.transformBlock(namespaceBytes, sizeof(GUID));
    algorithm.transformBlock(nameBytes, (DWORD)name.size());
    auto hash = algorithm.computeHash();

    // most bytes from the hash are copied straight to the bytes of the new GUID (steps 5-7, 9, 11-12)
    auto newGuid = (BYTE*)&result;
    memcpy_s(newGuid, sizeof(GUID), hash.data(), sizeof(GUID));

    const auto version = 5;

    // set the four most significant bits (bits 12 through 15) of the time_hi_and_version field to the appropriate 4-bit version number from Section 4.1.3 (step 8)
    newGuid[6] = (BYTE)((newGuid[6] & 0x0F) | (version << 4));

    // set the two most significant bits (bits 6 and 7) of the clock_seq_hi_and_reserved to zero and one, respectively (step 10)
    newGuid[8] = (BYTE)((newGuid[8] & 0x3F) | 0x80);

    // convert the resulting UUID to local byte order (step 13)
    SwapByteOrder(newGuid);
  }

  const GUID theExcelDnaNamespaceGuid =
    { 0x306D016E, 0xCCE8, 0x4861, { 0x9D, 0xA1, 0x51, 0xA2, 0x7C, 0xBE, 0x34, 0x1A} };
}
namespace xloil
{
  // Return a stable Guid from the xll path - used for COM registration and helper functions
  // Uses the .ToUpperInvariant() of the path name.
  void stableGuidFromString(GUID& result, const GUID& id, const std::wstring& path)
  {
    wstring upper;
    upper.resize(path.size());
    LCMapStringW(LOCALE_INVARIANT, LCMAP_UPPERCASE, path.data(), (int)path.size(), upper.data(), (int)upper.size());
    const string utf8 = xloil::utf16ToUtf8(upper);
    return Create(result, id, utf8);
  }

  std::wstring guidToWString(const GUID& guid, GuidToString mode)
  {
    wchar_t result[39]; // 32 hex chars + 4 hyphens + two braces + null terminator
    int ret = 0;

    switch (mode)
    {
    case GuidToString::PUNCTUATED:
      ret = StringFromGUID2(guid, result, _countof(result));
      break;
    case GuidToString::HEX:
      ret = _snwprintf_s(
        result, sizeof(result),
        L"%08x%04x%04x%02x%02x%02x%02x%02x%02x%02x%02x",
        guid.Data1,    guid.Data2,    guid.Data3,
        guid.Data4[0], guid.Data4[1], guid.Data4[2],
        guid.Data4[3], guid.Data4[4], guid.Data4[5],
        guid.Data4[6], guid.Data4[7]);
      break;
    case GuidToString::BASE62:
    {
      auto written = 0u;
      auto uintPtr = (const size_t*)&guid;
      for (auto i = 0; i < sizeof(guid) / sizeof(*uintPtr); ++i)
        written += unsignedToString<62>(uintPtr[i], result + written, _countof(result) - written);

#ifdef _WIN64
      static_assert(sizeof(guid) == 2 * sizeof(*uintPtr));
#else
      static_assert(sizeof(guid) == 4 * sizeof(*uintPtr));
#endif

      result[written] = 0; // null terminator
      ret = written > 0;
      break;
    }
    }

    if (ret <= 0)
      throw Exception("Failed to write GUID as string");

    return result;
  }

  bool createGuid(_GUID& guid)
  {
    return CoCreateGuid(&guid) == 0;
  }
}
