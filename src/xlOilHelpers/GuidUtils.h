#pragma once
#include <string>
#include <xloil/ExportMacro.h>
#include <guiddef.h>

namespace xloil
{
  /// <summary>
  /// Creates a name-based UUID using the algorithm from RFC 4122 §4.3. (for SHA-1 hashing)
  /// </summary>
  /// <param name="namespaceId">The ID of the namespace.</param>
  /// <param name="name">The name (within that namespace).</param>
  /// <returns>A UUID derived from the namespace and name.</returns>
  /// <remarks>See <a href="https://faithlife.codes/blog/2011/04/generating_a_deterministic_guid/">
  /// Generating a deterministic GUID</a>.
  /// </remarks>
  void stableGuidFromString(_GUID& result, const _GUID& id, const std::wstring& path);

  enum class GuidToString
  {
    HEX,
    PUNCTUATED,
    BASE62
  };
  /// <summary>
  /// Wrties the guid in a string of the form 
  /// 
  ///   * PUNCTUATED: '{V-W-X-Y-Z}' length 38 + null terminator
  ///   * HEX : 'VWXYZ' with length 32 + null terminator
  ///   * BASE62: [0-0a-zA-Z]* with length max 23 + null terminator
  ///   
  /// </summary>
  std::wstring guidToWString(const _GUID& guid, GuidToString mode = GuidToString::PUNCTUATED);

  bool createGuid(_GUID& guid);
}