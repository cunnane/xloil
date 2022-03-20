#pragma once
#include <string>
#include <xloil/ExportMacro.h>
struct _GUID;

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

  std::wstring guidToWString(const _GUID& guid, bool withPunctuation = true);
}