#pragma once
#include <functional>

namespace xloil
{
  /// <summary>
  /// Species the format used to write sheet addresses
  /// </summary>
  enum class AddressStyle : int
  {
    /// <summary>
    /// A1 Format: '[Book1]Sheet1'!A1:B2
    /// </summary>
    A1 = 0,
    /// <summary>
    /// RC Format: '[Book1]Sheet1'!R1C1:R2C2
    /// </summary>
    RC = 1,
    /// <summary>
    /// Makes the address absolute, e.g. $A$1
    /// </summary>
    ROW_FIXED = 2,
    /// <summary>
    /// Makes the address absolute, e.g. $A$1
    /// </summary>
    COL_FIXED = 4,
    /// <summary>
    /// Omits the sheet/workbook name
    /// </summary>
    LOCAL = 8,
    /// <summary>
    /// Does not quote sheet name, e.g. [Book1]Sheet1!A1:B2
    /// </summary>
    NOQUOTE = 16,
  };

  namespace detail
  {
    template<class T> struct isFlagEnum        { static constexpr bool value = false; };
    template<> struct isFlagEnum<AddressStyle> { static constexpr bool value = true; };
  }

  // Based on https://stackoverflow.com/questions/49653901/

  template< typename ENUM, typename std::enable_if_t<detail::isFlagEnum<ENUM>::value, int>* = nullptr>
  inline constexpr ENUM operator |(ENUM lhs, ENUM rhs)
  {
    return static_cast<ENUM>(static_cast<std::underlying_type_t<ENUM>>(lhs) | static_cast<std::underlying_type_t<ENUM>>(rhs));
  }

  template< typename ENUM, typename std::enable_if_t<detail::isFlagEnum<ENUM>::value, int>* = nullptr>
  inline constexpr ENUM& operator |=(ENUM& lhs, ENUM rhs)
  {
    lhs = lhs | rhs;
    return lhs;
  }

  template< typename ENUM, typename std::enable_if_t<detail::isFlagEnum<ENUM>::value, int>* = nullptr>
  inline constexpr auto operator &(ENUM lhs, ENUM rhs)
  {
    return static_cast<std::underlying_type_t<ENUM>>(lhs) & static_cast<std::underlying_type_t<ENUM>>(rhs);
  }

  template< typename ENUM, typename std::enable_if_t<detail::isFlagEnum<ENUM>::value, int>* = nullptr>
  inline constexpr ENUM& operator &=(ENUM& lhs, ENUM rhs)
  {
    lhs = lhs & rhs;
    return lhs;
  }

  template< typename ENUM, typename std::enable_if_t<detail::isFlagEnum<ENUM>::value, int>* = nullptr>
  inline constexpr ENUM& operator &=(ENUM& lhs, std::underlying_type_t<ENUM> rhs)
  {
    lhs = static_cast<ENUM>(static_cast<std::underlying_type_t<ENUM>>(lhs) & rhs);
    return lhs;
  }
}