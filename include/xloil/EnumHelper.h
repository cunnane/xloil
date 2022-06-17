#pragma once
#include <functional>

namespace xloil
{
  // TODO: obviously having to declare things like this is a bit naff... could improve here
  enum class AddressStyle : int;

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