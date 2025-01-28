#pragma once
#include <list>


namespace xloil
{
  /// <summary>
  /// Move elements from *from* list to *to* list if they match the predicate 
  /// </summary>
  template<class T, class Pred>
  void move_if(std::list<T>& from, std::list<T>& to, Pred predicate)
  {
    // Note the postfix i++ as the iterator passed to splice will be invalidated.
    for (auto i = from.begin(); i != from.end(); ++i)
      if (predicate(*i))
        to.splice(to.end(), from, i++);
  }
}
