#pragma once
#include <boost/preprocessor/repeat_from_to.hpp>
#include <boost/preprocessor/arithmetic.hpp>
#include <boost/preprocessor/repetition/enum_shifted_params.hpp>
#include <boost/preprocessor/tuple.hpp>

// Wraps boost preprocessor to create some handy macros

#define XLO_DECLARE_ARGS(N, prefix) BOOST_PP_ENUM_SHIFTED_PARAMS(BOOST_PP_ADD(N,1), const ExcelObj& prefix)

#define XLO_ARG_ARRAY_I(z,N,tuple) BOOST_PP_TUPLE_ELEM(2,0,tuple)[N - 1] = &BOOST_PP_TUPLE_ELEM(2,1,tuple)##N;

#define XLO_ARG_ARRAY(N, arrayName, prefix) \
  const ExcelObj* arrayName[N]; \
  BOOST_PP_REPEAT_FROM_TO(1, BOOST_PP_ADD(N, 1), XLO_ARG_ARRAY_I, (arrayName, prefix))

#define XLO_WSTR(x) L ## #x
