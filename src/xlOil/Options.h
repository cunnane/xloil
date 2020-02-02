#pragma once

// No need to change any of this. We only really need 1 stub function
#define XLOIL_MAX_FUNCS 4
#define XLOIL_STUB_NAME xloil_stub
#define XLOIL_STUB(n) BOOST_PP_CAT(XLOIL_STUB_NAME, n)

constexpr char* XLOIL_SETTINGS_FILE_EXT = "ini";
