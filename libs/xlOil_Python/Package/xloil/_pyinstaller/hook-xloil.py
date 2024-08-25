from importlib.metadata import files
_BINARY_SUFFIXES = [".dll", ".xll", ".pyd"]
binaries=[
    (f.locate(), "share/xloil") 
    for f in files("xloil") 
    if "share" in str(f) and f.suffix in _BINARY_SUFFIXES
]
datas=[
    (f.locate(), "share/xloil") 
    for f in files("xloil") 
    if "share" in str(f) and f.suffix not in _BINARY_SUFFIXES
]
hiddenimports = [
    "xloil.xloil_ribbon", 
    "xloil._event_loop", 
    "xloil.excelfuncs"
]
