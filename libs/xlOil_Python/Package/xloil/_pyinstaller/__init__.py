import os

def get_hook_dirs():
    # Tell PyInstaller where to find hooks provided by this distribution
    return [os.path.dirname(__file__)]
