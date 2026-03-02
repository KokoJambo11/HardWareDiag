# Runtime hook for PyInstaller: ensure builtins.exit exists (flet.utils.pip may call it).
import builtins
if not hasattr(builtins, 'exit'):
    builtins.exit = lambda code=None: __import__('sys').exit(code)
