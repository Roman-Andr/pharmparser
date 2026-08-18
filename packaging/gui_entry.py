"""PyInstaller entry point for the GUI.

PyInstaller runs its entry script as ``__main__``, so pointing it straight at
``src/pharmparser/__main__.py`` breaks every relative import in the package. This
launcher imports ``pharmparser`` as a package instead.
"""

from multiprocessing import freeze_support

from pharmparser.__main__ import main

if __name__ == "__main__":
    freeze_support()
    raise SystemExit(main())
