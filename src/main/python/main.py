# main.py
import sys
from package.excel_viewer import ExcelViewerApp
from pathlib import Path
from package.resourcesPath import AppContext
from package.IPR import message_box

if __name__ == "__main__":
    ctx = AppContext.get()
    window = ExcelViewerApp()
    message_box.root = window
    window.show()
    window.raise_()         # Place la fenêtre au-dessus des autres
    window.activateWindow() # Active la fenêtre pour qu'elle reçoive le focus
    exit_code = ctx.app.exec()
    sys.exit(exit_code)

