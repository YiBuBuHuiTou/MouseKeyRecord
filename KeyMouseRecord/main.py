import record_frame
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QIODevice


if __name__ == '__main__':
    app = QApplication([])
    frame = record_frame.create()
    sys.exit(app.exec_())
