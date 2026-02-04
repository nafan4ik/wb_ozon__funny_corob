import sys
from PySide6.QtWidgets import QApplication

from ui_main import MainWindow


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(1000, 600)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
