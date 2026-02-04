from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton
from PySide6.QtGui import QMovie
from PySide6.QtMultimedia import QSoundEffect
from PySide6.QtCore import Qt, QUrl
from pathlib import Path

ASSETS_DIR = Path(__file__).resolve().parent / "assets"


class GoblinDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Ну ё-моё…")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.setFixedSize(360, 260)

        layout = QVBoxLayout(self)

        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)

        self.movie = QMovie(str(ASSETS_DIR / "goblin.gif"))
        self.label.setMovie(self.movie)
        layout.addWidget(self.label)

        btn = QPushButton("Выключить 😭")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)

        self.sound = QSoundEffect(self)
        self.sound.setSource(
            QUrl.fromLocalFile(str(ASSETS_DIR / "goblin.wav"))
        )
        self.sound.setVolume(0.9)

    def showEvent(self, event):
        self.movie.start()
        self.sound.play()
        super().showEvent(event)

    def closeEvent(self, event):
        self.movie.stop()
        self.sound.stop()
        super().closeEvent(event)
