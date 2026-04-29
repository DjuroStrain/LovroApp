import sys
import os
import cv2
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QLabel, QMessageBox
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QImage, QPixmap
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


CAPTURES_DIR = "captures"
MAX_PHOTOS = 4


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Photo Capture")
        self.setMinimumSize(800, 620)

        self.captured_paths = []
        self.doc = Document()
        self.camera = None
        self.timer = QTimer()

        self._init_ui()
        self._init_camera()

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        self.preview_label = QLabel("Initializing camera...")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumHeight(480)
        self.preview_label.setStyleSheet("background: #111; color: #aaa; font-size: 14px;")
        layout.addWidget(self.preview_label)

        controls = QHBoxLayout()

        self.counter_label = QLabel(f"Photos: 0 / {MAX_PHOTOS}")
        self.counter_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        controls.addWidget(self.counter_label)

        controls.addStretch()

        self.capture_btn = QPushButton("Capture Photo")
        self.capture_btn.setFixedHeight(40)
        self.capture_btn.setStyleSheet("font-size: 14px; padding: 0 20px;")
        self.capture_btn.clicked.connect(self.capture_photo)
        controls.addWidget(self.capture_btn)

        self.save_btn = QPushButton("Save Document")
        self.save_btn.setFixedHeight(40)
        self.save_btn.setStyleSheet("font-size: 14px; padding: 0 20px;")
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self.save_document)
        controls.addWidget(self.save_btn)

        layout.addLayout(controls)

    def _init_camera(self):
        self.camera = cv2.VideoCapture(0)
        if not self.camera.isOpened():
            QMessageBox.critical(self, "Camera Error", "No webcam detected. Please connect a USB camera and restart the app.")
            self.capture_btn.setEnabled(False)
            return

        self.timer.timeout.connect(self._update_frame)
        self.timer.start(33)

    def _update_frame(self):
        if self.camera is None or not self.camera.isOpened():
            return
        ret, frame = self.camera.read()
        if not ret:
            return
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = frame_rgb.shape
        img = QImage(frame_rgb.data, w, h, ch * w, QImage.Format.Format_RGB888)
        pixmap = QPixmap.fromImage(img).scaled(
            self.preview_label.width(),
            self.preview_label.height(),
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self.preview_label.setPixmap(pixmap)

    def capture_photo(self):
        if self.camera is None or not self.camera.isOpened():
            QMessageBox.warning(self, "Camera Error", "Camera is not available.")
            return

        ret, frame = self.camera.read()
        if not ret:
            QMessageBox.warning(self, "Capture Failed", "Failed to capture image from camera.")
            return

        os.makedirs(CAPTURES_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = os.path.join(CAPTURES_DIR, f"photo_{timestamp}.jpg")
        cv2.imwrite(filename, frame)
        self.captured_paths.append(filename)

        count = len(self.captured_paths)
        self.counter_label.setText(f"Photos: {count} / {MAX_PHOTOS}")

        if count >= MAX_PHOTOS:
            self.capture_btn.setEnabled(False)
            self.save_btn.setEnabled(True)

    def save_document(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=len(self.captured_paths))

        for i, path in enumerate(self.captured_paths):
            cell = table.cell(0, i)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            try:
                run.add_picture(path, width=Inches(4))
            except Exception as e:
                QMessageBox.critical(self, "Document Error", f"Failed to embed image {i + 1}:\n{e}")
                return

        docs_folder = os.path.expanduser("~/Documents")
        os.makedirs(docs_folder, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        out_path = os.path.join(docs_folder, f"captures_{timestamp}.docx")

        try:
            doc.save(out_path)
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Could not save document:\n{e}")
            return

        QMessageBox.information(self, "Saved", f"Document saved to:\n{out_path}")

        self.captured_paths = []
        self.counter_label.setText(f"Photos: 0 / {MAX_PHOTOS}")
        self.capture_btn.setEnabled(True)
        self.save_btn.setEnabled(False)

    def closeEvent(self, event):
        self.timer.stop()
        if self.camera and self.camera.isOpened():
            self.camera.release()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()