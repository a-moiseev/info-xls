import sys

from PySide6.QtCore import Slot
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from calc import CalcZP
from config import Config


class MainWindow(QMainWindow):
    def __init__(self, config: Config):
        super().__init__()

        self.config = config

        self.setWindowTitle("Расчет ЗП")
        self.setGeometry(100, 100, 500, 300)

        self.zp_label = QLabel(f"Файл расчета ЗП:\n{config.info_path}")
        self.from_label = QLabel(f"Папка с файлами:\n{config.from_files_path}")
        self.to_label = QLabel(f"Папка для результатов:\n{config.to_files_path}")
        self.zp_label.setWordWrap(True)

        self.select_zp_button = QPushButton("Выбрать")
        self.select_from_button = QPushButton("Выбрать")
        self.select_to_button = QPushButton("Выбрать")
        self.start_button = QPushButton("Старт")
        self.start_button.clicked.connect(self.on_start_clicked)

        self.select_zp_button.clicked.connect(self.on_select_file)
        self.select_from_button.clicked.connect(self.on_select_from_dir)
        self.select_to_button.clicked.connect(self.on_select_to_dir)

        main_layout = QVBoxLayout()

        pairs = [
            (self.zp_label, self.select_zp_button),
            (self.from_label, self.select_from_button),
            (self.to_label, self.select_to_button),
        ]

        for label, button in pairs:
            h_layout = QHBoxLayout()
            h_layout.addWidget(label, stretch=1)
            h_layout.addWidget(button)
            main_layout.addLayout(h_layout)

        # Добавляем прогресс-бар
        self.progress_bar = QProgressBar()
        main_layout.addWidget(self.progress_bar)
        
        main_layout.addWidget(self.start_button)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        self.calc_zp = CalcZP(self.config)

    @Slot()
    def on_select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл расчета ЗП",
            str(self.config.current_path),
            "Excel files (*.xlsx)",
        )
        if file_path:
            # Обновляем конфиг
            self.config.update_param("info_path", file_path)
            self.zp_label.setText(f"Файл расчета ЗП:\n{self.config.info_path}")

    @Slot()
    def on_select_from_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "Выберите папку с файлами", str(self.config.current_path)
        )
        if dir_path:
            self.config.update_param("files_path", dir_path)
            self.from_label.setText(f"Папка с файлами:\n{self.config.from_files_path}")

    @Slot()
    def on_select_to_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "Выберите папку для результатов", str(self.config.current_path)
        )
        if dir_path:
            self.config.update_param("to_files_path", dir_path)
            self.to_label.setText(
                f"Папка для результатов:\n{self.config.to_files_path}"
            )

    @Slot()
    def on_start_clicked(self):
        # Настраиваем прогресс-бар
        files_count = len(self.calc_zp.get_files_df())
        # Учитываем все этапы: для каждого файла (копирование + конвертация + обработка + сохранение) + итоговый файл
        total_steps = files_count * 4 + 1  # 4 этапа на файл + 1 этап сохранения итогового файла
        self.progress_bar.setMaximum(total_steps)
        self.progress_bar.setValue(0)
        
        # Передаем callback для обновления прогресса
        self.calc_zp.calculate(progress_callback=self.update_progress)
        
    def update_progress(self, current_value):
        """Обновляет прогресс-бар."""
        self.progress_bar.setValue(current_value)
        QApplication.processEvents()  # Обновляем GUI


if __name__ == "__main__":
    conf = Config()

    app = QApplication(sys.argv)
    window = MainWindow(config=conf)
    window.show()
    sys.exit(app.exec())
