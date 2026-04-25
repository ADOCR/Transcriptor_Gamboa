# -*- coding: utf-8 -*-
"""
Aplicacion de escritorio para Windows:
Gamboa Desarrollos - Transcriptor Local y Minutador.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import Qt, QThread, QTimer, Signal, QUrl
from PySide6.QtGui import QDesktopServices, QIcon, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app_core import (
    APP_NAME,
    APP_SHORT_NAME,
    AppConfig,
    GamboaProcessor,
    OutputOptions,
    ProcessCallbacks,
    ProcessSummary,
    abrir_carpeta,
    exportar_reporte_ejecucion,
    verificar_cuda,
    verificar_ffmpeg,
    verificar_ollama,
    verificar_whisper_import,
)


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def format_seconds(seconds: float) -> str:
    seconds = max(0, int(seconds))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


class ProcessWorker(QThread):
    log = Signal(str)
    status = Signal(str)
    general_progress = Signal(int, int)
    file_progress = Signal(int)
    current_file = Signal(str)
    error = Signal(str)
    finished_summary = Signal(object)
    fatal_error = Signal(str)

    def __init__(self, mode: str, inputs: List[Path], config: AppConfig):
        super().__init__()
        self.mode = mode
        self.inputs = inputs
        self.config = config
        self._stop_after_current = False
        self.processor: Optional[GamboaProcessor] = None

    def stop_after_current(self) -> None:
        self._stop_after_current = True

    def run(self) -> None:
        callbacks = ProcessCallbacks(
            log=self.log.emit,
            status=self.status.emit,
            general_progress=self.general_progress.emit,
            file_progress=self.file_progress.emit,
            current_file=self.current_file.emit,
            error=self.error.emit,
            should_stop_after_current=lambda: self._stop_after_current,
        )
        try:
            self.processor = GamboaProcessor(self.config, callbacks)
            summary = self.processor.process(self.mode, self.inputs)
            self.finished_summary.emit(summary)
        except Exception as exc:
            self.fatal_error.emit(str(exc))


class SettingsDialog(QDialog):
    def __init__(self, parent: "MainWindow"):
        super().__init__(parent)
        self.setWindowTitle("Configuracion avanzada")
        self.setMinimumWidth(520)
        self.parent_window = parent

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.cpu_threads = QSpinBox()
        self.cpu_threads.setRange(1, 32)
        self.cpu_threads.setValue(parent.cpu_threads)

        self.beam_size = QSpinBox()
        self.beam_size.setRange(1, 10)
        self.beam_size.setValue(parent.beam_size)

        self.ollama_timeout = QSpinBox()
        self.ollama_timeout.setRange(60, 7200)
        self.ollama_timeout.setValue(parent.ollama_timeout)
        self.ollama_timeout.setSuffix(" s")

        self.ollama_url = QLineEdit(parent.ollama_url)
        self.ollama_tags_url = QLineEdit(parent.ollama_tags_url)
        self.vad_filter = QCheckBox("Activar VAD filter")
        self.vad_filter.setChecked(parent.vad_filter)
        self.cuda_agresiva = QCheckBox("Permitir limpieza CUDA agresiva")
        self.cuda_agresiva.setChecked(parent.limpieza_cuda_agresiva)

        form.addRow("Hilos CPU", self.cpu_threads)
        form.addRow("Beam size", self.beam_size)
        form.addRow("Timeout Ollama", self.ollama_timeout)
        form.addRow("URL generate", self.ollama_url)
        form.addRow("URL tags", self.ollama_tags_url)
        form.addRow("", self.vad_filter)
        form.addRow("", self.cuda_agresiva)

        note = QLabel(
            "La limpieza CUDA agresiva puede ayudar a liberar VRAM, pero en algunos equipos "
            "puede cerrar la aplicacion. Mantengala desactivada salvo que sea necesario."
        )
        note.setWordWrap(True)
        note.setObjectName("MutedLabel")

        buttons = QHBoxLayout()
        save = QPushButton("Guardar")
        cancel = QPushButton("Cancelar")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        buttons.addStretch()
        buttons.addWidget(cancel)
        buttons.addWidget(save)

        layout.addLayout(form)
        layout.addWidget(note)
        layout.addLayout(buttons)

    def accept(self) -> None:
        parent = self.parent_window
        parent.cpu_threads = self.cpu_threads.value()
        parent.beam_size = self.beam_size.value()
        parent.ollama_timeout = self.ollama_timeout.value()
        parent.ollama_url = self.ollama_url.text().strip()
        parent.ollama_tags_url = self.ollama_tags_url.text().strip()
        parent.vad_filter = self.vad_filter.isChecked()
        parent.limpieza_cuda_agresiva = self.cuda_agresiva.isChecked()
        super().accept()


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ayuda - Gamboa Transcriptor")
        self.resize(760, 620)

        layout = QVBoxLayout(self)
        text = QTextEdit()
        text.setReadOnly(True)
        text.setHtml(
            """
            <h1>Gamboa Transcriptor</h1>
            <h2>Ollama</h2>
            <p>Para generar minutas, instale Ollama desde <b>https://ollama.com</b>.</p>
            <p>Luego abra una terminal y ejecute:</p>
            <pre>ollama pull qwen3:8b</pre>

            <h2>Recomendaciones RTX 3070 8 GB</h2>
            <ul>
              <li><b>large-v3</b>: maxima calidad, mayor consumo de VRAM.</li>
              <li><b>medium</b>: mas rapido y consume menos VRAM.</li>
              <li><b>num_ctx=4096</b>: opcion mas estable para generar minutas.</li>
              <li>Para muchos archivos: use primero <b>Solo transcripcion</b> y luego <b>Minuta desde TXT</b>.</li>
            </ul>

            <h2>Archivos generados</h2>
            <ul>
              <li><b>TXT con tiempos</b>: texto con marcas [HH:MM:SS - HH:MM:SS].</li>
              <li><b>TXT limpio</b>: transcripcion corrida sin marcas.</li>
              <li><b>SRT</b>: subtitulos compatibles con reproductores y editores.</li>
              <li><b>DOCX</b>: transcripcion o minuta en Word.</li>
              <li><b>MINUTA</b>: resumen estructurado generado por Ollama local.</li>
            </ul>

            <h2>Privacidad</h2>
            <p>La transcripcion y la generacion de minutas se realizan localmente en el equipo del usuario.</p>
            """
        )
        close = QPushButton("Cerrar")
        close.clicked.connect(self.accept)
        layout.addWidget(text)
        layout.addWidget(close, alignment=Qt.AlignRight)


class ResultDialog(QDialog):
    def __init__(self, summary: ProcessSummary, parent=None):
        super().__init__(parent)
        self.summary = summary
        self.setWindowTitle("Resumen de ejecucion")
        self.resize(620, 420)

        layout = QVBoxLayout(self)
        title = QLabel("Proceso terminado")
        title.setObjectName("DialogTitle")

        details = QLabel(
            f"Archivos correctos: {summary.ok_count}\n"
            f"Archivos con error: {summary.error_count}\n"
            f"Minutas generadas: {summary.minute_count}\n"
            f"Tiempo total: {format_seconds(summary.elapsed_seconds)}\n"
            f"Carpeta de salida: {summary.output_root or 'No especificada'}"
        )
        details.setTextInteractionFlags(Qt.TextSelectableByMouse)

        buttons = QHBoxLayout()
        open_btn = QPushButton("Abrir carpeta")
        close_btn = QPushButton("Cerrar")
        open_btn.clicked.connect(self.open_folder)
        close_btn.clicked.connect(self.accept)
        buttons.addStretch()
        buttons.addWidget(open_btn)
        buttons.addWidget(close_btn)

        layout.addWidget(title)
        layout.addWidget(details)
        layout.addStretch()
        layout.addLayout(buttons)

    def open_folder(self) -> None:
        if self.summary.output_root:
            abrir_carpeta(self.summary.output_root)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        icon = resource_path("assets/icon.ico")
        if icon.exists():
            self.setWindowIcon(QIcon(str(icon)))

        self.resize(1220, 820)
        self.selected_inputs: List[Path] = []
        self.input_kind: Optional[str] = None
        self.output_auto = True
        self.current_summary: Optional[ProcessSummary] = None
        self.worker: Optional[ProcessWorker] = None
        self.started_at: Optional[float] = None
        self.progress_done = 0
        self.progress_total = 0

        self.cpu_threads = 6
        self.beam_size = 5
        self.ollama_timeout = 900
        self.ollama_url = "http://localhost:11434/api/generate"
        self.ollama_tags_url = "http://localhost:11434/api/tags"
        self.vad_filter = True
        self.limpieza_cuda_agresiva = False

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time_labels)

        self.build_ui()
        self.apply_style()

    # ------------------------------------------------------------------ UI
    def build_ui(self) -> None:
        central = QWidget()
        main = QVBoxLayout(central)
        main.setContentsMargins(18, 18, 18, 18)
        main.setSpacing(14)
        main.addWidget(self.build_header())

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        grid = QGridLayout(content)
        grid.setSpacing(14)

        left = QVBoxLayout()
        right = QVBoxLayout()
        left.addWidget(self.build_mode_section())
        left.addWidget(self.build_input_section())
        left.addWidget(self.build_output_section())
        left.addWidget(self.build_config_section())
        left.addStretch()

        right.addWidget(self.build_system_section())
        right.addWidget(self.build_execution_section())
        right.addWidget(self.build_log_section(), stretch=1)

        grid.addLayout(left, 0, 0)
        grid.addLayout(right, 0, 1)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        scroll.setWidget(content)
        main.addWidget(scroll, stretch=1)
        self.setCentralWidget(central)

    def build_header(self) -> QWidget:
        header = QFrame()
        header.setObjectName("Header")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(18, 16, 18, 16)

        logo = QLabel()
        logo.setFixedSize(72, 72)
        logo.setObjectName("LogoBox")
        logo_path = resource_path("assets/logo.png")
        if logo_path.exists():
            pix = QPixmap(str(logo_path))
            logo.setPixmap(pix.scaled(72, 72, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            logo.setText("GD")
            logo.setAlignment(Qt.AlignCenter)

        title_box = QVBoxLayout()
        company = QLabel("Gamboa Desarrollos")
        company.setObjectName("CompanyTitle")
        title = QLabel("Transcriptor Local y Generador de Minutas")
        title.setObjectName("MainTitle")
        subtitle = QLabel("Transcripcion y minutas locales para audio y video")
        subtitle.setObjectName("Subtitle")
        title_box.addWidget(company)
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        settings = QPushButton("Configuracion")
        settings.setObjectName("SecondaryButton")
        settings.clicked.connect(self.open_settings)
        help_btn = QPushButton("Ayuda")
        help_btn.setObjectName("SecondaryButton")
        help_btn.clicked.connect(self.open_help)

        layout.addWidget(logo)
        layout.addLayout(title_box)
        layout.addStretch()
        layout.addWidget(help_btn)
        layout.addWidget(settings)
        return header

    def card(self, title: str) -> QGroupBox:
        box = QGroupBox(title)
        box.setObjectName("Card")
        return box

    def build_mode_section(self) -> QWidget:
        box = self.card("A. Modo de proceso")
        layout = QVBoxLayout(box)
        self.mode_group = QButtonGroup(self)

        self.mode_transcribe = QRadioButton("Solo transcripcion")
        self.mode_both = QRadioButton("Transcripcion + resumen/minuta")
        self.mode_txt = QRadioButton("Generar resumen/minuta desde TXT existente")
        self.mode_transcribe.setChecked(True)

        self.mode_group.addButton(self.mode_transcribe)
        self.mode_group.addButton(self.mode_both)
        self.mode_group.addButton(self.mode_txt)
        self.mode_group.buttonClicked.connect(lambda _button: self.update_mode_hint())

        layout.addWidget(self.mode_transcribe)
        layout.addWidget(self.description("Genera TXT, SRT y DOCX. No requiere Ollama."))
        layout.addWidget(self.mode_both)
        layout.addWidget(self.description("Usa Whisper y luego Ollama local para crear la minuta."))
        layout.addWidget(self.mode_txt)
        layout.addWidget(self.description("Recomendado si se quiere separar transcripcion y minuta para cuidar VRAM."))
        self.mode_hint = QLabel("")
        self.mode_hint.setObjectName("WarningLabel")
        self.mode_hint.setWordWrap(True)
        layout.addWidget(self.mode_hint)
        return box

    def build_input_section(self) -> QWidget:
        box = self.card("B. Entrada")
        layout = QVBoxLayout(box)

        buttons = QHBoxLayout()
        file_btn = QPushButton("Seleccionar archivo")
        folder_btn = QPushButton("Seleccionar carpeta")
        txt_btn = QPushButton("Seleccionar TXT para minuta")
        file_btn.clicked.connect(self.select_file)
        folder_btn.clicked.connect(self.select_folder)
        txt_btn.clicked.connect(self.select_txt)
        buttons.addWidget(file_btn)
        buttons.addWidget(folder_btn)
        buttons.addWidget(txt_btn)

        self.input_path = QLineEdit()
        self.input_path.setReadOnly(True)
        self.input_path.setPlaceholderText("Seleccione un archivo, carpeta o TXT...")

        layout.addLayout(buttons)
        layout.addWidget(self.input_path)
        return box

    def build_output_section(self) -> QWidget:
        box = self.card("C. Carpeta de salida")
        layout = QVBoxLayout(box)
        self.output_path = QLineEdit()
        self.output_path.setReadOnly(True)
        self.output_path.setPlaceholderText("Carpeta automatica junto al archivo origen")

        buttons = QHBoxLayout()
        auto_btn = QPushButton("Usar carpeta automatica")
        change_btn = QPushButton("Cambiar carpeta de salida")
        auto_btn.clicked.connect(self.use_auto_output)
        change_btn.clicked.connect(self.change_output_folder)
        buttons.addWidget(auto_btn)
        buttons.addWidget(change_btn)

        layout.addWidget(self.output_path)
        layout.addLayout(buttons)
        return box

    def build_config_section(self) -> QWidget:
        box = self.card("D. Configuracion rapida")
        layout = QGridLayout(box)

        self.model_combo = QComboBox()
        self.model_combo.addItems(["large-v3", "medium", "small"])
        self.model_combo.currentTextChanged.connect(lambda _text: self.update_mode_hint())

        self.device_combo = QComboBox()
        self.device_combo.addItems(["Auto", "CUDA", "CPU"])

        self.context_combo = QComboBox()
        self.context_combo.addItems(["4096", "8192"])
        self.fragment_combo = QComboBox()
        self.fragment_combo.addItems(["8000", "12000"])
        self.ollama_model = QLineEdit("qwen3:8b")

        self.out_times = QCheckBox("TXT con tiempos")
        self.out_clean = QCheckBox("TXT limpio")
        self.out_srt = QCheckBox("SRT")
        self.out_docx = QCheckBox("DOCX transcripcion")
        self.out_minute_docx = QCheckBox("DOCX minuta")
        for cb in [self.out_times, self.out_clean, self.out_srt, self.out_docx, self.out_minute_docx]:
            cb.setChecked(True)

        self.overwrite_outputs = QCheckBox("Sobrescribir salidas existentes")
        self.release_whisper = QCheckBox("Liberar modelo Whisper antes de generar minutas")

        layout.addWidget(QLabel("Modelo Whisper"), 0, 0)
        layout.addWidget(self.model_combo, 0, 1)
        layout.addWidget(QLabel("Dispositivo"), 0, 2)
        layout.addWidget(self.device_combo, 0, 3)
        layout.addWidget(QLabel("Ollama modelo"), 1, 0)
        layout.addWidget(self.ollama_model, 1, 1)
        layout.addWidget(QLabel("Contexto"), 1, 2)
        layout.addWidget(self.context_combo, 1, 3)
        layout.addWidget(QLabel("Fragmento minuta"), 2, 0)
        layout.addWidget(self.fragment_combo, 2, 1)

        checks = QGridLayout()
        checks.addWidget(self.out_times, 0, 0)
        checks.addWidget(self.out_clean, 0, 1)
        checks.addWidget(self.out_srt, 1, 0)
        checks.addWidget(self.out_docx, 1, 1)
        checks.addWidget(self.out_minute_docx, 2, 0)
        checks.addWidget(self.overwrite_outputs, 2, 1)
        checks.addWidget(self.release_whisper, 3, 0, 1, 2)
        layout.addLayout(checks, 3, 0, 1, 4)
        return box

    def build_system_section(self) -> QWidget:
        box = self.card("E. Verificaciones del sistema")
        layout = QVBoxLayout(box)

        self.cuda_status = self.status_card("CUDA", "Sin verificar", "status-blue")
        self.ollama_status = self.status_card("Ollama", "Sin verificar", "status-blue")
        self.ffmpeg_status = self.status_card("FFmpeg/PyAV", "Sin verificar", "status-blue")
        layout.addWidget(self.cuda_status)
        layout.addWidget(self.ollama_status)
        layout.addWidget(self.ffmpeg_status)

        buttons = QGridLayout()
        cuda_btn = QPushButton("Verificar CUDA")
        ollama_btn = QPushButton("Verificar Ollama")
        open_ollama_btn = QPushButton("Abrir Ollama")
        guide_btn = QPushButton("Ver guia de instalacion")
        cuda_btn.clicked.connect(self.check_cuda)
        ollama_btn.clicked.connect(self.check_ollama)
        open_ollama_btn.clicked.connect(self.open_ollama)
        guide_btn.clicked.connect(self.open_help)
        buttons.addWidget(cuda_btn, 0, 0)
        buttons.addWidget(ollama_btn, 0, 1)
        buttons.addWidget(open_ollama_btn, 1, 0)
        buttons.addWidget(guide_btn, 1, 1)
        layout.addLayout(buttons)
        return box

    def build_execution_section(self) -> QWidget:
        box = self.card("F. Ejecucion")
        layout = QVBoxLayout(box)

        self.start_btn = QPushButton("INICIAR PROCESO")
        self.start_btn.setObjectName("PrimaryButton")
        self.stop_btn = QPushButton("DETENER DESPUES DEL ARCHIVO ACTUAL")
        self.open_results_btn = QPushButton("ABRIR CARPETA DE RESULTADOS")
        self.export_report_btn = QPushButton("EXPORTAR REPORTE DE EJECUCION")
        self.stop_btn.setEnabled(False)
        self.open_results_btn.setEnabled(False)
        self.export_report_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_process)
        self.stop_btn.clicked.connect(self.stop_after_current)
        self.open_results_btn.clicked.connect(self.open_results_folder)
        self.export_report_btn.clicked.connect(self.export_report)

        self.general_bar = QProgressBar()
        self.file_bar = QProgressBar()
        self.status_label = QLabel("Listo")
        self.current_file_label = QLabel("Archivo actual: -")
        self.elapsed_label = QLabel("Tiempo transcurrido: 00:00:00")
        self.eta_label = QLabel("Tiempo estimado restante: -")

        layout.addWidget(self.start_btn)
        layout.addWidget(self.stop_btn)
        layout.addWidget(QLabel("Progreso general"))
        layout.addWidget(self.general_bar)
        layout.addWidget(QLabel("Progreso del archivo actual"))
        layout.addWidget(self.file_bar)
        layout.addWidget(self.status_label)
        layout.addWidget(self.current_file_label)
        layout.addWidget(self.elapsed_label)
        layout.addWidget(self.eta_label)
        layout.addWidget(self.open_results_btn)
        layout.addWidget(self.export_report_btn)
        return box

    def build_log_section(self) -> QWidget:
        box = self.card("Registro")
        layout = QVBoxLayout(box)
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMinimumHeight(220)

        buttons = QHBoxLayout()
        copy_btn = QPushButton("Copiar registro")
        save_btn = QPushButton("Guardar registro TXT")
        clear_btn = QPushButton("Limpiar registro")
        copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(self.log_view.toPlainText()))
        save_btn.clicked.connect(self.save_log)
        clear_btn.clicked.connect(self.log_view.clear)
        buttons.addWidget(copy_btn)
        buttons.addWidget(save_btn)
        buttons.addWidget(clear_btn)

        layout.addWidget(self.log_view)
        layout.addLayout(buttons)
        return box

    def description(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("MutedLabel")
        label.setWordWrap(True)
        return label

    def status_card(self, title: str, text: str, status_class: str) -> QLabel:
        label = QLabel(f"{title}: {text}")
        label.setProperty("statusClass", status_class)
        label.setObjectName("StatusCard")
        label.setWordWrap(True)
        return label

    # ------------------------------------------------------------------ Actions
    def select_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar audio o video",
            "",
            "Audio/video (*.mp3 *.wav *.m4a *.aac *.flac *.ogg *.wma *.mp4 *.mkv *.mov *.avi *.webm *.m4v *.wmv);;Todos (*.*)",
        )
        if path:
            self.selected_inputs = [Path(path)]
            self.input_kind = "file"
            self.input_path.setText(path)

    def select_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta")
        if path:
            self.selected_inputs = [Path(path)]
            self.input_kind = "folder"
            self.input_path.setText(path)

    def select_txt(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Seleccionar TXT limpio", "", "TXT (*.txt);;Todos (*.*)")
        if path:
            self.selected_inputs = [Path(path)]
            self.input_kind = "txt"
            self.input_path.setText(path)
            self.mode_txt.setChecked(True)
            self.update_mode_hint()

    def use_auto_output(self) -> None:
        self.output_auto = True
        self.output_path.clear()
        self.output_path.setPlaceholderText("Carpeta automatica junto al archivo origen")

    def change_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de salida")
        if path:
            self.output_auto = False
            self.output_path.setText(path)

    def open_settings(self) -> None:
        SettingsDialog(self).exec()

    def open_help(self) -> None:
        HelpDialog(self).exec()

    def update_mode_hint(self) -> None:
        if self.mode_both.isChecked() and self.model_combo.currentText() == "large-v3":
            self.mode_hint.setText(
                "Advertencia: large-v3 puede ocupar bastante VRAM. Si Ollama queda lento, "
                "use Solo transcripcion y luego Minuta desde TXT."
            )
        else:
            self.mode_hint.setText("")

    def append_log(self, message: str) -> None:
        self.log_view.append(message)

    def set_status_label(self, label: QLabel, title: str, text: str, color: str) -> None:
        label.setText(f"{title}: {text}")
        label.setProperty("statusClass", color)
        label.style().unpolish(label)
        label.style().polish(label)

    def check_cuda(self) -> None:
        whisper = verificar_whisper_import()
        cuda = verificar_cuda()
        color = "status-green" if cuda["available"] else "status-yellow"
        self.set_status_label(self.cuda_status, "CUDA", str(cuda["message"]), color)
        self.append_log(str(whisper["message"]))
        self.append_log(str(cuda["message"]))

    def check_ollama(self) -> None:
        status = verificar_ollama(self.ollama_model.text().strip() or "qwen3:8b", self.ollama_tags_url)
        if status["available"] and status["model_installed"]:
            color = "status-green"
        elif status["available"]:
            color = "status-yellow"
        else:
            color = "status-red"
        self.set_status_label(self.ollama_status, "Ollama", str(status["message"]), color)
        self.append_log(str(status["message"]))

        ff = verificar_ffmpeg()
        self.set_status_label(self.ffmpeg_status, "FFmpeg/PyAV", str(ff["message"]), "status-green" if ff["available"] else "status-yellow")

    def open_ollama(self) -> None:
        exe = shutil.which("ollama")
        if not exe:
            QMessageBox.warning(self, "Ollama no encontrado", "No se encontro ollama.exe en PATH. Instale Ollama desde https://ollama.com.")
            return
        try:
            flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
            subprocess.Popen([exe, "serve"], creationflags=flags)
            self.append_log("Se intento iniciar Ollama local con 'ollama serve'.")
        except Exception as exc:
            QMessageBox.warning(self, "No se pudo abrir Ollama", str(exc))

    def current_mode(self) -> str:
        if self.mode_txt.isChecked():
            return "txt_to_minute"
        if self.mode_both.isChecked():
            return "transcribe_minute"
        return "transcribe"

    def build_config(self) -> AppConfig:
        output_root = Path(self.output_path.text()) if self.output_path.text().strip() else None
        num_ctx = int(self.context_combo.currentText())
        fragment = int(self.fragment_combo.currentText())
        return AppConfig(
            whisper_model_name=self.model_combo.currentText(),
            device_mode=self.device_combo.currentText().lower(),
            beam_size=self.beam_size,
            vad_filter=self.vad_filter,
            cpu_threads=self.cpu_threads,
            ollama_model=self.ollama_model.text().strip() or "qwen3:8b",
            ollama_url=self.ollama_url,
            ollama_tags_url=self.ollama_tags_url,
            ollama_timeout=self.ollama_timeout,
            ollama_options={"temperature": 0.1, "num_ctx": num_ctx},
            tamano_fragmento_minuta=fragment,
            output_root=output_root,
            usar_carpeta_automatica=self.output_auto,
            sobrescribir_salidas=self.overwrite_outputs.isChecked(),
            liberar_whisper_antes_ollama=self.release_whisper.isChecked(),
            limpieza_cuda_agresiva=self.limpieza_cuda_agresiva,
            output_options=OutputOptions(
                txt_con_tiempos=self.out_times.isChecked(),
                txt_limpio=self.out_clean.isChecked(),
                srt=self.out_srt.isChecked(),
                docx_transcripcion=self.out_docx.isChecked(),
                docx_minuta=self.out_minute_docx.isChecked(),
            ),
        )

    def validate_inputs(self) -> bool:
        if not self.selected_inputs:
            QMessageBox.warning(self, "Entrada requerida", "Seleccione un archivo, carpeta o TXT.")
            return False
        if self.current_mode() == "txt_to_minute":
            if not all(p.suffix.lower() == ".txt" for p in self.selected_inputs):
                QMessageBox.warning(self, "TXT requerido", "Para este modo seleccione un archivo TXT.")
                return False
        else:
            if self.input_kind == "txt":
                QMessageBox.warning(self, "Entrada incompatible", "Para transcribir seleccione audio/video o una carpeta.")
                return False
        return True

    def start_process(self) -> None:
        if self.worker and self.worker.isRunning():
            return
        if not self.validate_inputs():
            return

        config = self.build_config()
        if self.current_mode() == "transcribe_minute" and config.whisper_model_name == "large-v3":
            QMessageBox.information(
                self,
                "Advertencia de VRAM",
                "large-v3 puede ocupar bastante VRAM. Si Ollama queda lento, use modo Solo transcripcion y luego Minuta desde TXT.",
            )

        self.current_summary = None
        self.started_at = time.time()
        self.progress_done = 0
        self.progress_total = 0
        self.general_bar.setValue(0)
        self.file_bar.setValue(0)
        self.open_results_btn.setEnabled(False)
        self.export_report_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("Iniciando...")
        self.timer.start(1000)

        self.worker = ProcessWorker(self.current_mode(), self.selected_inputs, config)
        self.worker.log.connect(self.append_log)
        self.worker.status.connect(self.status_label.setText)
        self.worker.general_progress.connect(self.on_general_progress)
        self.worker.file_progress.connect(self.file_bar.setValue)
        self.worker.current_file.connect(lambda p: self.current_file_label.setText(f"Archivo actual: {p}"))
        self.worker.error.connect(lambda e: self.append_log(f"ERROR: {e}"))
        self.worker.finished_summary.connect(self.on_finished)
        self.worker.fatal_error.connect(self.on_fatal_error)
        self.worker.start()

    def stop_after_current(self) -> None:
        if self.worker and self.worker.isRunning():
            self.worker.stop_after_current()
            self.append_log("Se detendra despues del archivo actual.")
            self.stop_btn.setEnabled(False)

    def on_general_progress(self, done: int, total: int) -> None:
        self.progress_done = done
        self.progress_total = total
        self.general_bar.setMaximum(max(1, total))
        self.general_bar.setValue(min(done, total))
        self.update_time_labels()

    def update_time_labels(self) -> None:
        if not self.started_at:
            return
        elapsed = time.time() - self.started_at
        self.elapsed_label.setText(f"Tiempo transcurrido: {format_seconds(elapsed)}")
        if self.progress_done > 0 and self.progress_total > self.progress_done:
            remaining = elapsed / self.progress_done * (self.progress_total - self.progress_done)
            self.eta_label.setText(f"Tiempo estimado restante: {format_seconds(remaining)}")
        else:
            self.eta_label.setText("Tiempo estimado restante: -")

    def on_finished(self, summary: ProcessSummary) -> None:
        self.timer.stop()
        self.current_summary = summary
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.open_results_btn.setEnabled(bool(summary.output_root))
        self.export_report_btn.setEnabled(True)
        self.status_label.setText("Proceso terminado.")
        self.append_log(
            f"Proceso terminado. Correctos: {summary.ok_count}, errores: {summary.error_count}, "
            f"minutas: {summary.minute_count}, tiempo: {format_seconds(summary.elapsed_seconds)}."
        )
        ResultDialog(summary, self).exec()

    def on_fatal_error(self, message: str) -> None:
        self.timer.stop()
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("Error")
        self.append_log(f"ERROR FATAL: {message}")
        QMessageBox.critical(self, "Error", message)

    def open_results_folder(self) -> None:
        if self.current_summary and self.current_summary.output_root:
            abrir_carpeta(self.current_summary.output_root)
            return
        if self.output_path.text().strip():
            abrir_carpeta(self.output_path.text().strip())

    def save_log(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Guardar registro", "registro_gamboa_transcriptor.txt", "TXT (*.txt)")
        if path:
            Path(path).write_text(self.log_view.toPlainText(), encoding="utf-8")

    def export_report(self) -> None:
        if not self.current_summary:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exportar reporte", "reporte_ejecucion.json", "JSON (*.json);;TXT (*.txt)")
        if path:
            out = exportar_reporte_ejecucion(self.current_summary, path)
            self.append_log(f"Reporte exportado: {out}")

    # ------------------------------------------------------------------ Style
    def apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background: #0c1117;
                color: #edf4f8;
                font-family: "Segoe UI";
                font-size: 10.5pt;
            }
            QScrollArea {
                border: 0;
            }
            #Header {
                background: #111a24;
                border: 1px solid #203040;
                border-radius: 10px;
            }
            #LogoBox {
                background: #173b4f;
                border: 1px solid #2a8fb0;
                border-radius: 8px;
                color: #f2c86b;
                font-weight: 800;
                font-size: 20pt;
            }
            #CompanyTitle {
                color: #f2c86b;
                font-weight: 700;
                font-size: 13pt;
            }
            #MainTitle {
                color: #ffffff;
                font-weight: 800;
                font-size: 20pt;
            }
            #Subtitle, #MutedLabel {
                color: #aebdcc;
            }
            #WarningLabel {
                color: #f2c86b;
                font-weight: 600;
            }
            #DialogTitle {
                font-size: 18pt;
                font-weight: 800;
                color: #ffffff;
            }
            QGroupBox#Card {
                background: #111820;
                border: 1px solid #263545;
                border-radius: 8px;
                margin-top: 18px;
                padding: 14px;
                font-weight: 700;
            }
            QGroupBox#Card::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                color: #d8e4ee;
            }
            QLineEdit, QTextEdit, QComboBox, QSpinBox {
                background: #0b1118;
                border: 1px solid #2c3d4f;
                border-radius: 6px;
                padding: 7px;
                color: #f2f6f8;
                selection-background-color: #1f6f8b;
            }
            QPushButton {
                background: #183348;
                border: 1px solid #2f6075;
                border-radius: 7px;
                padding: 8px 12px;
                color: #f4fbff;
                font-weight: 650;
            }
            QPushButton:hover {
                background: #21455f;
            }
            QPushButton:disabled {
                background: #1a222b;
                color: #657381;
                border-color: #26313c;
            }
            QPushButton#PrimaryButton {
                background: #0f7f5c;
                border: 1px solid #22a879;
                font-size: 13pt;
                padding: 12px;
            }
            QPushButton#PrimaryButton:hover {
                background: #12956c;
            }
            QPushButton#SecondaryButton {
                background: #162536;
            }
            QRadioButton, QCheckBox {
                spacing: 8px;
                color: #edf4f8;
            }
            QProgressBar {
                background: #0b1118;
                border: 1px solid #2c3d4f;
                border-radius: 6px;
                text-align: center;
                height: 22px;
            }
            QProgressBar::chunk {
                background: #0f7f5c;
                border-radius: 5px;
            }
            QLabel#StatusCard {
                background: #0b1118;
                border-radius: 7px;
                padding: 10px;
                border: 1px solid #2c3d4f;
            }
            QLabel#StatusCard[statusClass="status-green"] {
                border-color: #1fa873;
                color: #a9f0cc;
            }
            QLabel#StatusCard[statusClass="status-yellow"] {
                border-color: #c49a38;
                color: #ffe0a3;
            }
            QLabel#StatusCard[statusClass="status-red"] {
                border-color: #bf4d55;
                color: #ffb5bd;
            }
            QLabel#StatusCard[statusClass="status-blue"] {
                border-color: #2c78a0;
                color: #bdddf1;
            }
            """
        )


def main() -> int:
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName(APP_SHORT_NAME)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
