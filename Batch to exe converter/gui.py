import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout,
    QWidget, QCheckBox, QLabel, QMessageBox, QProgressBar, QComboBox, QGroupBox, QLineEdit
)
from PyQt6.QtCore import Qt, QSettings, pyqtSignal
from PyQt6.QtGui import QIcon, QDragEnterEvent, QDropEvent
import json
import os
import tempfile
import shutil
import subprocess
from converter import ConversionWorker

class DropWidget(QWidget):
    fileDropped = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for file in files:
            if file.endswith('.bat'):
                self.fileDropped.emit(file)
                break

class BatchConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_compiler()  # Verificar y configurar el compilador C#
    
        self.settings = QSettings('BatchConverter', 'Settings')
        self.output_dir = ''
        self.load_settings()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Batch to EXE Converter')
        self.setGeometry(100, 100, 600, 400)
        self.setWindowIcon(QIcon("favicon.ico"))
        self.setWindowFlags(Qt.WindowType.WindowCloseButtonHint | Qt.WindowType.WindowMinimizeButtonHint)


        # Widget central con soporte para drag and drop
        central_widget = DropWidget()
        central_widget.fileDropped.connect(self.handle_dropped_file)
        self.setCentralWidget(central_widget)
        
        # Crear el layout principal
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # Grupo de selección de archivos
        file_group = QGroupBox("Selección de archivos")
        file_layout = QVBoxLayout()
        
        self.file_label = QLabel('Arrastra y suelta un archivo .bat o selecciónalo')
        file_layout.addWidget(self.file_label)
        
        self.select_button = QPushButton('Seleccionar archivo .bat', self)
        self.select_button.clicked.connect(self.select_file)
        file_layout.addWidget(self.select_button)
        
        self.icon_button = QPushButton('Seleccionar icono (.ico)', self)
        self.icon_button.clicked.connect(self.select_icon)
        file_layout.addWidget(self.icon_button)
        
        self.output_dir_button = QPushButton('Seleccionar carpeta de salida', self)
        self.output_dir_button.clicked.connect(self.select_output_dir)
        file_layout.addWidget(self.output_dir_button)

        # Etiqueta para mostrar la ruta de salida (con el mismo estilo que file_label)
        self.output_dir_label = QLabel('Carpeta de salida: No seleccionada')
        self.output_dir_label.setWordWrap(True)  # Permite que el texto se ajuste en múltiples líneas
        file_layout.addWidget(self.output_dir_label)
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # Grupo de opciones
        options_group = QGroupBox("Opciones de conversión")
        options_layout = QVBoxLayout()
        
        self.output_name = QLineEdit(self)
        self.output_name.setPlaceholderText('Nombre del ejecutable (obligatorio)')
        options_layout.addWidget(self.output_name)
                
        self.console_checkbox = QCheckBox('Mostrar consola', self)
        options_layout.addWidget(self.console_checkbox)
        
        self.center_checkbox = QCheckBox('Centrar ventana al ejecutar', self)
        options_layout.addWidget(self.center_checkbox)

        self.admin_checkbox = QCheckBox('Ejecutar como administrador', self)
        options_layout.addWidget(self.admin_checkbox)
        
        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)

        # Grupo de tema
        theme_group = QGroupBox("Personalización")
        theme_layout = QVBoxLayout()
        
        theme_label = QLabel('Tema:')
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(['Claro', 'Oscuro', 'Azul', 'Verde'])
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        theme_layout.addWidget(theme_label)
        theme_layout.addWidget(self.theme_combo)
        
        theme_group.setLayout(theme_layout)
        main_layout.addWidget(theme_group)

        # Etiqueta de estado
        self.status_label = QLabel('Listo')
        main_layout.addWidget(self.status_label)

        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Botón de conversión
        self.convert_button = QPushButton('Convertir a EXE', self)
        self.convert_button.clicked.connect(self.convert_to_exe)
        main_layout.addWidget(self.convert_button)

        # Variables para almacenar rutas
        self.batch_file = ''
        self.icon_file = ''

        # Cargar preferencias guardadas
        self.load_preferences()
        
        # Centrar la ventana
        self.center()

    def update_status(self, status):
        self.status_label.setText(status)            

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            'Seleccionar archivo batch',
            '',
            'Batch files (*.bat)'
        )
        if file_name:
            self.batch_file = file_name
            self.file_label.setText(f'Archivo seleccionado: {os.path.basename(file_name)}')
            self.save_settings()

    def select_icon(self):
        icon_name, _ = QFileDialog.getOpenFileName(
            self,
            'Seleccionar icono',
            '',
            'Icon files (*.ico)'
        )
        if icon_name:
            self.icon_file = icon_name
            self.save_settings()

    def select_output_dir(self):
        dir_name = QFileDialog.getExistingDirectory(self, 'Seleccionar carpeta de salida')
        if dir_name:
            self.output_dir = dir_name
            self.output_dir_label.setText(f'Ruta de salida: {dir_name}')
            self.save_settings()

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def load_settings(self):
        self.batch_file = self.settings.value('batch_file', '')
        self.icon_file = self.settings.value('icon_file', '')
        self.output_dir = self.settings.value('output_dir', '')
        self.theme = self.settings.value('theme', 'Claro')

    def save_settings(self):
        self.settings.setValue('batch_file', self.batch_file)
        self.settings.setValue('icon_file', self.icon_file)
        self.settings.setValue('output_dir', self.output_dir)
        self.settings.setValue('theme', self.theme_combo.currentText())

    def handle_dropped_file(self, file_path):
        self.batch_file = file_path
        self.file_label.setText(f'Archivo seleccionado: {os.path.basename(file_path)}')
        self.save_settings()

    def change_theme(self, theme_name):
        themes = {
            'Claro': """
                QMainWindow, QWidget { background-color: #ffffff; color: #000000; }
                QPushButton { background-color: #e0e0e0; border: 1px solid #b0b0b0; padding: 5px; }
                QPushButton:hover { background-color: #d0d0d0; }
                QLineEdit { padding: 5px; border: 1px solid #b0b0b0; }
            """,
            'Oscuro': """
                QMainWindow, QWidget { background-color: #2b2b2b; color: #ffffff; }
                QPushButton { background-color: #404040; color: #ffffff; border: 1px solid #555555; padding: 5px; }
                QPushButton:hover { background-color: #505050; }
                QLineEdit { padding: 5px; border: 1px solid #555555; background-color: #404040; color: #ffffff; }
            """,
            'Azul': """
                QMainWindow, QWidget { background-color: #1e3d59; color: #ffffff; }
                QPushButton { background-color: #2a5f8f; color: #ffffff; border: 1px solid #3d84c6; padding: 5px; }
                QPushButton:hover { background-color: #3d84c6; }
                QLineEdit { padding: 5px; border: 1px solid #3d84c6; background-color: #2a5f8f; color: #ffffff; }
            """,
            'Verde': """
                QMainWindow, QWidget { background-color: #2e4a3e; color: #ffffff; }
                QPushButton { background-color: #3e6351; color: #ffffff; border: 1px solid #528f6f; padding: 5px; }
                QPushButton:hover { background-color: #528f6f; }
                QLineEdit { padding: 5px; border: 1px solid #528f6f; background-color: #3e6351; color: #ffffff; }
            """
        }
        self.setStyleSheet(themes.get(theme_name, themes['Claro']))
        self.save_settings()

    def convert_to_exe(self):
        if not self.batch_file:
            QMessageBox.warning(self, 'Error', 'Por favor seleccione un archivo batch')
            return
        
        if not self.output_name.text().strip():
            QMessageBox.warning(self, 'Error', 'Debe ingresar un nombre para el ejecutable')
            self.output_name.setFocus() 
            return

        # Verificar permisos de escritura en el directorio de salida
        output_dir = self.output_dir or 'dist'
        try:
            os.makedirs(output_dir, exist_ok=True)
            test_file = os.path.join(output_dir, 'test.tmp')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
        except Exception as e:
            QMessageBox.critical(
                self, 
                'Error', 
                f'No hay permisos de escritura en el directorio de salida: {str(e)}'
            )
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.convert_button.setEnabled(False)
        self.status_label.setText('Iniciando conversión...')

        config = {
            'batch_file': self.batch_file,
            'icon_file': self.icon_file,
            'output_dir': output_dir,
            'output_name': self.output_name.text(),
            'console': self.console_checkbox.isChecked(),
            'center_window': self.center_checkbox.isChecked(),
            'admin_required': self.admin_checkbox.isChecked(),
            'keep_temp_files': False 
        }

        self.worker = ConversionWorker(config)
        self.worker.progress.connect(self.update_progress)
        self.worker.status.connect(self.update_status)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_error)
        self.worker.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def conversion_finished(self):
        self.progress_bar.setVisible(False)
        self.convert_button.setEnabled(True)
        output_path = os.path.join(
            self.output_dir or 'dist',
            f"{self.output_name.text()}.exe"
        )
        QMessageBox.information(
            self, 
            'Éxito', 
            f'Conversión completada exitosamente.\nArchivo creado: {output_path}'
        )

    def conversion_error(self, error_message):
        self.progress_bar.setVisible(False)
        self.convert_button.setEnabled(True)
        QMessageBox.critical(self, 'Error', f'Error durante la conversión: {error_message}')

    def load_preferences(self):
        try:
            if os.path.exists('preferences.json'):
                with open('preferences.json', 'r') as f:
                    prefs = json.load(f)
                    self.theme_combo.setCurrentText(prefs.get('theme', 'Claro'))
                    self.console_checkbox.setChecked(prefs.get('console', False))
                    self.center_checkbox.setChecked(prefs.get('center_window', False))
                    self.admin_checkbox.setChecked(prefs.get('admin_required', False))
                    self.output_name.setText(prefs.get('output_name', ''))
        except Exception as e:
            print(f"Error loading preferences: {e}")

    def save_preferences(self):
        prefs = {
            'theme': self.theme_combo.currentText(),
            'console': self.console_checkbox.isChecked(),
            'center_window': self.center_checkbox.isChecked(),
            'admin_required': self.admin_checkbox.isChecked(),
            'output_name': self.output_name.text()
        }
        try:
            with open('preferences.json', 'w') as f:
                json.dump(prefs, f)
        except Exception as e:
            print(f"Error saving preferences: {e}")

    def closeEvent(self, event):
        self.cleanup_temp_files()
        self.save_preferences()
        self.save_settings()
        event.accept()

    def cleanup_temp_files(self):
        """Limpia archivos temporales residuales"""
        temp_dir = os.path.join(tempfile.gettempdir(), 'batch_converter_temp')
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except Exception as e:
            print(f"Error al limpiar archivos temporales: {e}")


    def show_info_message(self):
        QMessageBox.information(
            self,
            'Información',
            'Esta aplicación requiere .NET Framework para funcionar.\n' +
            'Asegúrese de tener instalado .NET Framework y que el compilador C# (csc.exe) ' +
            'esté en el PATH del sistema.'
        ) 


    def check_csc_compiler(self):
        """Verifica la disponibilidad del compilador C# y su ubicación"""
        possible_paths = [
            r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
            r"C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
        ]
        
        # Primero intentar con el PATH del sistema
        try:
            subprocess.run(
                ['csc', '/help'],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return True, "PATH del sistema"
        except FileNotFoundError:
            # Si no está en el PATH, buscar en las ubicaciones conocidas
            for path in possible_paths:
                if os.path.exists(path):
                    return True, path
            
            return False, None

    def setup_compiler(self):
        """Configura el acceso al compilador C#"""
        is_available, compiler_path = self.check_csc_compiler()
        
        if not is_available:
            QMessageBox.critical(
                self,
                'Error',
                'No se encontró el compilador C# (csc.exe).\n'
                'Para resolver esto:\n\n'
                '1. Abra el Panel de Control\n'
                '2. Vaya a "Programas y características"\n'
                '3. Habilite ".NET Framework 4.8 Advanced Services"\n'
                '4. Reinicie el sistema'
            )
            sys.exit(1)
        else:
            # Si el compilador no está en el PATH pero existe en una ubicación conocida
            if compiler_path != "PATH del sistema":
                # Agregar al PATH temporal
                os.environ['PATH'] = f"{os.path.dirname(compiler_path)};{os.environ['PATH']}"

    def show_dotnet_download_info(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText("Descarga de .NET Framework")
        msg.setInformativeText(
            "Se recomienda .NET Framework 4.8 o superior.\n"
            "¿Desea abrir la página de descarga?"
        )
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg.setDefaultButton(QMessageBox.StandardButton.Yes)
        
        if msg.exec() == QMessageBox.StandardButton.Yes:
            import webbrowser
            webbrowser.open("https://dotnet.microsoft.com/download/dotnet-framework")                      

def main():
    app = QApplication(sys.argv)
    ex = BatchConverter()
    ex.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()