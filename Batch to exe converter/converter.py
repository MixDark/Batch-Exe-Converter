import sys
import time
import os
import subprocess
import tempfile
import shutil
from PyQt6.QtCore import QThread, pyqtSignal
import logging
from datetime import datetime

# Configuración del sistema de logging
def setup_logging(level=logging.INFO):
    log_directory = "logs"
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)
    
    log_filename = os.path.join(log_directory, f"converter_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return logging.getLogger(__name__)

# Crear el logger
logger = setup_logging()

class ConversionWorker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, config):
        super().__init__()
        self.config = config
        self.require_admin = config.get('require_admin', True)
        self.process = None
        self.temp_files = []
        logger.debug("ConversionWorker inicializado con config: %s", self.config)

    def cleanup_temp_files(self):
        """Limpia todos los archivos temporales generados"""
        if not self.config.get('keep_temp_files', False):
            logger.debug("Iniciando limpieza de archivos temporales")
            for file_path in self.temp_files:
                try:
                    if os.path.exists(file_path):
                        if os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                        else:
                            os.remove(file_path)
                        logger.debug("Eliminado: %s", file_path)
                except Exception as e:
                    logger.error("Error limpiando archivo temporal %s: %s", file_path, e)
        else:
            logger.debug("No se eliminarán los archivos temporales")

    def generate_cs_template(self, bat_content):
        """Genera el template de C# para la conversión"""
        logger.debug("Generando template C#")

        try:
            # Escapar el contenido del batch para C#
            escaped_bat_content = (
                bat_content
                .replace("\\", "\\\\")
                .replace("\"", "\\\"")
                .replace("\r\n", "\\r\\n")
                .replace("\n", "\\r\\n")
                .replace("\t", "\\t")
                .replace("$", "$$")
            )

            # Verificar que el contenido escapado sea válido
            logger.debug("Generando template con contenido escapado")

            # Solo incluir la clase RequireAdministrator si se requieren privilegios de administrador
            admin_check_class = '''
        [System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
        [System.Runtime.InteropServices.ComVisible(false)]
        public class RequireAdministrator
        {
            [System.Runtime.InteropServices.DllImport("shell32.dll", PreserveSig = false)]
            private static extern void IsUserAnAdmin();

            public static bool Check()
            {
                try
                {
                    IsUserAnAdmin();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }''' if self.config.get('admin_required', False) else ''

            template = f'''
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Text;
    using System.Windows.Forms;
    using System.Threading;
    using System.Threading.Tasks;

    namespace BatchExecutor
    {{
        public class BatchExecutorConfig
        {{
            private string _batFilePrefix = "batch_";
            private int _deleteDelayMs = 1000;
            private string _batchContent = "{escaped_bat_content}";

            public string BatFilePrefix
            {{
                get {{ return _batFilePrefix; }}
                set {{ _batFilePrefix = value; }}
            }}

            public int DeleteDelayMs
            {{
                get {{ return _deleteDelayMs; }}
                set {{ _deleteDelayMs = value; }}
            }}

            public string BatchContent
            {{
                get {{ return _batchContent; }}
                set {{ _batchContent = value; }}
            }}
        }}

        public class BatchExecutor : IDisposable
        {{
            private readonly string _tempBatFile;
            private readonly BatchExecutorConfig _config;
            private bool _disposed;

            public BatchExecutor(BatchExecutorConfig config)
            {{
                _config = config ?? new BatchExecutorConfig();
                _tempBatFile = GenerateTempBatPath();
            }}

            private string GenerateTempBatPath()
            {{
                return Path.Combine(
                    Path.GetTempPath(),
                    string.Format("{{0}}{{1}}.bat", _config.BatFilePrefix, Guid.NewGuid().ToString("N"))
                );
            }}

            private ProcessStartInfo CreateStartInfo()
            {{
                return new ProcessStartInfo
                {{
                    FileName = "cmd.exe",
                    Arguments = string.Format("/C \\"{{0}}\\"", _tempBatFile),
                    UseShellExecute = true,
                    WorkingDirectory = Application.StartupPath,
                    WindowStyle = ProcessWindowStyle.Normal,
                    CreateNoWindow = false
                }};
            }}

            public void Execute()
            {{
                try
                {{
                    CreateBatchFile();
                    ExecuteProcess();
                }}
                catch (Exception ex)
                {{
                    MessageBox.Show(
                        string.Format("Error ejecutando el archivo batch:\\n{{0}}", ex.Message),
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    throw;
                }}
                finally
                {{
                    CleanupTempFile();
                }}
            }}

            private void CreateBatchFile()
            {{
                ValidateNotDisposed();
                try
                {{
                    Encoding encoding;
                    try
                    {{
                        encoding = Encoding.GetEncoding("IBM437");
                    }}
                    catch
                    {{
                        try
                        {{
                            encoding = Encoding.GetEncoding("Windows-1252");
                        }}
                        catch
                        {{
                            encoding = Encoding.Default;
                        }}
                    }}

                    using (var writer = new StreamWriter(_tempBatFile, false, encoding))
                    {{
                        writer.Write(_config.BatchContent);
                        writer.Flush();
                    }}
                }}
                catch (Exception ex)
                {{
                    MessageBox.Show(
                        string.Format("Error al crear el archivo batch: {{0}}", ex.Message),
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    throw;
                }}
            }}

            private void ExecuteProcess()
            {{
                ValidateNotDisposed();
                using (var process = Process.Start(CreateStartInfo()))
                {{
                    if (process == null)
                    {{
                        throw new InvalidOperationException("No se pudo iniciar el proceso");
                    }}
                    process.WaitForExit();
                }}
            }}

            private void CleanupTempFile()
            {{
                if (File.Exists(_tempBatFile))
                {{
                    try
                    {{
                        Thread.Sleep(_config.DeleteDelayMs);
                        File.Delete(_tempBatFile);
                    }}
                    catch (Exception)
                    {{
                        // Ignorar errores al eliminar archivo temporal
                    }}
                }}
            }}

            private void ValidateNotDisposed()
            {{
                if (_disposed)
                {{
                    throw new ObjectDisposedException("BatchExecutor");
                }}
            }}

            public void Dispose()
            {{
                if (!_disposed)
                {{
                    CleanupTempFile();
                    _disposed = true;
                }}
            }}
        }}

        {admin_check_class}

        public class Program
        {{
            [STAThread]
            static void Main()
            {{
                try
                {{
                    {'''if (!RequireAdministrator.Check())
                    {
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.UseShellExecute = true;
                        startInfo.WorkingDirectory = Environment.CurrentDirectory;
                        startInfo.FileName = Application.ExecutablePath;
                        startInfo.Verb = "runas";
                        
                        try
                        {
                            Process.Start(startInfo);
                            return;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(
                                "Esta aplicación requiere privilegios de administrador para ejecutarse.",
                                "Error de Privilegios",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                            return;
                        }
                    }''' if self.config.get('admin_required', False) else ''}

                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    var config = new BatchExecutorConfig();
                    
                    using (var executor = new BatchExecutor(config))
                    {{
                        executor.Execute();
                    }}
                }}
                catch (UnauthorizedAccessException ex)
                {{
                    MessageBox.Show(
                        string.Format("Error de permisos:\\n\\n{{0}}", ex.Message),
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }}
                catch (IOException ex)
                {{
                    MessageBox.Show(
                        string.Format("Error de E/S:\\n\\n{{0}}", ex.Message),
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }}
                catch (Exception ex)
                {{
                    MessageBox.Show(
                        string.Format("Error inesperado:\\n\\n{{0}}", ex.Message),
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }}
            }}
        }}
    }}'''

            # Verificar que el template se generó correctamente
            if not template or len(template.strip()) == 0:
                raise ValueError("El template generado está vacío")

            logger.debug("Template C# generado exitosamente")
            
            # Guardar el template en un archivo temporal para debugging si está habilitado
            if self.config.get('debug_mode', False):
                debug_file = "debug_template.cs"
                with open(debug_file, 'w', encoding='utf-8') as f:
                    f.write(template)
                logger.debug(f"Template guardado para debugging en: {debug_file}")

            return template

        except Exception as e:
            logger.error(f"Error generando el template de C#: {str(e)}")
            raise Exception(f"Error al generar el template de C#: {str(e)}")

    def compile_cs_to_exe(self, cs_file, output_exe):
        """Compila el archivo C# a ejecutable"""
        try:
            # Asegurar que el directorio de salida existe
            output_dir = os.path.dirname(output_exe)
            os.makedirs(output_dir, exist_ok=True)

            # Verificar que el archivo C# existe
            if not os.path.exists(cs_file):
                raise FileNotFoundError(f"No se encuentra el archivo fuente: {cs_file}")

            # Verificar el contenido del archivo
            with open(cs_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if not content.strip():
                    raise ValueError("El archivo fuente está vacío")

            # Buscar el compilador de C#
            framework_paths = [
                r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319",
                r"C:\Windows\Microsoft.NET\Framework\v4.0.30319",
            ]

            csc_path = None
            for base_path in framework_paths:
                possible_csc = os.path.join(base_path, "csc.exe")
                if os.path.exists(possible_csc):
                    csc_path = possible_csc
                    break

            if not csc_path:
                raise FileNotFoundError("No se encontró el compilador de C# (csc.exe)")

            logger.info(f"Usando compilador: {csc_path}")

            # Solo crear el archivo de manifiesto si se requieren privilegios de administrador
            manifest_file = None
            if self.config.get('admin_required', False):
                manifest_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
        <assemblyIdentity version="1.0.0.0" name="MyApplication.app"/>
        <trustInfo xmlns="urn:schemas-microsoft-com:asm.v2">
            <security>
                <requestedPrivileges xmlns="urn:schemas-microsoft-com:asm.v3">
                    <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
                </requestedPrivileges>
            </security>
        </trustInfo>
    </assembly>'''

                manifest_file = os.path.join(os.path.dirname(cs_file), "app.manifest")
                with open(manifest_file, 'w', encoding='utf-8') as f:
                    f.write(manifest_content)
                self.temp_files.append(manifest_file)

            # Construir el comando de compilación
            command = [
                csc_path,
                '/nologo',
                '/target:winexe',
                '/platform:anycpu',
                '/optimize+',
                '/debug-',
                '/reference:System.dll',
                '/reference:System.Windows.Forms.dll',
                '/reference:System.Drawing.dll',
                '/reference:System.Core.dll',
            ]

            # Agregar el manifiesto solo si existe
            if manifest_file:
                command.append(f'/win32manifest:{manifest_file}')

            # Agregar el resto de los parámetros
            command.extend([
                f'/out:{output_exe}',
                cs_file
            ])

            # Agregar icono si existe
            if self.config.get('icon_file') and os.path.exists(self.config['icon_file']):
                command.append(f'/win32icon:{self.config["icon_file"]}')

            # Ejecutar el proceso de compilación
            logger.debug(f"Ejecutando comando: {' '.join(command)}")
            
            # Crear el proceso con un shell explícito
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                env={
                    "PATH": os.environ["PATH"],
                    "SystemRoot": os.environ["SystemRoot"],
                    "TEMP": os.environ["TEMP"],
                    "TMP": os.environ["TMP"]
                }
            )

            # Obtener la salida con timeout
            try:
                stdout, stderr = process.communicate(timeout=30)
                
                # Registrar la salida para debugging
                logger.debug(f"Salida estándar: {stdout}")
                logger.debug(f"Salida de error: {stderr}")

                if process.returncode != 0:
                    error_msg = stderr.strip() if stderr else stdout.strip()
                    if not error_msg:
                        error_msg = f"Error de compilación con código {process.returncode}"
                    raise Exception(f"Error en la compilación: {error_msg}")

                # Verificar que el archivo se creó
                if not os.path.exists(output_exe):
                    raise FileNotFoundError(f"No se generó el archivo ejecutable: {output_exe}")

                # Verificar el tamaño del archivo
                if os.path.getsize(output_exe) == 0:
                    raise ValueError("El archivo ejecutable generado está vacío")

                logger.info("Compilación exitosa")
                return True

            except subprocess.TimeoutExpired:
                process.kill()
                raise Exception("Tiempo de espera agotado durante la compilación")

        except Exception as e:
            logger.error(f"Error durante la compilación: {str(e)}")
            # Intentar obtener más información sobre el error
            if os.path.exists(cs_file):
                logger.debug(f"Contenido del archivo fuente:")
                with open(cs_file, 'r', encoding='utf-8') as f:
                    logger.debug(f.read())
            raise Exception(f"Error en la compilación: {str(e)}")

    def check_csc_compiler(self):
        """Verifica que el compilador C# esté disponible"""
        logger.debug("Verificando disponibilidad del compilador C#")
        try:
            process = subprocess.run(
                ['csc', '/help'], 
                capture_output=True, 
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return process.returncode == 0
        except FileNotFoundError:
            logger.error("Compilador C# no encontrado")
            return False

    def check_dependencies(self):
        """Verifica todas las dependencias necesarias"""
        logger.debug("Verificando dependencias del sistema")
        
        # Verificar .NET Framework
        try:
            key_path = r'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                version = winreg.QueryValueEx(key, 'Version')[0]
                logger.info(f"Versión de .NET Framework encontrada: {version}")
        except Exception as e:
            logger.error("No se encontró .NET Framework 4.0 o superior")
            raise Exception("Se requiere .NET Framework 4.0 o superior")


    def run(self):
        """Ejecuta el proceso de conversión"""
        logger.info("Iniciando proceso de conversión")
        temp_cs_file = os.path.abspath('temp_script.cs')
        self.temp_files.append(temp_cs_file)

        try:
            # Verificar dependencias
            self.status.emit("Verificando dependencias...")
            self.progress.emit(5)
            self.check_dependencies()

            # Verificar compilador C#
            self.status.emit("Verificando compilador C#...")
            self.progress.emit(10)
            if not self.check_csc_compiler():
                raise Exception("Compilador C# no encontrado")

            # Leer archivo BAT
            self.status.emit("Leyendo archivo batch...")
            self.progress.emit(20)
            with open(self.config['batch_file'], 'r', encoding='utf-8', errors='replace') as f:
                bat_content = f.read()

            # Generar archivo C#
            self.status.emit("Generando código C#...")
            self.progress.emit(40)
            cs_content = self.generate_cs_template(bat_content)
            
            # Guardar archivo C#
            with open(temp_cs_file, 'w', encoding='utf-8') as f:
                f.write(cs_content)

            # Verificar archivo generado
            if not os.path.exists(temp_cs_file):
                raise FileNotFoundError(f"No se pudo crear el archivo: {temp_cs_file}")

            # Preparar compilación
            self.status.emit("Preparando compilación...")
            self.progress.emit(60)
            
            output_dir = self.config.get('output_dir', 'dist')
            os.makedirs(output_dir, exist_ok=True)
            
            output_exe = os.path.join(
                output_dir,
                self.config.get('output_name', 'output') + '.exe'
            )

            # Compilar
            self.status.emit("Compilando ejecutable...")
            self.progress.emit(80)
            
            success = self.compile_cs_to_exe(temp_cs_file, output_exe)
            
            if not success:
                raise Exception("La compilación falló sin error específico")

            self.progress.emit(100)
            self.status.emit("¡Conversión completada!")
            self.finished.emit()

        except Exception as e:
            logger.exception("Error durante la conversión")
            self.error.emit(str(e))
        finally:
            self.cleanup_temp_files()