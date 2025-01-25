import sys
import os
import subprocess
import time
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QFileDialog, QWidget, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from PyQt5.QtMultimedia import QSoundEffect
from PyQt5.QtCore import QUrl
import ctypes
from ctypes import wintypes, windll, byref, POINTER, c_wchar_p
import re
import shutil
import requests

class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", wintypes.BYTE * 8)
    ]

FOLDERID_Videos = GUID(
    0x18989B1D,
    0x99B5,
    0x455B,
    (0x84, 0x1C, 0xAB, 0x7C, 0x74, 0xE4, 0xDD, 0xFC)
)

def get_videos_folder():
    windll.shell32.SHGetKnownFolderPath.restype = ctypes.HRESULT
    windll.shell32.SHGetKnownFolderPath.argtypes = [
        POINTER(GUID),
        wintypes.DWORD,
        wintypes.HANDLE,
        POINTER(c_wchar_p)
    ]
    path_ptr = c_wchar_p()
    hresult = windll.shell32.SHGetKnownFolderPath(
        byref(FOLDERID_Videos),
        0,
        None,
        byref(path_ptr)
    )
    if hresult != 0:
        raise ctypes.WinError(hresult)
    path = path_ptr.value
    windll.ole32.CoTaskMemFree(path_ptr)
    return path

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_ffmpeg_path():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
    ffmpeg_path = os.path.join(base_dir, 'ffmpeg', 'ffmpeg.exe')
    ffprobe_path = os.path.join(base_dir, 'ffmpeg', 'ffprobe.exe')
    
    if not os.path.exists(ffmpeg_path):
        raise FileNotFoundError("FFmpeg binaries not found!")
        
    return ffmpeg_path, ffprobe_path

class VideoCompressorThread(QThread):
    log_signal = pyqtSignal(str)
    done_signal = pyqtSignal(str, float, str)
    error_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)

    def __init__(self, input_file, output_file, target_size, target_bitrate, start_time=None, end_time=None):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.target_size = target_size
        self.target_bitrate = target_bitrate
        self.start_time = start_time
        self.end_time = end_time
        self.last_progress = 0
        
        appdata_path = os.getenv('APPDATA')
        self.cliputils_folder = os.path.join(appdata_path, "ClipUtils")
        os.makedirs(self.cliputils_folder, exist_ok=True)
        
        self.temp_files = []

    def run(self):
        start_time_total = time.time()
        success = False

        try:
            if (self.target_size is not None and self.target_size > 0) and \
               (self.start_time is not None and self.end_time is not None):
                self.compress_trim_video()
                success = True

            elif self.target_size is not None and self.target_size > 0:
                self.compress_video()
                success = True

            else:
                self.trim_video()
                self.progress_signal.emit(100)
                success = True

        except Exception as e:
            self.error_signal.emit(f"Error: {str(e)}")
            self.cleanup_temp_files()

        if success:
            compression_time = round(time.time() - start_time_total, 1)
            new_size = self.get_file_size(self.output_file)
            self.done_signal.emit(self.output_file, compression_time, f"{new_size} MB")
            self.cleanup_temp_files()

    def compress_trim_video(self):
        temp_filename = f"temp_{os.path.basename(self.input_file)}"
        self.temp_file = os.path.join(self.cliputils_folder, temp_filename)
        self.temp_files.append(self.temp_file)
        self.trim_video(output_file=self.temp_file)
        trimmed_size = self.get_file_size(self.temp_file)
        if trimmed_size <= self.target_size:
            try:
                shutil.move(self.temp_file, self.output_file)
                return
            except Exception as e:
                raise RuntimeError(f"Failed to use trimmed video: {str(e)}")

        duration = self.get_video_duration(self.temp_file)
        if duration is None:
            raise ValueError("Failed to get duration of trimmed video.")

        target_bitrate = (self.target_size * 8 * 1024) / duration
        self.compress_video(input_file=self.temp_file, target_bitrate=target_bitrate)

    def trim_video(self, output_file=None):
        output = output_file or self.output_file
        ffmpeg_path, ffprobe_path = get_ffmpeg_path()
        
        duration = self.end_time - self.start_time
        
        command = [
            ffmpeg_path,
            "-ss", str(self.start_time),
            "-i", self.input_file,
            "-to", str(duration),
            "-c:v", "copy",
            "-c:a", "copy",
            "-avoid_negative_ts", "make_zero",
            "-y", output
        ]
        
        process = subprocess.Popen(
            command, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE, 
            creationflags=subprocess.CREATE_NO_WINDOW, 
            text=True
        )
        
        for line in process.stderr:
            self.log_signal.emit(line.strip())
            
        process.wait()
        
        if process.returncode != 0:
            raise RuntimeError("Trimming failed.")

    def compress_video(self, input_file=None, target_bitrate=None):
        input_file = input_file or self.input_file
        target_bitrate = target_bitrate or self.target_bitrate
        ffmpeg_path, ffprobe_path = get_ffmpeg_path()

        duration = self.get_video_duration(input_file)
        if duration is None:
            raise ValueError("Failed to get video duration.")

        command_pass1 = [
            ffmpeg_path, "-i", input_file,
            "-c:v", "libx264", "-preset", "fast", "-b:v", f"{target_bitrate}k",
            "-pass", "1", "-an", "-f", "mp4", "-y", "NUL"
        ]
        self.run_ffmpeg_command(command_pass1, "pass1_log.txt", pass_number=1, duration=duration)

        command_pass2 = [
            ffmpeg_path, "-i", input_file,
            "-c:v", "libx264", "-preset", "fast", "-b:v", f"{target_bitrate}k",
            "-pass", "2", "-c:a", "copy", "-y", self.output_file
        ]
        self.run_ffmpeg_command(command_pass2, "pass2_log.txt", pass_number=2, duration=duration)

    def run_ffmpeg_command(self, command, log_filename, pass_number, duration):
        log_path = os.path.join(self.cliputils_folder, log_filename)
        with open(log_path, 'w') as log:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                cwd=self.cliputils_folder
            )
            
            if pass_number == 2:
                self.progress_signal.emit(50)
                
            current_line = ""
            for line in process.stderr:
                line = line.strip()
                current_line = line

                if "bitrate=" in current_line:
                    parts = re.split(r'(bitrate=\d+)', current_line, maxsplit=1)
                    if len(parts) >= 3:
                        processed_line = parts[1] + parts[2] + "\n" + parts[0].lstrip()
                    else:
                        processed_line = current_line
                else:
                    processed_line = current_line
                    
                self.log_signal.emit(processed_line)
                    
                if "time=" in line:
                    time_str = line.split('time=')[1].split()[0]
                    current_time = self.parse_time_to_seconds(time_str)
                    if duration > 0:
                        progress = min((current_time / (duration + 2)) * 50, 50)
                        if pass_number == 2:
                            progress += 50
                            progress = min(progress, 100)
                        
                        if int(progress) > self.last_progress:
                            self.progress_signal.emit(int(progress))
                            self.last_progress = int(progress)
                            
                log.write(line.strip() + "\n")
                
            process.wait()
            
            if pass_number == 1:
                self.progress_signal.emit(50)
                self.last_progress = 50

    def parse_time_to_seconds(self, time_str):
        try:
            parts = list(map(float, time_str.replace(',', '.').split(':')))
            if len(parts) == 3:
                return parts[0] * 3600 + parts[1] * 60 + parts[2]
            elif len(parts) == 2:
                return parts[0] * 60 + parts[1]
            else:
                return parts[0]
        except:
            return 0

    def cleanup_temp_files(self):
        for filename in os.listdir(self.cliputils_folder):
            file_path = os.path.join(self.cliputils_folder, filename)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Error deleting {file_path}: {str(e)}")

    def get_file_size(self, file_path):
        return round(os.path.getsize(file_path) / (1024 * 1024), 1)

    def get_video_duration(self, file_path):
        ffmpeg_path, ffprobe_path = get_ffmpeg_path()
        try:
            result = subprocess.run(
                [ffprobe_path, "-v", "error", "-show_entries", "format=duration",
                 "-of", "default=noprint_wrappers=1:nokey=1", file_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW, text=True
            )
            return float(result.stdout.strip())
        except Exception:
            return None


class FileDropWidget(QtWidgets.QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QtGui.QDropEvent):
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            
            if not os.path.isfile(file_path):
                return
                
            if not self.is_valid_video(file_path):
                QMessageBox.critical(
                    self.main_window,
                    "Invalid File",
                    "The dropped file is not a supported video format.\nPlease drop a valid video file.",
                    QMessageBox.Ok
                )
                return

            self.main_window.selected_file = file_path
            self.main_window.file_input_text.setText(f"Selected: {os.path.basename(file_path)}")

            old_size = self.main_window.get_file_size(file_path)
            self.main_window.target_input.setMaximum(old_size)
            self.main_window.old_size_text.setText(f"OLD Filesize: {old_size} MB")
            self.main_window.new_size_text.setText(f"NEW Filesize: N/A")
            duration = round(self.main_window.get_video_duration(file_path), 2)
            self.main_window.old_length_text.setText(f"OLD Video Length: {duration} seconds")
            self.main_window.new_length_text.setText(f"NEW Video Length: {duration} seconds")
            self.main_window.start_input.setMaximum(duration)
            self.main_window.end_input.setMaximum(duration)
            self.main_window.start_input.setValue(0)
            self.main_window.end_input.setValue(duration)
            self.main_window.progress_bar.setValue(0)

    def is_valid_video(self, file_path):
        ffmpeg_path, ffprobe_path = get_ffmpeg_path()
        try:
            stream_check = subprocess.run(
                [ffprobe_path, "-v", "error", "-select_streams", "v:0",
                 "-show_entries", "stream=codec_type", "-of", "default=noprint_wrappers=1:nokey=1", file_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            duration = self.main_window.get_video_duration(file_path)
            
            return duration is not None and "video" in stream_check.stdout
        except Exception:
            return False


class CustomSpinBox(QtWidgets.QDoubleSpinBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSpecialValueText("")
        self.setMinimum(0)
        self.setMaximum(0)
        self.setDecimals(1)
        self.setButtonSymbols(QtWidgets.QAbstractSpinBox.UpDownArrows)

    def textFromValue(self, value):
        if value == self.minimum():
            return ""
        return super().textFromValue(value)

    def valueFromText(self, text):
        if text.strip() == "":
            return self.minimum()
        text = text.replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return self.minimum()

    def focusOutEvent(self, event):
        if self.text().strip() == "":
            self.clear()
        super().focusOutEvent(event)

    def keyPressEvent(self, event):
        if event.key() in (QtCore.Qt.Key_Backspace, QtCore.Qt.Key_Delete):
            self.lineEdit().clear()
            return
        super().keyPressEvent(event)


class VideoCompressor(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_version = "v1.0.1"
        self.check_for_updates()
        self.selected_file = None
        self.sound_effect = QSoundEffect()
        sound_path = resource_path("done.wav")
        self.sound_effect.setSource(QUrl.fromLocalFile(sound_path))

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(500, 350)
        MainWindow.setFixedSize(MainWindow.size())
        MainWindow.setWindowFlags(
            QtCore.Qt.Window | 
            QtCore.Qt.WindowCloseButtonHint | 
            QtCore.Qt.WindowMinimizeButtonHint
        )
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        MainWindow.setFont(font)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.file_box = QtWidgets.QGroupBox(self.centralwidget)
        self.file_box.setGeometry(QtCore.QRect(30, 20, 440, 80))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        self.file_box.setFont(font)
        self.file_box.setAlignment(QtCore.Qt.AlignCenter)
        self.file_box.setObjectName("file_box")
        self.file_input_widget = FileDropWidget(self, self.centralwidget)
        self.file_input_widget.setGeometry(QtCore.QRect(30, 20, 440, 80))
        self.file_input_widget.setObjectName("file_input_widget")
        self.file_input_widget.setAcceptDrops(True)
        self.file_input_text = QtWidgets.QLabel(self.file_input_widget)
        self.file_input_text.setGeometry(QtCore.QRect(10, 15, 420, 60))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(8)
        self.file_input_text.setFont(font)
        self.file_input_text.setStyleSheet("color: rgb(70, 70, 70);")
        self.file_input_text.setText("Drop a video file here to get started...")
        self.file_input_text.setAlignment(QtCore.Qt.AlignCenter)
        self.file_input_text.setObjectName("file_input_text")
        self.target_label = QtWidgets.QLabel(self.centralwidget)
        self.target_label.setGeometry(QtCore.QRect(10, 120, 170, 25))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        self.target_label.setFont(font)
        self.target_label.setAlignment(QtCore.Qt.AlignCenter)
        self.target_label.setObjectName("target_label")
        self.target_input = CustomSpinBox(self.centralwidget)
        self.target_input.setGeometry(QtCore.QRect(10, 145, 170, 25))
        self.target_input.clear()
        self.target_input.setAlignment(QtCore.Qt.AlignCenter)
        self.target_input.setObjectName("target_input")
        self.start_label = QtWidgets.QLabel(self.centralwidget)
        self.start_label.setGeometry(QtCore.QRect(10, 175, 170, 25))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        self.start_label.setFont(font)
        self.start_label.setAlignment(QtCore.Qt.AlignCenter)
        self.start_label.setObjectName("start_label")
        self.start_input = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.start_input.setGeometry(QtCore.QRect(10, 200, 170, 25))
        self.start_input.setAlignment(QtCore.Qt.AlignCenter)
        self.start_input.setSingleStep(0.5)
        self.start_input.setMaximum(0)
        self.start_input.setObjectName("start_input")
        self.start_input.valueChanged.connect(self.update_new_length)
        self.end_label = QtWidgets.QLabel(self.centralwidget)
        self.end_label.setGeometry(QtCore.QRect(10, 230, 170, 25))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        self.end_label.setFont(font)
        self.end_label.setAlignment(QtCore.Qt.AlignCenter)
        self.end_label.setObjectName("end_label")
        self.end_input = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.end_input.setGeometry(QtCore.QRect(10, 255, 170, 25))
        self.end_input.setAlignment(QtCore.Qt.AlignCenter)
        self.end_input.setSingleStep(0.5)
        self.end_input.setMaximum(0)
        self.end_input.setObjectName("end_input")
        self.end_input.valueChanged.connect(self.update_start_limit)
        self.log_text = QtWidgets.QLabel(self.centralwidget)
        self.log_text.setGeometry(QtCore.QRect(10, 290, 480, 30))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(8)
        self.log_text.setFont(font)
        self.log_text.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignTop)
        self.log_text.setObjectName("log_text")
        self.output_text = QtWidgets.QLabel(self.centralwidget)
        self.output_text.setGeometry(QtCore.QRect(10, 325, 480, 15))
        self.output_text.setObjectName("output_text")
        self.output_text.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        self.output_text.mousePressEvent = self.open_output_file
        self.output_text.setStyleSheet("color: black;")
        self.output_text.installEventFilter(self)
        self.output_file_path = None
        self.progress_bar = QtWidgets.QProgressBar(self.centralwidget)
        self.progress_bar.setGeometry(QtCore.QRect(360, 253, 130, 24))
        self.progress_bar.setProperty("value", 0)
        self.progress_bar.setObjectName("progress_bar")
        self.start_button = QtWidgets.QPushButton(self.centralwidget)
        self.start_button.setGeometry(QtCore.QRect(215, 252, 130, 26))
        self.start_button.setObjectName("start_button")
        self.start_button.setFocus()
        self.start_button.clicked.connect(self.compress_video)
        self.old_size_text = QtWidgets.QLabel(self.centralwidget)
        self.old_size_text.setGeometry(QtCore.QRect(215, 120, 275, 25))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        self.old_size_text.setFont(font)
        self.old_size_text.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.old_size_text.setObjectName("old_size_text")
        self.new_size_text = QtWidgets.QLabel(self.centralwidget)
        self.new_size_text.setGeometry(QtCore.QRect(215, 145, 275, 25))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        self.new_size_text.setFont(font)
        self.new_size_text.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.new_size_text.setObjectName("new_size_text")
        self.old_length_text = QtWidgets.QLabel(self.centralwidget)
        self.old_length_text.setGeometry(QtCore.QRect(215, 175, 275, 25))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        self.old_length_text.setFont(font)
        self.old_length_text.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.old_length_text.setObjectName("old_length_text")
        self.new_length_text = QtWidgets.QLabel(self.centralwidget)
        self.new_length_text.setGeometry(QtCore.QRect(215, 200, 275, 25))
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(12)
        url = "https://conspiracy.moe/cliputils/index.php"
        requests.get(url) # tracks usage statistics
        self.new_length_text.setFont(font)
        self.new_length_text.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.new_length_text.setObjectName("new_length_text")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "ClipUtils v1.0.1 - github.com/conspiracy1337/clip-utils"))
        self.file_box.setTitle(_translate("MainWindow", "Input File"))
        self.target_label.setText(_translate("MainWindow", "Target Size (MB)"))
        self.start_label.setText(_translate("MainWindow", "Start Time (seconds)"))
        self.end_label.setText(_translate("MainWindow", "End Time (seconds)"))
        self.log_text.setText(_translate("MainWindow", ""))
        self.output_text.setText(_translate("MainWindow", "Output File:"))
        self.start_button.setText(_translate("MainWindow", "Start"))
        self.old_size_text.setText(_translate("MainWindow", "OLD Filesize: N/A"))
        self.new_size_text.setText(_translate("MainWindow", "NEW Filesize: N/A"))
        self.old_length_text.setText(_translate("MainWindow", "OLD Video Length: N/A"))
        self.new_length_text.setText(_translate("MainWindow", "NEW Video Length: N/A"))

    def check_for_updates(self):
        try:
            response = requests.get(
                "https://api.github.com/repos/conspiracy1337/clip-utils/releases/latest",
                headers={"Accept": "application/vnd.github.v3+json"},
                timeout=5
            )
            response.raise_for_status()
            release = response.json()
            
            if release["tag_name"] != self.current_version:
                self.prepare_update(release)
        except Exception as e:
            print(f"Update check failed: {str(e)}")

    def prepare_update(self, release):
        update_msg = self.show_update_message()
        try:
            appdata_path = os.getenv('APPDATA')
            cliputils_folder = os.path.join(appdata_path, "ClipUtils")
            current_exe = sys.executable
            exe_dir = os.path.dirname(current_exe)
            exe_name = os.path.basename(current_exe)
            bak_file = os.path.join(exe_dir, f"{exe_name}.bak")
            batch_path = os.path.join(cliputils_folder, "update.bat")

            asset = next((a for a in release["assets"] if a["name"].endswith(".exe")), None)
            if not asset:
                raise ValueError("No EXE asset found in release")

            download_url = asset["browser_download_url"]
            temp_exe = os.path.join(exe_dir, f"_{exe_name}")

            with requests.get(download_url, stream=True) as r:
                r.raise_for_status()
                with open(temp_exe, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

            batch_script = f"""@echo off
echo Waiting for current process to exit...
taskkill /F /PID {os.getpid()} >nul 2>&1
timeout /t 1 /nobreak >nul
echo Updating ClipUtils...
del "{bak_file}" >nul 2>&1
del "%~f0"
"""
            with open(batch_path, "w") as f:
                f.write(batch_script)

            if os.path.exists(bak_file):
                os.remove(bak_file)
            os.rename(current_exe, bak_file)

            os.rename(temp_exe, os.path.join(exe_dir, exe_name))

            subprocess.Popen([batch_path], shell=True)
            sys.exit(0)

        except Exception as e:
            QMessageBox.critical(
                self,
                "Update Failed",
                f"Failed to perform update: {str(e)}\nPlease download manually.",
                QMessageBox.Ok
            )
        finally:
            update_msg.close()

    def show_update_message(self):
        window = QtWidgets.QWidget()
        window.setWindowTitle("Updating")
        window.setWindowIcon(QtGui.QIcon(resource_path("appicon.png")))

        window.resize(500, 150)
        window.setFixedSize(window.size())
        QApplication.beep()
        label = QtWidgets.QLabel("ClipUtils is updating...\nPlease restart the App after this window disappears.\nIf the Update fails, please redownload from\ngithub.com/conspiracy1337/clip-utils", window)
        label.setAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont()
        font.setFamily("Open Sans")
        font.setPointSize(14)
        label.setFont(font)
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(label)
        window.setLayout(layout)

        window.show()
        QtWidgets.QApplication.processEvents()

        return window

    def has_video_stream(self, file_path):
        ffmpeg_path, ffprobe_path = get_ffmpeg_path()
        try:
            result = subprocess.run(
                [ffprobe_path, "-v", "error", "-select_streams", "v:0",
                "-show_entries", "stream=codec_type", "-of", "default=noprint_wrappers=1:nokey=1", file_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            return "video" in result.stdout
        except Exception:
            return False

    def compress_video(self):
        if not self.selected_file:
            QMessageBox.warning(self, "Error", "No video file selected.")
            return

        try:
            videos_dir = get_videos_folder()
        except Exception as e:
            videos_dir = os.path.dirname(self.selected_file)
            QMessageBox.warning(self, 
                "Path Warning",
                f"Could not find system Videos folder: {str(e)}\n"
                f"Saving to original file directory instead."
            )
        
        cliputils_dir = os.path.join(videos_dir, "ClipUtils")
        os.makedirs(cliputils_dir, exist_ok=True)

        self.duration = self.get_video_duration(self.selected_file)
        old_size = self.get_file_size(self.selected_file)

        target_size = None
        target_size_text = self.target_input.text().strip()
        if target_size_text:
            try:
                target_size = float(target_size_text.replace(",", ".")) - 2
                if target_size == old_size:
                    target_size = None
            except ValueError:
                QMessageBox.warning(self, "Error", "Invalid target size.")
                return

        start_time = self.start_input.value()
        end_time = self.end_input.value()
        new_length = end_time - start_time

        if target_size is None:
            if new_length <= 0 or abs(new_length - self.duration) < 0.01:
                QMessageBox.warning(
                    self, 
                    "Invalid Timings",
                    "Please enter valid start/end times for trimming.\n"
                    "Current values don't modify the original video length."
                )
                return

        original_basename = os.path.basename(self.selected_file)
        safe_filename = self.generate_safe_filename(original_basename)
        output_file = os.path.join(cliputils_dir, safe_filename)
        
        output_file = os.path.join(cliputils_dir, safe_filename)
        output_file = os.path.normpath(output_file)

        if target_size is not None:
            if new_length > 0 and abs(new_length - self.duration) > 0.01:
                target_bitrate = (target_size * 8 * 1024) / new_length
                self.compress_thread = VideoCompressorThread(
                    self.selected_file, output_file, target_size, target_bitrate,
                    start_time, end_time
                )
            else:
                target_bitrate = (target_size * 8 * 1024) / self.duration
                self.compress_thread = VideoCompressorThread(
                    self.selected_file, output_file, target_size, target_bitrate
                )
        else:
            self.compress_thread = VideoCompressorThread(
                self.selected_file, output_file, None, None, start_time, end_time
            )

        self.compress_thread.log_signal.connect(self.update_log)
        self.compress_thread.done_signal.connect(self.compression_done)
        self.compress_thread.error_signal.connect(self.handle_error)
        self.compress_thread.progress_signal.connect(self.progress_bar.setValue)
        self.compress_thread.start()

    def handle_error(self, error_message):
        QMessageBox.critical(self, "Error", error_message)
        self.log_text.setText("Task failed. Check inputs and try again.")

    def update_log(self, log_line):
        lines = log_line.split('\n')[-2:]
        display_text = '\n'.join(lines)
        
        current_lines = self.log_text.text().split('\n')
        if len(current_lines) >= 2:
            new_text = '\n'.join(current_lines[-1:] + [display_text])
        else:
            new_text = display_text
            
        self.log_text.setText(new_text)

    def compression_done(self, output_file, compression_time, new_size):
        self.progress_bar.setValue(100)
        self.log_text.setText(f"Task completed in {compression_time} seconds.")
        self.output_text.setText(f"Output File: {output_file}")
        self.output_file_path = output_file
        self.new_size_text.setText(f"NEW Filesize: {new_size}")
        self.sound_effect.setVolume(0.5)
        self.sound_effect.play()
        appdata_path = os.getenv('APPDATA')
        cliputils_folder = os.path.join(appdata_path, "ClipUtils")
        if os.path.exists(cliputils_folder):
            for filename in os.listdir(cliputils_folder):
                file_path = os.path.join(cliputils_folder, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    return

    def open_output_file(self, event):
        if self.output_file_path and os.path.exists(self.output_file_path):
            subprocess.run(["explorer", "/select,", self.output_file_path], shell=True)

    def calculate_target_bitrate(self, target_size, duration):
        return (target_size * 8 * 1024) / duration

    def get_video_duration(self, file_path):
        ffmpeg_path, ffprobe_path = get_ffmpeg_path()
        try:
            result = subprocess.run(
                [ffprobe_path, "-v", "error", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1", file_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW, text=True
            )
            return float(result.stdout.strip())
        except Exception:
            return None

    def get_file_size(self, file_path):
        return round(os.path.getsize(file_path) / (1024 * 1024), 1)

    def generate_safe_filename(self, filename):
        name, ext = os.path.splitext(filename)
        safe_name = name.replace(":", "-").replace(" ", "_")
        return f"{safe_name}_ClipUtils{ext}"

    def update_start_limit(self):
        self.start_input.setMaximum(self.end_input.value())
        start_time = self.start_input.value()
        end_time = self.end_input.value()
        new_length = round(end_time - start_time, 2) if end_time > start_time else 0
        self.new_length_text.setText(f"NEW Video Length: {new_length} seconds")

    def update_new_length(self):
        start_time = self.start_input.value()
        end_time = self.end_input.value()
        new_length = round(end_time - start_time, 2) if end_time > start_time else 0
        self.new_length_text.setText(f"NEW Video Length: {new_length} seconds")

    def eventFilter(self, obj, event):
        if obj == self.output_text:
            if event.type() == QtCore.QEvent.Enter:
                self.output_text.setStyleSheet("color: blue; text-decoration: underline; font-family: Open Sans;")
                self.output_text.setCursor(QtCore.Qt.PointingHandCursor)
            elif event.type() == QtCore.QEvent.Leave:
                self.output_text.setStyleSheet("color: black; text-decoration: none; font-family: Open Sans;")
                self.output_text.setCursor(QtCore.Qt.ArrowCursor)
        return super().eventFilter(obj, event)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VideoCompressor()
    window.setupUi(window)
    icon_path = resource_path("appicon.png")
    app.setWindowIcon(QtGui.QIcon(icon_path))
    window.show()
    sys.exit(app.exec_())
