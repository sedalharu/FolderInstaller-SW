# 필수 시스템 모듈
import sys
import os
import time
import subprocess
import ctypes

# PyQt5 관련 모듈
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QProgressBar, QScrollArea, QFileDialog,
    QMessageBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QFontMetrics
from PyQt5.QtGui import QFontDatabase


def resource_path(relative_path):
    """Get absolute path to resource"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class InstallerThread(QThread):
    progress_update = pyqtSignal(str, int, str)  # 파일명, 진행률, 상태 메시지
    installation_complete = pyqtSignal(bool)  # 설치 성공 여부

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def is_process_running(self, process):
        """프로세스가 현재 실행 중인지 확인"""
        try:
            return process.poll() is None  # None이면 프로세스가 아직 실행 중
        except:
            return False

    def run(self):
        """설치 프로세스 실행 및 모니터링"""
        try:
            # 설치 시작 알림
            self.progress_update.emit(os.path.basename(self.file_path), 0, "설치 진행중...")

            # 설치 프로세스 시작
            process = subprocess.Popen(
                self.file_path,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )

            # 진행 상태 추적 변수
            progress = 0
            previous_time = time.time()

            # 설치 진행 상태 모니터링
            while self.is_process_running(process):
                current_time = time.time()

                # 출력 확인 (실시간 로그 모니터링 방식으로 변경)
                if process.stdout:
                    output = process.stdout.readline()
                    if "completed" in output.lower():  # 로그에서 완료 상태 탐지
                        progress = 90  # 90%로 설정

                # 진행 상태 업데이트
                if progress < 90:
                    if current_time - previous_time >= 1.0:  # 1.0초마다 진행률 업데이트
                        progress += 1
                        self.progress_update.emit(
                            os.path.basename(self.file_path),
                            int(progress),
                            "설치 진행중..."
                        )
                        previous_time = current_time

                self.msleep(100)

            # 설치 프로세스 완료 대기
            process.wait()

            # 설치 결과 확인 및 최종 상태 업데이트
            if process.returncode == 0:  # 정상 종료
                # 90%에서 100%까지 천천히 증가
                for i in range(int(progress), 101):
                    self.progress_update.emit(
                        os.path.basename(self.file_path),
                        i,
                        "설치 진행중..."
                    )
                    self.msleep(50)

                self.progress_update.emit(
                    os.path.basename(self.file_path),
                    100,
                    "설치 완료"
                )
                self.installation_complete.emit(True)
            else:  # 비정상 종료
                self.progress_update.emit(
                    os.path.basename(self.file_path),
                    0,
                    "설치 실패"
                )
                self.installation_complete.emit(False)

        except Exception as e:
            # 예외 발생 시 설치 실패로 처리
            self.progress_update.emit(
                os.path.basename(self.file_path),
                0,
                f"설치 실패: {str(e)}"
            )
            self.installation_complete.emit(False)


class InstallProgressWidget(QWidget):
    """개별 설치 프로그램 진행 상태 표시"""

    def __init__(self, file_name):
        super().__init__()

        # 레이아웃 설정
        layout = QHBoxLayout()
        layout.setContentsMargins(10, 5, 10, 5)

        # 파일명 레이블 설정
        self.file_label = QLabel(file_name)
        self.file_label.setFont(QFont(FolderInstaller.font_family, 10))
        self.file_label.setFixedWidth(200)
        self.file_label.setToolTip(file_name)

        # 긴 파일명 처리 (말줄임표 추가)
        metrics = QFontMetrics(self.file_label.font())
        elidedText = metrics.elidedText(file_name, Qt.ElideMiddle, self.file_label.width())
        self.file_label.setText(elidedText)

        # 프로그레스 바 설정
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumWidth(400)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #2196F3;
                border-radius: 5px;
                text-align: center;
                min-width: 400px;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
            }
        """)

        # 상태 레이블 설정
        self.status_label = QLabel("대기중")
        self.status_label.setFont(QFont(FolderInstaller.font_family, 10))
        self.status_label.setFixedWidth(100)

        # 위젯 배치
        layout.addWidget(self.file_label)
        layout.addWidget(self.progress_bar, 1)
        layout.addWidget(self.status_label)

        self.setLayout(layout)


class FolderInstaller(QMainWindow):
    """메인 창 클래스"""
    font_family = "Arial"  # 기본 폰트

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        """UI 초기화 및 기본 설정"""
        # 관리자 권한 확인
        if not ctypes.windll.shell32.IsUserAnAdmin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()

        # 폰트 로드
        font_path = resource_path('SCDream4.otf')
        font_id = QFontDatabase.addApplicationFont(font_path)
        if font_id >= 0:
            font_families = QFontDatabase.applicationFontFamilies(font_id)
            FolderInstaller.font_family = font_families[0]

        # 창 기본 설정
        self.setWindowTitle('Folder Installer')
        self.setFixedSize(800, 600)

        # 윈도우 아이콘 설정
        icon_path = resource_path('app.ico')
        self.setWindowIcon(QIcon(icon_path))

        # 중앙 위젯 생성
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 메인 레이아웃 설정
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setAlignment(Qt.AlignCenter)

        # 상단 여백 추가
        self.main_layout.addSpacing(50)

        # 타이틀 레이블 설정
        title_label = QLabel('Folder Installer')
        title_label.setFont(QFont(FolderInstaller.font_family, 26, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)

        # 설명 텍스트 추가
        description_label = QLabel('선택한 폴더 내의 모든 설치 프로그램을 자동으로 설치합니다.')
        description_label.setFont(QFont(FolderInstaller.font_family, 12))
        description_label.setAlignment(Qt.AlignCenter)

        # 지원 형식 레이블 설정
        format_label = QLabel('< 지원형식: .exe .msi >')
        format_label.setFont(QFont(FolderInstaller.font_family, 10))
        format_label.setAlignment(Qt.AlignCenter)

        # 폴더 선택 버튼 설정
        self.select_button = QPushButton('설치 폴더 선택')
        self.select_button.setFont(QFont(FolderInstaller.font_family, 12, QFont.Bold))
        self.select_button.setFixedSize(200, 50)
        self.select_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)

        # 스크롤 영역 설정
        self.scroll_area = QScrollArea()
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_area.setWidget(self.scroll_widget)
        self.scroll_area.setWidgetResizable(True)

        # 레이아웃 위젯 추가
        self.main_layout.addWidget(title_label)
        self.main_layout.addSpacing(40)  # 타이틀과 설명 텍스트 사이 간격
        self.main_layout.addWidget(description_label)
        self.main_layout.addSpacing(5)  # 설명 텍스트와 지원 형식 사이 간격
        self.main_layout.addWidget(format_label)
        self.main_layout.addSpacing(30)  # 지원 형식 레이블과 버튼 사이 간격
        self.main_layout.addWidget(self.select_button, alignment=Qt.AlignCenter)
        self.main_layout.addSpacing(40)  # 하단 여백 추가

        # 버튼 클릭 시그널 연결
        self.select_button.clicked.connect(self.select_folder)

        # 설치 관련 변수 초기화
        self.installer_threads = []
        self.installation_results = {'success': 0, 'fail': 0}
        self.total_installations = 0

    def select_folder(self):
        """폴더 선택 및 설치 프로세스 시작"""
        folder_path = QFileDialog.getExistingDirectory(self, '폴더 선택')
        if folder_path:
            # 버튼 숨기기
            self.select_button.hide()

            # 스크롤 영역 추가
            self.main_layout.addWidget(self.scroll_area)

            # 기존 설치 정보 초기화
            for i in reversed(range(self.scroll_layout.count())):
                self.scroll_layout.itemAt(i).widget().setParent(None)
            self.installer_threads.clear()

            # 설치 결과 초기화
            self.installation_results = {'success': 0, 'fail': 0}

            # 설치 파일 찾기
            install_files = []
            for file in os.listdir(folder_path):
                if file.lower().endswith(('.exe', '.msi')):
                    install_files.append(os.path.join(folder_path, file))

            if install_files:
                self.total_installations = len(install_files)

                # 각 설치 파일에 대한 진행 위젯 생성 및 설치 시작
                for file_path in install_files:
                    progress_widget = InstallProgressWidget(os.path.basename(file_path))
                    self.scroll_layout.addWidget(progress_widget)

                    installer_thread = InstallerThread(file_path)
                    installer_thread.progress_update.connect(
                        lambda fn, p, s, w=progress_widget: self.update_progress(w, fn, p, s)
                    )
                    installer_thread.installation_complete.connect(self.check_installation_result)
                    self.installer_threads.append(installer_thread)
                    installer_thread.start()

    def check_installation_result(self, success):
        """개별 설치 완료 결과 확인"""
        if success:
            self.installation_results['success'] += 1
        else:
            self.installation_results['fail'] += 1

        # 모든 설치가 완료되었는지 확인
        completed = self.installation_results['success'] + self.installation_results['fail']
        if completed == self.total_installations:
            msg = QMessageBox()
            msg.setWindowTitle("설치 완료")
            msg.setText(f"설치 성공: {self.installation_results['success']}개\n"
                        f"설치 실패: {self.installation_results['fail']}개")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.buttonClicked.connect(self.close)

            # 레이아웃 마진 조정
            layout = msg.layout()
            layout.setContentsMargins(20, 10, 30, 10)  # 좌, 상, 우, 하 여백

            msg.exec_()

    def update_progress(self, widget, file_name, progress, status):
        """설치 진행 상태 업데이트"""
        widget.progress_bar.setValue(progress)
        widget.status_label.setText(status)


def main():
    """프로그램 메인 함수"""
    # 콘솔창 숨기기
    kernel32 = ctypes.WinDLL('kernel32')
    user32 = ctypes.WinDLL('user32')
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, 0)

    # 애플리케이션 실행
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # 앱 아이콘 설정
    app_icon = QIcon(resource_path('app.ico'))
    app.setWindowIcon(app_icon)

    window = FolderInstaller()
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()