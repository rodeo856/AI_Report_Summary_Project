import os
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import QFile, Qt
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QMainWindow, QFileDialog, QMessageBox, QApplication

from analyzer.excel_manager import ExcelManager
from analyzer.options import AnalyzeOptions


class DropAreaEventFilter:
    def __init__(self, parent):
        self.parent = parent

    def eventFilter(self, obj, event):
        if event.type() == event.Type.DragEnter:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True
        if event.type() == event.Type.Drop:
            urls = event.mimeData().urls()
            if urls:
                self.parent.set_file(urls[0].toLocalFile())
                event.acceptProposedAction()
                return True
        return False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_file = ""
        self.generated_file = ""
        self._drop_filter = DropAreaEventFilter(self)
        loaded_ui = self._load_ui()
        if isinstance(loaded_ui, QMainWindow):
            self.ui = loaded_ui
            self.setCentralWidget(loaded_ui.centralWidget())
        else:
            self.ui = loaded_ui
            self.setCentralWidget(self.ui)
        self.setWindowTitle("Progress Checker")
        self.resize(900, 650)

        try:
            self.ui.label_title.setText("Progress Checker")
        except Exception:
            pass

        try:
            self.ui.label_dropArea.setText("Report 파일을 여기에 Drag & Drop 하세요")
        except Exception:
            pass

        # 기존 Status 삭제 체크박스만 남기고 나머지 옵션 제거
        for name in [
            "chk_waveformCheck",
            "chk_dailyProgress",
            "chk_dataTableCheck",
            "chk_imageExistCheck",
            "chk_waveform",
            "chk_image_exist",
        ]:
            if hasattr(self.ui, name):
                getattr(self.ui, name).hide()

        try:
            self.ui.chk_recreateSummary.setText("기존 Status 시트 삭제")
        except Exception:
            pass

        # -----------------------------------
        # 체크박스 숨기기
        # -----------------------------------

        try:
            self.ui.chk_waveform.hide()
        except:
            pass

        try:
            self.ui.chk_image_exist.hide()
        except:
            pass

        # -----------------------------------
        self._connect_signals()
        self._setup_drag_drop()
        self.log("대기 중...")

    def _resource_path(self, relative_path: str) -> str:
        base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])
        return str(Path(base_path) / relative_path)

    def _load_ui(self):
        ui_path = self._resource_path("ui/main.ui")

        print(f"UI PATH = {ui_path}")
        print(f"EXISTS = {os.path.exists(ui_path)}")

        ui_file = QFile(ui_path)
        if not ui_file.open(QFile.ReadOnly):
            raise FileNotFoundError(f"UI 파일을 열 수 없습니다: {ui_path}")

        loader = QUiLoader()
        loaded = loader.load(ui_file)
        ui_file.close()

        if loaded is None:
            raise RuntimeError("UI 로드 실패")

        return loaded

    def _connect_signals(self):
        self.ui.btn_selectFile.clicked.connect(self.select_file)
        self.ui.btn_createSummary.clicked.connect(self.create_summary)
        self.ui.btn_openExcel.clicked.connect(self.open_excel_file)

    def _setup_drag_drop(self):
        self.ui.label_dropArea.setAcceptDrops(False)
        #self.ui.label_dropArea.installEventFilter(self._drop_filter)

    def _get_options(self):
        return AnalyzeOptions(
            recreate_summary=self.ui.chk_recreateSummary.isChecked(),
            waveform_check=True,
            daily_progress=True,
            data_table_check=True,
            image_exist_check=True,
        )

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Report 선택",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)",
        )
        if file_path:
            self.set_file(file_path)

    def set_file(self, file_path: str):
        file_path = file_path.strip()
        if not file_path.lower().endswith((".xlsx", ".xlsm", ".xls")):
            QMessageBox.warning(self, "확인", "Excel 파일만 선택할 수 있습니다.")
            return
        if os.path.basename(file_path).startswith("~$"):
            QMessageBox.warning(self, "확인", "Excel 임시 잠금 파일(~$...)은 선택할 수 없습니다.\n원본 AER 파일을 선택하세요.")
            return
        self.selected_file = file_path
        self.ui.lineEdit_filePath.setText(file_path)
        self.log(f"파일 선택: {file_path}")

    def create_summary(self):
        file_path = self.ui.lineEdit_filePath.text().strip()
        if not file_path:
            QMessageBox.warning(self, "확인", "분석할 리포트 파일을 먼저 선택하세요.")
            return
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "확인", "파일 경로를 찾을 수 없습니다.")
            return
        if os.path.basename(file_path).startswith("~$"):
            QMessageBox.warning(self, "확인", "Excel 임시 잠금 파일(~$...)은 선택할 수 없습니다.\n원본 AER 파일을 선택하세요.")
            return

        ok = QMessageBox.question(
            self,
            "Summary 생성",
            "선택한 리포트의 복사본에 Status 시트를 생성할까요?\n\n"
            "기존 Status 시트가 있으면 삭제 후 재생성됩니다.",
        )
        if ok != QMessageBox.Yes:
            return

        self.ui.progressBar.setValue(0)
        self.ui.textEdit_log.clear()
        self.log("분석 시작")
        self.ui.btn_createSummary.setEnabled(False)

        try:
            self.ui.progressBar.setValue(10)
            self.log("[10%] Excel Manager 생성")
            self.ui.progressBar.repaint()
            manager = ExcelManager(self.log, self._get_options())

            self.log("[10%] Excel 분석 함수 진입")
            self.ui.progressBar.repaint()
            self.generated_file = manager.analyze_and_write_summary(file_path)

            self.log("[90%] Excel 분석 함수 종료")
            self.ui.progressBar.setValue(90)
            self.ui.progressBar.repaint()

            self.log("[100%] Status 생성 완료")
            self.ui.progressBar.setValue(100)
            self.ui.progressBar.repaint()
            QMessageBox.information(self, "완료", "Status 시트 생성이 완료되었습니다.")

        except Exception as exc:
            self.ui.progressBar.setValue(100)
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "오류", str(exc))

        finally:
            self.ui.btn_createSummary.setEnabled(True)

    def open_excel_file(self):
        file_path = self.generated_file or self.ui.lineEdit_filePath.text().strip()

        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "확인", "열 파일이 없습니다.")
            return

        os.startfile(file_path)

    def log(self, message: str):
        self.ui.textEdit_log.append(str(message))
        self.ui.textEdit_log.verticalScrollBar().setValue(
            self.ui.textEdit_log.verticalScrollBar().maximum()
        )
        QApplication.processEvents()
