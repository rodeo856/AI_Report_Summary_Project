import sys
import traceback
from PySide6.QtWidgets import QApplication, QMessageBox
from src.ui.main_window import MainWindow


def main():
    try:
        print("[MAIN] QApplication 생성", flush=True)
        app = QApplication(sys.argv)

        print("[MAIN] MainWindow 생성", flush=True)
        window = MainWindow()

        print("[MAIN] MainWindow 표시", flush=True)
        window.show()

        print("[MAIN] Event Loop 시작", flush=True)
        sys.exit(app.exec())

    except Exception as e:
        print("[MAIN ERROR]", flush=True)
        print(traceback.format_exc(), flush=True)
        try:
            QMessageBox.critical(None, "실행 오류", str(e))
        except Exception:
            pass


if __name__ == "__main__":
    main()
