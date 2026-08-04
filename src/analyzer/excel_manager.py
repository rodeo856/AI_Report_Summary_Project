import os
import pythoncom
import win32com.client as win32
import shutil

from analyzer.options import AnalyzeOptions
from analyzer.report_parser import ReportParser
from analyzer.summary_generator import SummaryGenerator
from pathlib import Path

def make_unique_status_path(file_path: str) -> str:
    src = Path(file_path)
    base = src.with_name(src.stem + "_Status.xlsx")

    if not base.exists():
        return str(base)

    idx = 1
    while True:
        candidate = src.with_name(f"{src.stem}_Status_{idx:03d}.xlsx")
        if not candidate.exists():
            return str(candidate)
        idx += 1

class ExcelManager:
    def __init__(self, log_func=None, options: AnalyzeOptions | None = None):
        self.log = log_func or (lambda msg: None)
        self.options = options or AnalyzeOptions()
        self.excel = None
        self.workbook = None

    def analyze_and_write_summary(self, file_path: str):
        print("[ExcelManager] 함수 진입", flush=True)
        output_path = None

        file_path = os.path.abspath(file_path)
        if os.path.basename(file_path).startswith("~$"):
            raise ValueError("Excel 임시 잠금 파일(~$...)이 선택되었습니다. 원본 AER 파일을 선택하세요.")
        if not os.path.exists(file_path):
            raise FileNotFoundError(file_path)

        self.log("[ExcelManager] COM 초기화 시작")
        pythoncom.CoInitialize()

        try:
            self.log("[ExcelManager] Excel Application 생성 전")
            self.excel = win32.DispatchEx("Excel.Application")
            self.log("[ExcelManager] Excel Application 생성 완료")

            # 디버깅 중에는 True 권장. 안정화 후 False로 바꿔도 됨.
            self.excel.Visible = True
            self.excel.DisplayAlerts = True

            self.log("[ExcelManager] Excel 옵션 설정 완료")

            output_path = make_unique_status_path(file_path)
            self.log(f"[ExcelManager] Status 복사본 경로 생성: {output_path}")

            self.log("[ExcelManager] 원본 파일 복사 시작")
            shutil.copy2(file_path, output_path)
            self.log("[ExcelManager] 원본 파일 복사 완료")

            self.log(f"[ExcelManager] Workbook.Open 시작: {output_path}")
            self.workbook = self.excel.Workbooks.Open(
                Filename=output_path,
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
                Editable=True,
            )
            self.log("[ExcelManager] Workbook.Open 완료")
            self.log(f"[ExcelManager] ReadOnly={self.workbook.ReadOnly}")
            self.log(f"[ExcelManager] FullName={self.workbook.FullName}")

            self.log("[ExcelManager] ReportParser 생성")
            parser = ReportParser(self.log)

            self.log("[ExcelManager] Workbook 파싱 시작")
            results = parser.parse_workbook(self.workbook)
            self.log(f"[ExcelManager] Workbook 파싱 완료 / 분석 항목 수: {len(results)}")

            self.log("[ExcelManager] Sheet Generator 생성")
            generator = SummaryGenerator(self.log, self.options)

            self.log("[ExcelManager] Sheet 생성 시작")
            generator.generate(self.workbook, results)
            self.log("[ExcelManager] Sheet 생성 완료")

            self.log("[ExcelManager] Workbook 저장 시작")

            self.excel.DisplayAlerts = False
            self.workbook.Save()
            self.workbook.Saved = True

            self.log("[ExcelManager] Workbook 저장 완료")

        finally:
            self.log("[ExcelManager] 종료 처리 시작")
            if self.workbook is not None:
                self.log("[ExcelManager] Workbook 닫기")
                self.workbook.Saved = True
                self.workbook.Close(SaveChanges=False)
                self.workbook = None
            if self.excel is not None:
                self.log("[ExcelManager] Excel 종료")
                self.excel.Quit()
                self.excel = None
            pythoncom.CoUninitialize()
            self.log("[ExcelManager] COM 해제 완료")

        return output_path
