"""
Excel → PDF 변환 핵심 로직.
win32com.client 을 통해 Excel COM 객체를 제어한다.
UI 스레드와 분리되어 별도 스레드에서 실행된다.
"""

import os
import threading
import time
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Callable, List, Optional

EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb"}

# ExportAsFixedFormat 에 전달할 PDF 포맷 상수
XL_TYPE_PDF = 0

# PrintArea 없이 전체 시트 인쇄 시 xlWorksheet
XL_PAPER_A4 = 9
XL_PORTRAIT = 1
XL_LANDSCAPE = 2


class SheetScope(Enum):
    ALL = "all"        # 모든 시트
    ACTIVE = "active"  # 활성(첫번째) 시트만


@dataclass
class ConversionOptions:
    sheet_scope: SheetScope = SheetScope.ALL
    fit_to_page: bool = True           # 열 너비를 한 페이지에 맞춤
    include_subfolders: bool = True
    output_same_folder: bool = True    # 원본 파일과 같은 폴더에 저장
    output_folder: str = ""            # output_same_folder=False 일 때 사용
    skip_hidden_sheets: bool = True


@dataclass
class ConversionResult:
    path: str
    success: bool
    message: str = ""
    duration: float = 0.0


@dataclass
class ConversionProgress:
    total: int = 0
    done: int = 0
    current_file: str = ""
    results: List[ConversionResult] = field(default_factory=list)

    @property
    def percent(self) -> float:
        return (self.done / self.total * 100) if self.total else 0.0


class Converter:
    """
    엑셀 파일을 PDF 로 변환하는 워커.
    run() 을 별도 스레드에서 호출한다.

    콜백:
        on_progress(progress: ConversionProgress) - 파일 1개 완료 시마다 호출
        on_finished(progress: ConversionProgress) - 전체 완료 시 호출
    """

    def __init__(
        self,
        files: List[str],
        options: ConversionOptions,
        on_progress: Optional[Callable] = None,
        on_finished: Optional[Callable] = None,
    ):
        self._files = files
        self._options = options
        self._on_progress = on_progress
        self._on_finished = on_finished
        self._stop_event = threading.Event()
        self._progress = ConversionProgress(total=len(files))

    def stop(self):
        self._stop_event.set()

    @property
    def is_running(self) -> bool:
        return not self._stop_event.is_set()

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def run(self):
        excel = None
        try:
            excel = self._create_excel_app()
            for file_path in self._files:
                if self._stop_event.is_set():
                    break
                self._progress.current_file = file_path
                result = self._convert_file(excel, file_path)
                self._progress.done += 1
                self._progress.results.append(result)
                if self._on_progress:
                    self._on_progress(self._progress)
        finally:
            self._quit_excel(excel)
            if self._on_finished:
                self._on_finished(self._progress)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _create_excel_app(self):
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        return excel

    def _quit_excel(self, excel):
        if excel is None:
            return
        try:
            excel.Quit()
        except Exception:
            pass
        # COM 레퍼런스 해제
        try:
            import win32com.client
            del excel
        except Exception:
            pass
        # 혹시 남은 프로세스 정리
        try:
            import gc
            gc.collect()
        except Exception:
            pass

    def _convert_file(self, excel, file_path: str) -> ConversionResult:
        t0 = time.monotonic()
        wb = None
        try:
            wb = excel.Workbooks.Open(
                file_path,
                UpdateLinks=False,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )

            pdf_path = self._build_pdf_path(file_path)

            if self._options.sheet_scope == SheetScope.ALL:
                self._export_all_sheets(excel, wb, pdf_path)
            else:
                self._export_active_sheet(excel, wb, pdf_path)

            duration = time.monotonic() - t0
            return ConversionResult(path=file_path, success=True, duration=duration)

        except Exception as exc:
            duration = time.monotonic() - t0
            return ConversionResult(
                path=file_path,
                success=False,
                message=str(exc),
                duration=duration,
            )
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def _build_pdf_path(self, excel_path: str) -> str:
        src = Path(excel_path)
        if self._options.output_same_folder:
            dest_dir = src.parent
        else:
            dest_dir = Path(self._options.output_folder)
            dest_dir.mkdir(parents=True, exist_ok=True)
        return str(dest_dir / (src.stem + ".pdf"))

    def _apply_fit_to_page(self, sheet):
        """한 페이지에 모든 열이 들어오도록 설정."""
        try:
            ps = sheet.PageSetup
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False  # 세로는 제한 없음
        except Exception:
            pass

    def _export_all_sheets(self, excel, wb, pdf_path: str):
        """
        모든 (인쇄 가능한) 시트를 선택한 뒤 하나의 PDF 로 내보낸다.
        빈 시트나 차트 시트는 건너뛴다.
        """
        import win32com.client
        printable = []
        for sh in wb.Sheets:
            # xlChart=3, xlWorksheet=1, xlDialogSheet=5 등
            if sh.Type != 1:  # xlWorksheet 만 처리
                continue
            if self._options.skip_hidden_sheets and sh.Visible != -1:  # xlSheetVisible=-1
                continue
            if self._options.fit_to_page:
                self._apply_fit_to_page(sh)
            printable.append(sh.Name)

        if not printable:
            raise RuntimeError("변환할 수 있는 워크시트가 없습니다.")

        # 여러 시트를 하나의 PDF 로 내보내려면 시트들을 동시에 선택 후 Export
        wb.Sheets(printable[0]).Select(Replace=True)
        for name in printable[1:]:
            wb.Sheets(name).Select(Replace=False)

        wb.ActiveSheet.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=pdf_path,
            Quality=0,          # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

    def _export_active_sheet(self, excel, wb, pdf_path: str):
        """활성(첫 번째 표시) 시트만 PDF 로 내보낸다."""
        sh = wb.ActiveSheet
        if self._options.fit_to_page:
            self._apply_fit_to_page(sh)
        sh.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=pdf_path,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )


# ------------------------------------------------------------------
# 파일 수집 유틸
# ------------------------------------------------------------------

def collect_excel_files(folder: str, recursive: bool) -> List[str]:
    """folder 안에서 엑셀 파일 목록을 수집한다."""
    result = []
    folder_path = Path(folder)
    pattern = "**/*" if recursive else "*"
    for p in sorted(folder_path.glob(pattern)):
        if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS:
            # 임시 파일(~$ 로 시작)은 제외
            if not p.name.startswith("~$"):
                result.append(str(p))
    return result


def is_excel_installed() -> bool:
    """Microsoft Excel COM 서버가 등록되어 있는지 확인."""
    try:
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Quit()
        del excel
        return True
    except Exception:
        return False
