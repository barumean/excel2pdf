"""
Excel → PDF 일괄 변환기  진입점.

실행 전 검사:
  1. Microsoft Excel (COM) 설치 여부
  2. tkinterdnd2 가용 여부 (없으면 DnD 기능만 비활성화)

사용법:
  python main.py
  또는 PyInstaller 로 빌드한 excel2pdf.exe 실행
"""

import sys
import tkinter as tk
from tkinter import messagebox


def _check_excel() -> bool:
    """Excel COM 서버 가용 여부 확인."""
    try:
        import win32com.client  # noqa: F401
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Quit()
        del excel
        return True
    except ImportError:
        return False
    except Exception:
        return False


def main():
    # 먼저 최소한의 루트 윈도우를 생성해 messagebox 를 표시할 수 있도록 한다.
    root = _build_root()

    if not _check_excel():
        messagebox.showerror(
            "Microsoft Excel 없음",
            "이 프로그램은 Microsoft Excel 이 설치된 Windows 환경에서만 실행됩니다.\n\n"
            "Excel 을 설치한 후 다시 시도해 주세요.",
        )
        root.destroy()
        sys.exit(1)

    from app.gui import App  # Excel 확인 후 임포트
    app = App(root)
    root.mainloop()


def _build_root() -> tk.Tk:
    """tkinterdnd2 가 있으면 DnD 지원 루트, 없으면 일반 루트를 반환."""
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
    return root


if __name__ == "__main__":
    main()
