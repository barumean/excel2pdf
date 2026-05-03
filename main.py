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

from app.converter import is_excel_installed


def main():
    root = _build_root()
    # Excel 확인이 완료되기 전까지 빈 창이 보이지 않도록 숨김
    root.withdraw()

    if not is_excel_installed():
        messagebox.showerror(
            "Microsoft Excel 없음",
            "이 프로그램은 Microsoft Excel 이 설치된 Windows 환경에서만 실행됩니다.\n\n"
            "Excel 을 설치한 후 다시 시도해 주세요.",
        )
        root.destroy()
        sys.exit(1)

    from app.gui import App
    App(root)
    root.deiconify()  # 준비 완료 후 창 표시
    root.mainloop()


def _build_root() -> tk.Tk:
    """tkinterdnd2 가 있으면 DnD 지원 루트, 없으면 일반 루트를 반환."""
    try:
        from tkinterdnd2 import TkinterDnD
        return TkinterDnD.Tk()
    except ImportError:
        return tk.Tk()


if __name__ == "__main__":
    main()
