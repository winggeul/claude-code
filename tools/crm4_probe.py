"""통합고객목록 창의 Win32 컨트롤을 조사한다.

UIA 로는 이 프로그램의 컨트롤이 전부 밋밋한 Pane 으로만 보였지만, 클래스 이름이
`WindowsForms10.EDIT` / `.BUTTON` / `.COMBOBOX` 인 것에서 알 수 있듯 실체는 전부
진짜 Win32 자식 윈도우다. 즉 win32 백엔드로는 핸들을 직접 잡아 값을 넣고 누를 수 있다.

읽지 못하는 것은 격자 하나뿐이고, 그건 클립보드로 우회한다.

이 스크립트는 자동화 코드가 집어야 할 컨트롤(등록일 입력칸, 검색 버튼, 탭, 전체선택)의
정확한 식별자를 얻기 위한 조사용이다.

실행:
    pip install -r tools/requirements.txt
    python tools/crm4_probe.py

통합고객목록 창을 띄운 상태에서 실행할 것. 탭을 바꿔가며 두세 번 돌리면 각 탭의
컨트롤이 모두 잡힌다.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

WINDOW_TITLE = "통합고객목록"


def dump_window(win) -> list[str]:
    """창 하나의 자식 컨트롤을 전부 한 줄씩 기록한다."""
    lines = [f"===== {win.window_text()!r} (handle={win.handle}) ====="]

    for ctrl in win.descendants():
        try:
            rect = ctrl.rectangle()
            text = ctrl.window_text()
            lines.append(
                " | ".join(
                    [
                        f"handle={ctrl.handle}",
                        f"ctrl_id={ctrl.control_id()}",
                        f"class={ctrl.class_name()}",
                        f"text={text[:50]!r}",
                        f"visible={ctrl.is_visible()}",
                        f"enabled={ctrl.is_enabled()}",
                        f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom})",
                    ]
                )
            )
        except Exception as exc:
            lines.append(f"(컨트롤 조회 실패: {exc})")

    return lines


def main() -> int:
    parser = argparse.ArgumentParser(description="통합고객목록 Win32 컨트롤 조사")
    parser.add_argument("--title", default=WINDOW_TITLE, help="대상 창 제목")
    parser.add_argument("--out", default=None, help="저장 경로")
    args = parser.parse_args()

    try:
        from pywinauto import Desktop
    except ImportError:
        print("pywinauto 가 없습니다. 먼저 실행하세요: pip install -r tools/requirements.txt")
        return 1

    windows = [
        w
        for w in Desktop(backend="win32").windows()
        if args.title in (w.window_text() or "")
    ]
    if not windows:
        print(f"'{args.title}' 창을 찾지 못했습니다. 창을 띄운 상태에서 실행하세요.")
        return 1

    out = [f"# 통합고객목록 Win32 덤프 - {datetime.now():%Y-%m-%d %H:%M:%S}", ""]
    for win in windows:
        out.extend(dump_window(win))
        out.append("")

    path = Path(args.out or f"crm4_probe_{datetime.now():%Y%m%d_%H%M%S}.txt")
    path.write_text("\n".join(out), encoding="utf-8")
    print(f"저장 완료: {path.resolve()}  ({len(out)} 줄)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
