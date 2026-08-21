"""CRM4enterprise 창의 UI 트리를 덤프한다.

CRM4enterprise 는 데이터를 로컬에 저장하지 않고 원격 서버(OtOSolution)에서 받아
화면에만 뿌린다. 따라서 자동 수집의 진입점은 파일이나 DB 가 아니라 화면 컨트롤이다.
이 스크립트는 그 화면을 자동화하기 전에, 어떤 컨트롤을 어떤 이름으로 집어야 하는지
확인하기 위한 1 회성 조사 도구다.

실행:
    pip install -r tools/requirements.txt
    python tools/crm4_inspect.py                     # 전체 트리를 파일로 덤프
    python tools/crm4_inspect.py --max-depth 4       # 얕게만
    python tools/crm4_inspect.py --grid-rows 5       # 표 컨트롤 미리보기 행 수

CRM4enterprise 를 실행하고 데이터가 보이는 화면(고객목록 등)을 띄운 상태에서 돌릴 것.
"""

from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime
from pathlib import Path

PROCESS_NAME = "CRM4enterprise.exe"


def iter_windows(process_name: str):
    """지정한 프로세스가 소유한 최상위 창을 모두 돌려준다."""
    from pywinauto import Desktop

    for win in Desktop(backend="uia").windows():
        try:
            if win.element_info.process_id and _exe_name(win.element_info.process_id) == process_name:
                yield win
        except Exception:
            continue


def _exe_name(pid: int) -> str:
    import psutil

    try:
        return psutil.Process(pid).name()
    except Exception:
        return ""


def describe(ctrl, depth: int) -> str:
    """컨트롤 한 개를 한 줄로 요약한다."""
    info = ctrl.element_info
    rect = info.rectangle
    parts = [
        f"{'  ' * depth}[{depth}]",
        f"type={info.control_type}",
        f"class={info.class_name or '-'}",
        f"auto_id={info.automation_id or '-'}",
        f"name={(info.name or '-')[:60]!r}",
        f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom})",
    ]
    return " ".join(parts)


def grid_preview(ctrl, max_rows: int) -> list[str]:
    """Grid 패턴을 지원하는 컨트롤이면 크기와 앞부분 몇 행을 미리 보여준다.

    WinForms 의 DataGridView 는 UIA 에서 Grid 패턴으로 노출되므로, 여기 걸리는
    컨트롤이 곧 고객목록/상담이력을 실제로 담고 있는 표다.
    """
    try:
        grid = ctrl.iface_grid
        rows, cols = grid.CurrentRowCount, grid.CurrentColumnCount
    except Exception:
        return []

    from pywinauto.controls.uiawrapper import UIAWrapper
    from pywinauto.uia_element_info import UIAElementInfo

    lines = [f"    >>> GRID 발견: {rows} 행 x {cols} 열"]
    for r in range(min(rows, max_rows)):
        cells = []
        for c in range(cols):
            try:
                cell = UIAWrapper(UIAElementInfo(grid.GetItem(r, c)))
                cells.append((cell.element_info.name or "").strip())
            except Exception:
                cells.append("?")
        lines.append(f"    >>> row {r}: {cells}")
    if rows > max_rows:
        lines.append(f"    >>> ... {rows - max_rows} 행 더 있음")
    return lines


def walk(ctrl, depth: int, max_depth: int, grid_rows: int, out: list[str]) -> None:
    out.append(describe(ctrl, depth))
    out.extend(grid_preview(ctrl, grid_rows))

    if depth >= max_depth:
        return
    try:
        children = ctrl.children()
    except Exception as exc:
        out.append(f"{'  ' * (depth + 1)}(자식 조회 실패: {exc})")
        return
    for child in children:
        walk(child, depth + 1, max_depth, grid_rows, out)


def main() -> int:
    parser = argparse.ArgumentParser(description="CRM4enterprise UI 트리 덤프")
    parser.add_argument("--process", default=PROCESS_NAME, help="대상 프로세스 이름")
    parser.add_argument("--max-depth", type=int, default=8, help="탐색 깊이 (기본 8)")
    parser.add_argument("--grid-rows", type=int, default=3, help="표 미리보기 행 수 (기본 3)")
    parser.add_argument("--out", default=None, help="저장할 파일 경로")
    args = parser.parse_args()

    try:
        import pywinauto  # noqa: F401
        import psutil  # noqa: F401
    except ImportError:
        print("의존성이 없습니다. 먼저 실행하세요: pip install -r tools/requirements.txt")
        return 1

    windows = list(iter_windows(args.process))
    if not windows:
        print(f"{args.process} 창을 찾지 못했습니다. 프로그램이 실행 중인지 확인하세요.")
        return 1

    out: list[str] = [
        f"# CRM4enterprise UI 덤프 - {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"# 창 {len(windows)} 개, max_depth={args.max_depth}",
        "",
    ]
    for win in windows:
        out.append(f"===== 창: {win.element_info.name!r} =====")
        try:
            walk(win, 0, args.max_depth, args.grid_rows, out)
        except Exception:
            out.append(traceback.format_exc())
        out.append("")

    path = Path(args.out or f"crm4_ui_dump_{datetime.now():%Y%m%d_%H%M%S}.txt")
    path.write_text("\n".join(out), encoding="utf-8")
    print(f"저장 완료: {path.resolve()}")
    print(f"총 {len(out)} 줄. GRID 라고 표시된 줄이 데이터가 실제로 들어있는 표입니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
