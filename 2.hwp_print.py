# -*- coding: utf-8 -*-
r"""
HWP 일괄 인쇄 스크립트 (2.hwp_print.py)
==================================================

여러 개의 한글(HWP) 파일을, 파일마다 지정한 부수만큼 자동으로 인쇄합니다.
한/글(Hancom Office) 프로그램이 설치된 Windows PC에서만 동작합니다
(한글의 인쇄 기능을 그대로 이용하기 때문입니다).

------------------------------------------------------------------
사전 준비 (최초 1회)
------------------------------------------------------------------
1. 명령 프롬프트(cmd)에서 아래 명령으로 필요한 라이브러리 설치
       pip install -r requirements.txt
   (pywin32 = 한글/프린터 제어용, openpyxl = 엑셀 목록 읽기용)

2. 컴퓨터에 한글(HWP) 프로그램이 설치되어 있어야 합니다.

3. 처음 실행하면 "다른 프로그램이 한글을 제어하려고 합니다" 같은 보안
   경고 창이 각 파일마다 뜰 수 있습니다. 0.setup_security_module.py를
   한 번 실행해 한컴 보안모듈을 등록하세요 (README.md 참고).

------------------------------------------------------------------
사용 방식 2가지 — 첫 번째 인자가 "폴더"인지 "목록 파일"인지로 자동 구분
------------------------------------------------------------------
[방식 A] 폴더 지정 (파일들이 한 폴더에 있고, 회차별로 경로가 안 바뀔 때)

    python 2.hwp_print.py "C:\경로\제출서류폴더"

    지정한 폴더 바로 안의 *.hwp 파일을 전부 찾아 인쇄합니다. 부수는
    - 폴더 안에 "부수목록.csv"(파일명,부수 2열)가 있으면 그 값을 사용
      (--csv 옵션으로 다른 CSV를 지정할 수도 있음)
    - 코드 위쪽 MANUAL_COPIES 딕셔너리에 적어둔 값 (CSV에 없는 파일만)
    - 둘 다 없으면 DEFAULT_COPIES(기본 1부)
    순서로 결정됩니다.

[방식 B] 목록 파일(.xlsx 또는 .csv) 지정 (폴더 복사 => 내용 수정 => 출력을
          반복하는 "회차"형 작업처럼, 경로 자체가 매번 조금씩 바뀔 때 추천)

    python 2.hwp_print.py "C:\경로\print_list_회차명.xlsx"
    python 2.hwp_print.py "C:\경로\부수목록.csv"

    1.hwp_list_maker.py가 만든 엑셀 목록을 그대로 넣으면 됩니다(첫 번째
    시트를 읽음). 각 행은 아래처럼 "경로 조각들 + 부수"로 구성합니다.
    (열 이름은 자유. 첫 번째 줄은 항상 제목 줄로 취급되어 건너뜁니다)

        1차경로,2차경로,파일명,부수
        C:\Users\me\제출서류,1차,사업계획서.hwp,3
        C:\Users\me\제출서류,1차,재무제표.hwp,5

    - 맨 마지막 열 = 부수
    - 그 앞의 모든 열 = 순서대로 이어붙여서(경로 조각1\경로 조각2\...\파일명)
      전체 파일 경로를 만듭니다. 경로 조각은 몇 개든 상관없습니다.
    - 다음 회차에는 폴더를 복사한 뒤, "2차경로" 열의 값만 바꿔주면
      전체 경로를 다시 안 적어도 됩니다.
    - 이 방식은 폴더를 스캔하지 않고 목록에 적힌 파일만 정확히 대상으로
      하므로, 하위 폴더 구조가 어떻든 상관없습니다.
    - 엑셀에서 부수를 고친 뒤에는 반드시 저장(Ctrl+S)하고 실행하세요.
      (저장하지 않은 수정 내용은 읽히지 않습니다)

------------------------------------------------------------------
공통 옵션
------------------------------------------------------------------
    --printer 이름    기본 프린터가 아닌 다른 프린터 사용
                     ([주의] 자동화 인쇄에서만 적용됨. 셸 인쇄로 전환된
                     문서는 항상 Windows 기본 프린터로 출력)
    --dry-run        실제로 인쇄하지 않고 "몇 부씩, 어느 경로로" 인쇄될지 확인
    --yes            확인 질문 없이 바로 인쇄 시작 (자동화용)
    --shell-only     자동화 인쇄를 건너뛰고 처음부터 셸 인쇄만 사용
                     (아래 "인쇄 방식" 참고)
    --visible        자동화 인쇄 중 한글 창을 화면에 보이게 함(디버깅용).
                     생략 시 창을 숨긴 채 진행
    --delay 초        인쇄 명령 후 대기 시간(기본 1.5초). 셸 인쇄에서는
                     SHELL_PRINT_GAP(6초)보다 작으면 6초가 적용됨
    --spool-timeout 초  인쇄 작업이 프린터 대기열에 접수되는지 확인하는 최대
                     대기 시간(기본 30초). 매우 큰 문서가 시간 초과로 실패
                     처리되면 늘려서 실행하세요.
    --csv PATH       [방식 A 전용] 부수 목록 CSV 경로 직접 지정

------------------------------------------------------------------
인쇄 방식 (기본: 자동화 인쇄 우선 → 실패 시 셸 인쇄로 전환)
------------------------------------------------------------------
[자동화(COM) 인쇄 - 1순위]
    한글 본체를 자동화로 제어해 인쇄합니다. --printer로 프린터를 지정할
    수 있어, 다른 프린터나 다른 PC에서도 쓸 수 있도록 1순위로 둡니다.
    한글 창은 기본적으로 숨긴 채 진행합니다(--visible로 표시 가능).

[셸 인쇄 - 우회 경로]
    탐색기에서 우클릭 → 인쇄와 같은 방식. 실제 출력은 한글 인쇄 관리자
    (HwpPrnMng.exe)가 담당하므로 한글 본체의 인쇄 오류를 피해 갑니다.
    자동화 인쇄가 한 번 실패하면 그 문서부터 끝까지 이 방식으로 진행하며,
    인쇄 명령 사이에 최소 6초 간격(코드 상단 SHELL_PRINT_GAP)을 두고
    대기열 접수를 확인한 뒤 다음 문서로 넘어갑니다.

[참고 - 이 PC(한글 2018 + Brother MFC-T800W)의 경우]
    자동화 인쇄 실행 순간 한글이 비정상 종료(MSVCR120.dll, 0xc0000409)
    하는 것이 확인되었습니다. 가상 PDF 프린터에서는 정상이고, Brother
    에서는 창 표시 여부와 무관하게 매번 종료되며, VC++ 2013 런타임을
    최신화해도 해결되지 않았습니다. 그래서 이 PC에서는 첫 문서에서 자동화가
    실패한 뒤 셸 인쇄로 넘어가게 되며(첫 문서도 셸로 다시 인쇄되므로 누락
    없음), 그 기록이 print_debug 로그에 남습니다. 이 과정을 생략하려면
    --shell-only로 실행하세요.

------------------------------------------------------------------
인쇄 누락 방지 (스풀러 접수 확인)
------------------------------------------------------------------
인쇄 명령을 보낸 뒤, 프린터 대기열(스풀러)에 작업이 실제로 나타나는
것을 확인하고 나서야 다음 부수/다음 파일로 넘어갑니다. 접수가 확인되지
않으면 그 파일은 '실패'로 기록되므로, 출력이 끝난 뒤 로그에서 실패
항목만 다시 인쇄하면 됩니다.
(이전 버전은 명령만 보내고 고정 시간만 기다렸기 때문에, 문서가 크거나
한글이 바쁘면 다음 명령과 겹쳐 일부 문서가 '성공'으로 기록되고도
실제로는 출력되지 않는 문제가 있었습니다)

------------------------------------------------------------------
실패 원인 진단 기록 (print_debug_*.log)
------------------------------------------------------------------
자동화(COM) 인쇄나 셸 인쇄가 실패하면, 그 원인(예외 메시지, 상세
traceback)과 환경 정보(한글 버전, 연결 방식, 보안모듈 등록 여부,
감시 프린터)를 로그 폴더의 print_debug_날짜시간.log 파일에 자동으로
남깁니다. 셸 방식 재시도가 성공해서 인쇄 자체는 됐더라도 실패 원인
기록은 보존되므로, 이 파일로 "왜 자동화 인쇄가 실패하는지"를 나중에
분석할 수 있습니다. 결과 로그(print_log_*.txt)에도 해당 파일 경로와
실패 원인 요약이 함께 적힙니다.
(실패가 한 건도 없으면 이 파일은 생성되지 않습니다)

예:
    python 2.hwp_print.py "C:\제출서류" --dry-run
    python 2.hwp_print.py "C:\제출서류\print_list_회차명.xlsx" --dry-run
    python 2.hwp_print.py "C:\제출서류\부수목록.csv" --yes
"""

import argparse
import csv
import datetime
import os
import sys
import time
import traceback
import unicodedata
from pathlib import Path

# ============================================================
# [방식 A 전용] 코드 안에 직접 부수를 적고 싶다면 여기에 채워 넣으세요.
#   예: MANUAL_COPIES = {"사업계획서.hwp": 3, "재무제표.hwp": 5}
# CSV 목록에 있는 파일은 CSV 값이 우선 적용되고, CSV에 없는 파일에
# 한해서 아래 값이 사용됩니다. (방식 B에서는 사용되지 않습니다)
# ============================================================
MANUAL_COPIES = {
    # "파일명.hwp": 부수,
}

# [방식 A 전용] CSV / MANUAL_COPIES 어디에도 없는 파일에 적용되는 기본 부수
DEFAULT_COPIES = 1

# [방식 A 전용] --csv 옵션 없이 실행했을 때, 폴더 안에서 자동으로 찾아볼 CSV 이름들
AUTO_CSV_NAMES = ["부수목록.csv", "copies.csv"]

# 인쇄 목록으로 받을 수 있는 파일 확장자 (엑셀은 첫 번째 시트를 읽음)
LIST_EXCEL_SUFFIXES = (".xlsx", ".xlsm")
LIST_SUFFIXES = (".csv",) + LIST_EXCEL_SUFFIXES

# 셸 인쇄 방식에서 인쇄 명령 사이에 두는 최소 간격(초).
# 대기열 접수 확인과 별개로, 한글이 문서를 닫고 다음 명령을 받을
# 준비가 되도록 여유를 둔다. (--delay 값이 이보다 크면 그 값을 사용)
SHELL_PRINT_GAP = 6.0


# ------------------------------------------------------------------
# 실패 원인 진단 기록 (print_debug_*.log)
# ------------------------------------------------------------------
# 자동화(COM) 인쇄 실패 원인은 콘솔에만 출력되고, 셸 방식 재시도가
# 성공하면 결과 로그에는 '성공'으로만 남아 원인이 사라진다. 그래서
# 실패가 발생하면 예외 상세와 환경 정보(한글 버전 등)를 별도 파일로
# 보존해 나중에 분석할 수 있게 한다. 실패가 한 건도 없으면 파일은
# 생성되지 않는다.

_diag = {"path": None, "context": [], "entries": 0}


def _diag_setup(log_dir):
    """진단 기록 파일 경로를 준비 (실제 파일은 첫 실패가 기록될 때 생성)"""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    _diag["path"] = Path(log_dir) / f"print_debug_{ts}.log"


def _diag_context(message):
    """실패 분석에 도움이 되는 배경 정보(버전/연결 방식 등)를 기억해 두었다가
    첫 실패가 기록될 때 함께 파일에 적는다. 같은 내용은 한 번만 기억한다."""
    for line in _diag["context"]:
        if line.endswith(message):
            return
    stamp = datetime.datetime.now().strftime("%H:%M:%S")
    _diag["context"].append(f"[{stamp}] {message}")


def _diag_log(message):
    """실패 내용을 진단 파일에 기록. 기록 실패가 인쇄를 막지는 않는다."""
    path = _diag["path"]
    if path is None:
        return
    try:
        with open(path, "a", encoding="utf-8") as f:
            if _diag["entries"] == 0:
                import platform
                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f.write(f"인쇄 실패 진단 기록 - {now}\n")
                f.write(f"Python {sys.version.split()[0]} / {platform.platform()}\n")
                for line in _diag["context"]:
                    f.write(line + "\n")
                f.write("=" * 60 + "\n")
            stamp = datetime.datetime.now().strftime("%H:%M:%S")
            f.write(f"[{stamp}] {message}\n")
        _diag["entries"] += 1
    except Exception:
        pass


def _normalize(name):
    return name.strip().lower()


def _display_width(s):
    """한글 등 전각 문자는 2칸으로 계산해 터미널 출력 정렬에 사용"""
    width = 0
    for ch in s:
        width += 2 if unicodedata.east_asian_width(ch) in ("F", "W") else 1
    return width


def _pad(s, width):
    return s + " " * max(0, width - _display_width(s))


def _read_csv_rows(csv_path):
    """CSV를 읽어 원시 행(row) 리스트로 반환. 인코딩을 자동으로 시도."""
    last_err = None
    # 엑셀에서 저장한 CSV는 CP949(EUC-KR)인 경우가 많아 여러 인코딩을 시도
    # (엑셀 "유니코드 텍스트" 저장을 대비해 UTF-16도 시도)
    text = None
    for encoding in ("utf-8-sig", "utf-8", "cp949", "utf-16"):
        try:
            with open(csv_path, "r", encoding=encoding, newline="") as f:
                text = f.read()
            break
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            continue
    if text is None:
        raise RuntimeError(f"CSV 파일 인코딩을 읽을 수 없습니다: {csv_path} ({last_err})")

    # 구분자 자동 감지: 탭이 쉼표보다 많이 쓰였으면 탭 구분 파일로 판단
    # (엑셀에서 "텍스트(탭으로 분리)" 형식으로 저장된 경우 대비)
    delimiter = "\t" if text.count("\t") > text.count(",") else ","
    return list(csv.reader(text.splitlines(), delimiter=delimiter))


def _read_xlsx_rows(xlsx_path):
    """엑셀(.xlsx) 첫 번째 시트를 읽어 CSV와 같은 문자열 행 리스트로 반환."""
    try:
        import openpyxl
    except ImportError:
        raise SystemExit(
            "엑셀 목록을 읽으려면 openpyxl이 필요합니다.\n"
            "명령 프롬프트에서 'pip install openpyxl' 을 실행한 뒤 다시 시도해 주세요."
        )

    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    except Exception as e:
        raise SystemExit(f"엑셀 파일을 열 수 없습니다: {xlsx_path} ({e})")

    def cell_text(v):
        if v is None:
            return ""
        # 엑셀은 숫자를 실수로 돌려주는 경우가 있어 2.0 → "2" 로 정리
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    try:
        ws = wb.worksheets[0]
        return [[cell_text(v) for v in row] for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _read_list_rows(list_path):
    """목록 파일(.xlsx 또는 .csv)을 확장자에 맞게 읽어 행 리스트로 반환"""
    if Path(list_path).suffix.lower() in LIST_EXCEL_SUFFIXES:
        return _read_xlsx_rows(list_path)
    return _read_csv_rows(list_path)


# ------------------------------------------------------------------
# 방식 A: 폴더 스캔 + (파일명, 부수) CSV
# ------------------------------------------------------------------

def load_csv_copies(csv_path):
    """CSV(파일명,부수)를 읽어 {파일명: 부수} 딕셔너리로 반환"""
    copies = {}
    if csv_path is None or not csv_path.exists():
        return copies

    rows = _read_csv_rows(csv_path)
    for row in rows:
        if not row or len(row) < 2:
            continue
        name, qty = row[0].strip(), row[1].strip()
        if not name or _normalize(name) in ("파일명", "filename", "file", "파일"):
            continue  # 빈 줄/헤더 행 건너뛰기
        try:
            copies[name] = int(qty)
        except ValueError:
            print(f"  [경고] CSV의 '{name}' 행 부수 값('{qty}')을 숫자로 읽을 수 없어 건너뜁니다.")
    return copies


def find_auto_csv(folder):
    for name in AUTO_CSV_NAMES:
        candidate = folder / name
        if candidate.exists():
            return candidate
    return None


def resolve_targets_from_folder(folder, csv_path):
    """폴더 내 hwp 파일 목록과 각 파일의 인쇄 부수를 결정 (방식 A)"""
    hwp_files = sorted(
        p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".hwp"
    )
    if not hwp_files:
        raise SystemExit(f"'{folder}' 폴더에서 .hwp 파일을 찾지 못했습니다.")

    csv_copies = load_csv_copies(csv_path)
    csv_norm = {_normalize(k): v for k, v in csv_copies.items()}
    manual_norm = {_normalize(k): v for k, v in MANUAL_COPIES.items()}

    targets = []
    unmatched = []
    zero_skipped = []
    for f in hwp_files:
        qty = csv_norm.get(_normalize(f.name))
        if qty is None:
            qty = csv_norm.get(_normalize(f.stem))
        if qty is None:
            qty = manual_norm.get(_normalize(f.name))
        if qty is None:
            qty = manual_norm.get(_normalize(f.stem))

        matched = qty is not None
        if qty is None:
            qty = DEFAULT_COPIES
            unmatched.append(f.name)

        qty = int(qty)
        if qty <= 0:
            zero_skipped.append(f.name)
            continue

        targets.append((f, qty, matched))

    return targets, unmatched, zero_skipped, []


# ------------------------------------------------------------------
# 방식 B: CSV의 각 행 = 경로 조각들 + 부수 (회차형 작업용)
# ------------------------------------------------------------------

def resolve_targets_from_path_csv(csv_path):
    """CSV의 각 행([경로조각...,파일명], 부수)을 읽어 전체 경로를 조립 (방식 B)

    - 첫 번째 줄은 항상 제목(헤더) 줄로 취급되어 건너뜁니다.
    - 마지막 열 = 부수, 그 앞의 모든 열 = 순서대로 이어붙일 경로 조각.
    - 조립한 경로가 상대경로이면 CSV 파일이 있는 폴더를 기준으로 해석합니다.
    """
    rows = _read_list_rows(csv_path)
    data_rows = rows[1:]  # 첫 줄은 헤더로 간주하고 건너뜀
    if not data_rows:
        raise SystemExit(f"목록 파일에 처리할 데이터 행이 없습니다 (첫 줄은 제목 줄로 건너뜁니다): {csv_path}")

    targets = []
    missing = []
    for line_no, row in enumerate(data_rows, start=2):
        cells = [c.strip() for c in row]
        if not any(cells):
            continue  # 완전히 빈 행 건너뛰기

        if len(cells) < 2:
            print(f"  [경고] {line_no}행: 경로/부수 열이 부족해 건너뜁니다 -> {row}")
            continue

        *path_parts, qty_str = cells
        path_parts = [p for p in path_parts if p]  # 중간에 비워둔 조각은 무시
        if not path_parts:
            print(f"  [경고] {line_no}행: 경로 정보가 없어 건너뜁니다.")
            continue

        try:
            qty = int(qty_str)
        except ValueError:
            print(f"  [경고] {line_no}행: 부수 값('{qty_str}')을 숫자로 읽을 수 없어 건너뜁니다.")
            continue

        if qty <= 0:
            continue  # 0부는 조용히 건너뜀

        full_path = Path(path_parts[0])
        for part in path_parts[1:]:
            full_path = full_path / part
        full_path = full_path.expanduser()
        if not full_path.is_absolute():
            full_path = csv_path.parent / full_path
        full_path = full_path.resolve()

        if not full_path.exists():
            missing.append((line_no, str(full_path)))
            continue

        targets.append((full_path, qty, True))

    return targets, [], [], missing


# ------------------------------------------------------------------
# 인쇄 실행 / 로그 (두 방식 공통)
# ------------------------------------------------------------------

def _set_print_option(hwp, key, value):
    """인쇄 설정 항목을 지정. SetItem 방식을 우선 쓰고, 안 되면 속성 대입 방식 시도.

    (pywin32가 만든 정적 타입 래퍼(gen_py 캐시)에서는 속성 대입이
    AttributeError로 실패하는 경우가 있어 두 방식을 모두 시도합니다)
    """
    try:
        hwp.HParameterSet.HPrint.HSet.SetItem(key, value)
        return True
    except Exception:
        pass
    try:
        setattr(hwp.HParameterSet.HPrint, key, value)
        return True
    except Exception:
        return False


# ------------------------------------------------------------------
# 프린터 대기열(스풀러) 확인 - 인쇄 명령이 실제로 접수됐는지 검증
# ------------------------------------------------------------------
# 인쇄는 "명령 전송"과 "실제 접수"가 분리된 비동기 동작이라, 명령만 보내고
# 다음 파일로 넘어가면 한글이 바쁠 때 보낸 명령이 조용히 유실될 수 있다.
# 그래서 명령을 보낸 뒤 프린터 대기열에 새 작업이 나타나는 것을 확인하고
# 나서야 다음 단계(다음 부수/다음 파일/문서 닫기)로 진행한다.

# 인쇄 작업 Status 비트: 스풀링(응용프로그램 → 스풀러 데이터 전달) 진행 중
_JOB_STATUS_SPOOLING = 0x00000008

# 안내 문구를 한 번만 출력하기 위한 상태 플래그
_spool_notice_shown = False
_shell_printer_notice_shown = False


def _spool_watch_printer(printer_name):
    """대기열 감시에 사용할 프린터 이름을 반환. 사용할 수 없으면 None.

    None이 반환되면 예전처럼 '고정 시간 대기' 방식으로만 동작한다
    (접수 여부를 검증하지 못하므로 인쇄 누락을 자동으로 감지할 수 없음).
    """
    global _spool_notice_shown
    try:
        import win32print
        name = printer_name or win32print.GetDefaultPrinter()
        handle = win32print.OpenPrinter(name)
        win32print.ClosePrinter(handle)
        _diag_context(f"감시 대상 프린터: {name}")
        return name
    except Exception as e:
        if not _spool_notice_shown:
            _spool_notice_shown = True
            print(f"  [안내] 프린터 대기열을 확인할 수 없어 고정 대기 방식으로 동작합니다: {e}")
            print("         (인쇄 누락을 자동으로 감지하지 못하니 출력물 매수를 꼭 확인하세요)")
            _diag_context(f"프린터 대기열 확인 불가: {type(e).__name__}: {e}")
        return None


def _query_jobs(printer):
    """해당 프린터 대기열의 {작업ID: 상태비트} 딕셔너리를 반환. 조회 실패 시 None."""
    try:
        import win32print
        handle = win32print.OpenPrinter(printer)
        try:
            jobs = win32print.EnumJobs(handle, 0, -1, 1)
        finally:
            win32print.ClosePrinter(handle)
        return {job["JobId"]: job["Status"] for job in jobs}
    except Exception:
        return None


def _snapshot_jobs(printer):
    """인쇄 명령을 보내기 직전의 대기열 작업 ID 집합. 조회 실패 시 None."""
    jobs = _query_jobs(printer)
    return set(jobs) if jobs is not None else None


def _wait_print_job(printer, before_ids, timeout):
    """인쇄 명령 후 새 작업이 대기열에 나타나 스풀링까지 끝나기를 기다린다.

    반환값:
      True  - 새 작업 접수 확인 (문서 데이터가 스풀러로 넘어감 → 닫아도 안전)
      False - timeout 안에 새 작업이 나타나지 않음 (명령이 유실된 것으로 판단)
    대기 중 대기열 조회 자체가 실패하면 검증을 포기하고 True를 반환한다.
    """
    deadline = time.time() + timeout
    new_job_id = None
    while True:
        jobs = _query_jobs(printer)
        if jobs is None:
            return True  # 스풀러 조회 불가 - 검증 생략 (기존 방식과 동일하게 진행)
        if new_job_id is None:
            fresh = set(jobs) - before_ids
            if fresh:
                new_job_id = min(fresh)
                deadline = time.time() + timeout  # 접수 확인 → 스풀링 완료 대기로 전환
        if new_job_id is not None:
            status = jobs.get(new_job_id)
            if status is None:
                return True  # 이미 인쇄를 마치고 대기열에서 사라짐
            if not (status & _JOB_STATUS_SPOOLING):
                return True  # 스풀링 완료 (이후 출력은 프린터가 알아서 진행)
        if time.time() >= deadline:
            # 작업이 보이긴 했다면 접수는 된 것 (스풀링이 오래 걸리는 큰 문서)
            return new_job_id is not None
        time.sleep(0.2)


def _shell_print(filepath, copies, delay, spool_timeout, printer_name=None):
    """Windows 셸 '인쇄' 명령으로 출력 (탐색기에서 우클릭 → 인쇄와 동일).

    자동화 인쇄가 실패했을 때의 우회 경로(--shell-only면 처음부터 사용).
    셸 인쇄는 한글 본체(자동화)가 아니라 한글 인쇄 관리자(HwpPrnMng.exe /p)가
    처리하므로, 자동화 인쇄가 프린터 드라이버와 충돌해 한글이 비정상
    종료되는 환경(예: 한글 2018 + Brother 드라이버)에서도 정상 출력된다.

    - 명령을 보낸 뒤 프린터 대기열에 작업이 실제로 나타나는지 확인하고
      나서야 다음 명령을 보낸다 (한글이 바쁠 때 명령이 유실되는 것을 감지).
    - 접수 확인과 별개로 명령 사이에 최소 SHELL_PRINT_GAP초 간격을 둔다.
    - 대기열 확인이 불가능한 환경에서는 간격 대기만으로 동작한다.
    - 같은 파일에 인쇄 명령을 연달아 보내면 한글이 두 번째부터는 무시하므로
      (이미 열린 문서로 판단), 2부째부터는 임시 복사본을 만들어 인쇄한다.
    - 셸 인쇄는 항상 '기본 프린터'로 출력된다 (--printer 적용 불가).
    """
    global _shell_printer_notice_shown
    import shutil
    import tempfile

    if printer_name and not _shell_printer_notice_shown:
        _shell_printer_notice_shown = True
        print(f"  [주의] 셸 인쇄 방식은 --printer('{printer_name}') 지정을 적용하지 못하고 "
              "기본 프린터로 인쇄합니다.")

    filepath = Path(filepath)
    watch = _spool_watch_printer(None)  # 셸 인쇄는 기본 프린터로 나간다
    gap = max(delay, SHELL_PRINT_GAP)  # 인쇄 명령 사이 최소 간격

    temp_dir = None
    try:
        # 2부 이상이면 부수-1개의 임시 복사본을 준비
        print_files = [filepath]
        if copies > 1:
            temp_dir = Path(tempfile.mkdtemp(prefix="hwp_print_"))
            for n in range(2, copies + 1):
                t = temp_dir / f"{filepath.stem}_{n}부째{filepath.suffix}"
                shutil.copy2(filepath, t)
                print_files.append(t)

        for i, target in enumerate(print_files, start=1):
            before = _snapshot_jobs(watch) if watch else None
            os.startfile(str(target), "print")
            if copies > 1:
                print(f"      ... {i}/{copies}부 인쇄 명령 전송(셸 방식)")
            if before is not None:
                if not _wait_print_job(watch, before, spool_timeout):
                    raise RuntimeError(
                        f"{i}부째 인쇄 작업이 {spool_timeout:.0f}초 안에 프린터 대기열에 "
                        "나타나지 않았습니다. 한글이 명령을 놓쳤거나 대화 상자에 막혀 "
                        "있을 수 있으니 화면의 한글 창과 프린터를 확인한 뒤, "
                        "로그의 실패 항목만 다시 인쇄하세요."
                    )
            time.sleep(gap)
    finally:
        if temp_dir is not None:
            # 접수가 확인된 경우 복사본은 이미 읽힌 뒤라 바로 지워도 안전하지만,
            # 확인 불가 환경(고정 대기)에서는 스풀로 넘어갈 시간을 넉넉히 준다
            time.sleep(gap if not watch else 1.0)
            shutil.rmtree(temp_dir, ignore_errors=True)


def _connect_hwp(win32, visible):
    """한글 프로그램에 새로 연결하고 기본 설정을 적용"""
    # dynamic.Dispatch: 타입 캐시(gen_py)를 거치지 않는 유연한 연결 방식.
    # 정적 래퍼에서는 인쇄 설정 속성 대입이 막히는 경우가 있어 이를 우선 사용.
    try:
        hwp = win32.dynamic.Dispatch("HWPFrame.HwpObject")
        dispatch_mode = "dynamic"
    except Exception:
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        dispatch_mode = "static(gen_py)"

    # 자동화 보안 팝업 우회: 0.setup_security_module.py로 등록한 모듈을 우선 사용
    # (등록 전이면 실패해도 무방 - 보안 창이 뜰 뿐 동작은 함)
    registered = None
    for module_name in ("FilePathCheckerModule", "FilePathCheckerModuleExample"):
        try:
            if hwp.RegisterModule("FilePathCheckDLL", module_name):
                registered = module_name
                break
        except Exception:
            continue

    try:
        hwp.XHwpWindows.Item(0).Visible = visible
    except Exception:
        pass

    # 실패 분석용 환경 정보 기억 (첫 실패가 기록될 때 진단 파일에 함께 적힘)
    try:
        hwp_version = hwp.Version
    except Exception:
        hwp_version = "(확인 불가)"
    _diag_context(
        f"한글 연결: 방식={dispatch_mode}, 한글 버전={hwp_version}, "
        f"보안모듈={registered or '미등록'}, visible={visible}"
    )

    return hwp


def print_hwp_files(targets, printer_name, visible, delay, spool_timeout, use_com=True):
    """targets를 차례로 인쇄하고 (경로, 부수, 결과) 목록을 반환.

    기본은 자동화(COM) 인쇄를 먼저 시도하고, 실패하면 그 파일부터 셸
    인쇄 방식으로 전환한다. use_com=False면 처음부터 셸 인쇄만 사용한다.
    """
    import win32com.client as win32

    results = []
    if use_com:
        hwp = _connect_hwp(win32, visible)
        # 자동화 인쇄가 실제로 접수됐는지 감시할 프린터 대기열
        watch = _spool_watch_printer(printer_name)
        use_shell = False  # 자동화 인쇄가 한 번 죽으면 이후 파일은 처음부터 셸 방식 사용
    else:
        hwp = None
        watch = None
        use_shell = True
    for idx, (filepath, copies, _matched) in enumerate(targets):
        if use_shell:
            try:
                _shell_print(filepath, copies, delay, spool_timeout, printer_name)
                results.append((str(filepath), copies, "성공(셸 인쇄 방식)"))
                print(f"  [완료] {filepath}  x {copies}부 (셸 인쇄 방식)")
            except Exception as e:
                _diag_log(
                    f"셸 인쇄 실패: {filepath} ({copies}부)\n"
                    f"    예외: {type(e).__name__}: {e}\n"
                    f"{traceback.format_exc()}"
                )
                results.append((str(filepath), copies, f"실패(셸 인쇄): {e}"))
                print(f"  [실패] {filepath} (셸 인쇄): {e}")
            continue

        step = "파일 열기"
        done_copies = 0  # 접수까지 확인된 부수 (셸 재시도 시 같은 부수를 또 찍지 않도록)
        try:
            # Open은 원래 (경로, 형식, 옵션) 3개 인자를 받는 명령.
            # dynamic 연결에서는 인자를 생략하면 "매개 변수 개수" 오류가
            # 날 수 있어 전부 채워서 호출하고, 안 되면 축약형도 시도.
            # "forceopen:true"는 "다른 곳에서 사용 중" 확인 창을 건너뜀.
            try:
                opened = hwp.Open(str(filepath), "HWP", "forceopen:true")
            except Exception:
                opened = hwp.Open(str(filepath))  # 3인자 호출이 안 되는 환경이면 축약형 재시도
            # Open은 실패해도 예외 대신 False를 반환하므로 결과를 꼭 확인해야 한다
            # (안 하면 직전/빈 문서를 대신 인쇄하고도 '성공'으로 기록됨)
            if not opened:
                raise RuntimeError("한글이 파일을 열지 못했습니다 (Open이 실패를 반환)")

            # 부수 설정(Copies)을 인쇄 명령에 끼워 넣는 방식은 한글 버전에
            # 따라 인쇄 순간 프로그램이 죽는 문제가 있어, 대신 "기본 인쇄를
            # 부수만큼 반복 실행"하는 방식을 사용 (가장 호환성이 높음)
            step = "인쇄 실행"
            for n in range(copies):
                before = _snapshot_jobs(watch) if watch else None
                hwp.HAction.GetDefault("Print", hwp.HParameterSet.HPrint.HSet)
                if printer_name:
                    _set_print_option(hwp, "PrinterName", printer_name)
                executed = hwp.HAction.Execute("Print", hwp.HParameterSet.HPrint.HSet)
                # Execute도 실패 시 예외 없이 False를 반환할 수 있다
                if not executed:
                    raise RuntimeError(f"{n + 1}부째 인쇄 명령이 실행되지 않았습니다 (Execute가 실패를 반환)")
                if before is not None:
                    # 인쇄는 비동기라, 접수 확인 없이 문서를 닫으면 작업이 유실될 수 있음
                    if not _wait_print_job(watch, before, spool_timeout):
                        raise RuntimeError(
                            f"{n + 1}부째 인쇄 작업이 {spool_timeout:.0f}초 안에 "
                            "프린터 대기열에 나타나지 않았습니다"
                        )
                done_copies += 1
                time.sleep(delay)  # 한글/스풀러가 다음 명령을 받을 여유
                if copies > 1:
                    suffix = "인쇄 접수 확인" if watch else "인쇄 명령 전송"
                    print(f"      ... {n + 1}/{copies}부 {suffix}")

            step = "문서 닫기"
            try:
                hwp.Clear(1)  # 저장하지 않고 문서 닫기
            except Exception:
                pass  # 이미 인쇄는 됐으므로 닫기 실패는 무시

            results.append((str(filepath), copies, "성공"))
            print(f"  [완료] {filepath}  x {copies}부")
        except Exception as e:
            # 실패 원인을 진단 파일에 보존 (콘솔에만 남고 사라지지 않도록).
            # com_error라면 args에 HRESULT 코드와 한글 쪽 상세 설명이 들어 있다.
            _diag_log(
                f"자동화 인쇄 실패: {filepath}\n"
                f"    단계: {step} / 접수 확인된 부수: {done_copies}/{copies}\n"
                f"    예외: {type(e).__name__}: {e}\n"
                f"    상세 args: {getattr(e, 'args', None)!r}\n"
                f"{traceback.format_exc()}"
            )

            # 셸 재시도 전에 자동화 쪽 연결부터 정리한다. 문서가 열린 채로
            # 파일을 잠그고 있으면 셸 인쇄가 '사용 중' 대화 상자에 막혀
            # 그 뒤 파일들 인쇄까지 연쇄로 실패할 수 있기 때문.
            try:
                hwp.Clear(1)
            except Exception:
                pass
            try:
                hwp.Quit()
            except Exception:
                pass
            hwp = None

            shell_ok = False
            if step == "인쇄 실행":
                # 자동화(COM) 인쇄가 한글을 죽이는 환경 대비:
                # 탐색기 우클릭 → 인쇄와 같은 셸 방식으로 자동 재시도.
                # 접수까지 확인된 부수는 빼고 남은 부수만 다시 인쇄한다.
                remaining = copies - done_copies
                print(f"  [재시도] 자동화 인쇄 실패({e}) → 남은 {remaining}부를 셸 인쇄 방식으로 재시도합니다...")
                try:
                    _shell_print(filepath, remaining, delay, spool_timeout, printer_name)
                    # 결과 로그에도 자동화 실패 원인을 한 줄로 남긴다 (상세는 진단 파일)
                    err_short = " ".join(str(e).split())
                    if len(err_short) > 100:
                        err_short = err_short[:100] + "…"
                    if done_copies:
                        status = f"성공(자동화 {done_copies}부 + 셸 인쇄 {remaining}부; 자동화 실패: {err_short})"
                    else:
                        status = f"성공(셸 인쇄 방식; 자동화 실패: {err_short})"
                    results.append((str(filepath), copies, status))
                    print(f"  [완료] {filepath}  x {copies}부 (셸 인쇄 방식)")
                    shell_ok = True
                    use_shell = True
                    _diag_log(f"→ 셸 인쇄 재시도 성공: {filepath} (남은 {remaining}부 출력)")
                    if idx + 1 < len(targets):
                        print("  [안내] 이 컴퓨터에서는 자동화 인쇄가 불안정하여, "
                              "남은 파일은 처음부터 셸 인쇄 방식으로 진행합니다.")
                except Exception as shell_err:
                    print(f"  [실패] 셸 인쇄도 실패했습니다: {shell_err}")
                    _diag_log(
                        f"→ 셸 인쇄 재시도도 실패: {filepath}\n"
                        f"    예외: {type(shell_err).__name__}: {shell_err}\n"
                        f"{traceback.format_exc()}"
                    )

            if not shell_ok:
                results.append((str(filepath), copies, f"실패({step} 단계): {e}"))
                print(f"  [실패] {filepath} ({step} 단계): {e}")

            if use_shell:
                continue  # 이후 파일은 셸 방식이므로 자동화 재연결 불필요

            # 다음 파일의 자동화 인쇄를 위해 새로 연결 시도
            try:
                time.sleep(1.0)
                hwp = _connect_hwp(win32, visible)
            except Exception as reconnect_err:
                _diag_log(
                    f"한글 재연결 실패 (남은 파일 미시도): {type(reconnect_err).__name__}: {reconnect_err}\n"
                    f"{traceback.format_exc()}"
                )
                print(f"  [중단] 한글 프로그램에 다시 연결하지 못했습니다: {reconnect_err}")
                # 남은 파일들을 실패로 기록하고 종료
                for rf, rc, _rm in targets[idx + 1:]:
                    results.append((str(rf), rc, "실패(한글 재연결 불가로 미시도)"))
                break

    if hwp is not None:
        try:
            hwp.Quit()
        except Exception:
            pass

    return results


def write_log(results, log_dir):
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_dir / f"print_log_{ts}.txt"
    total = 0
    fail_count = 0
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"인쇄 작업 기록 - {ts}\n")
        f.write("=" * 50 + "\n")
        for name, copies, status in results:
            f.write(f"{name}\t{copies}부\t{status}\n")
            if status.startswith("성공"):
                total += copies
            else:
                fail_count += 1
        f.write("=" * 50 + "\n")
        f.write(f"성공한 문서 기준 총 출력 부수: {total}부\n")
        if fail_count:
            f.write(f"실패한 파일 수: {fail_count}건 (위 목록에서 '실패'로 표시된 항목 확인)\n")
        if _diag["entries"] and _diag["path"]:
            f.write(f"실패 원인 상세 기록(분석용): {_diag['path']}\n")
    return log_path


def main():
    parser = argparse.ArgumentParser(
        description="HWP 파일을 부수에 맞춰 일괄 인쇄합니다. (폴더 또는 목록 파일(.xlsx/.csv)을 지정)"
    )
    parser.add_argument("target", help="hwp 파일들이 든 폴더 경로, 또는 [경로조각...,부수] 형식의 목록 파일(.xlsx/.csv) 경로")
    parser.add_argument("--csv", help="[폴더 지정 시] 파일명,부수 로 이루어진 CSV 경로 (생략 시 폴더 안 '부수목록.csv' 자동 탐색)")
    parser.add_argument("--printer", help="사용할 프린터 이름 (생략 시 기본 프린터 사용)")
    parser.add_argument("--visible", action="store_true",
                        help="자동화 인쇄 중 한글 프로그램 창을 화면에 표시 (생략 시 숨김)")
    parser.add_argument("--delay", type=float, default=1.5,
                        help=f"인쇄 명령 후 대기 시간(초). 기본 1.5초. 셸 인쇄에서는 최소 {SHELL_PRINT_GAP:g}초 적용")
    parser.add_argument("--spool-timeout", type=float, default=30.0,
                        help="인쇄 작업이 프린터 대기열에 접수되는지 확인하는 최대 대기 시간(초). 기본 30초")
    parser.add_argument("--shell-only", action="store_true",
                        help="자동화(COM) 인쇄를 건너뛰고 처음부터 셸 인쇄 방식만 사용. "
                             "생략 시 자동화 인쇄를 먼저 시도하고 실패하면 셸 인쇄로 전환")
    parser.add_argument("--yes", action="store_true", help="확인 절차 없이 바로 인쇄 시작")
    parser.add_argument("--dry-run", action="store_true", help="실제로 인쇄하지 않고 목록/부수만 확인")
    args = parser.parse_args()

    target = Path(args.target).expanduser().resolve()
    if not target.exists():
        raise SystemExit(f"경로를 찾을 수 없습니다: {target}")

    if target.is_dir():
        mode = "folder"
        if args.csv:
            csv_path = Path(args.csv).expanduser().resolve()
            if not csv_path.exists():
                raise SystemExit(f"지정한 CSV 파일을 찾을 수 없습니다: {csv_path}")
        else:
            csv_path = find_auto_csv(target)
        targets, unmatched, zero_skipped, missing = resolve_targets_from_folder(target, csv_path)
        log_dir = target
    elif target.is_file() and target.suffix.lower() in LIST_SUFFIXES:
        mode = "csv"
        if args.csv:
            print("[안내] 목록 파일을 직접 지정했기 때문에 --csv 옵션은 무시됩니다.")
        targets, unmatched, zero_skipped, missing = resolve_targets_from_path_csv(target)
        log_dir = target.parent
    else:
        raise SystemExit(
            f"'{target}' 은(는) 폴더도 아니고 목록 파일({', '.join(LIST_SUFFIXES)})도 아닙니다."
        )

    _diag_setup(log_dir)  # 인쇄 실패 시 원인을 남길 진단 파일 경로 준비

    print()
    if mode == "folder":
        print(f"대상 폴더: {target}")
        print(f"부수 목록 CSV: {csv_path if csv_path else '(없음 - MANUAL_COPIES/기본값 사용)'}")
    else:
        print(f"대상 목록 파일(경로 조합 방식): {target}")
    print("-" * 55)
    total_copies = 0
    for f, qty, matched in targets:
        mark = "" if matched else "  ← 목록에 없어 기본값 적용, 확인 필요!"
        # 폴더 방식은 파일명만, CSV 경로조합 방식은 조립된 전체 경로를 보여줌
        label = f.name if mode == "folder" else str(f)
        pad_width = 40 if mode == "folder" else 60
        print(f"  {_pad(label, pad_width)} {qty:>3d}부{mark}")
        total_copies += qty
    print("-" * 55)
    print(f"총 {len(targets)}개 파일 / {total_copies}부 인쇄 예정")

    if zero_skipped:
        print(f"\n[안내] 부수가 0으로 지정되어 건너뛴 파일: {', '.join(zero_skipped)}")
    if unmatched:
        print(f"[주의] CSV/코드 목록에 없어 기본값({DEFAULT_COPIES}부)이 적용된 파일이 있습니다:")
        for name in unmatched:
            print(f"   - {name}")
        print("       실제 필요한 부수가 맞는지 인쇄 전에 꼭 확인하세요.")
    if missing:
        print(f"\n[경고] 경로를 조합했지만 실제로 존재하지 않는 파일이 있어 인쇄 대상에서 제외했습니다:")
        for line_no, path_str in missing:
            print(f"   - {line_no}행 -> {path_str}")
        print("       회차 폴더명을 안 바꿨거나 오타가 없는지 확인해 주세요.")

    if args.dry_run:
        print("\n(dry-run 모드: 실제 인쇄는 진행하지 않았습니다)")
        return

    if not targets:
        print("\n인쇄할 파일이 없습니다. 종료합니다.")
        return

    if not args.yes:
        answer = input("\n위 내용대로 인쇄를 시작할까요? (y/n): ").strip().lower()
        if answer != "y":
            print("취소되었습니다.")
            return

    gap = max(args.delay, SHELL_PRINT_GAP)
    if args.shell_only:
        print(f"\n인쇄를 시작합니다. (셸 인쇄 방식만 사용, 인쇄 명령 간격 {gap:g}초)\n")
    else:
        print(f"\n인쇄를 시작합니다. (자동화 인쇄 우선, 실패 시 셸 인쇄로 전환 - "
              f"셸 인쇄 간격 {gap:g}초)\n")
    try:
        results = print_hwp_files(targets, args.printer, args.visible, args.delay,
                                  args.spool_timeout, use_com=not args.shell_only)
    except ImportError:
        raise SystemExit(
            "pywin32 모듈을 찾을 수 없습니다. 명령 프롬프트에서 'pip install pywin32'를 "
            "실행해 설치한 뒤 다시 시도해 주세요. (이 스크립트는 한/글 프로그램이 설치된 "
            "Windows PC에서만 실제 인쇄가 가능합니다)"
        )
    except Exception as e:
        raise SystemExit(
            f"한글 프로그램을 제어하는 중 오류가 발생했습니다: {e}\n"
            "- 한/글(Hancom Office) 프로그램이 설치되어 있는지 확인해 주세요.\n"
            "- 처음 실행이라면 보안 경고 팝업이 화면 뒤에 떠 있을 수 있습니다. "
            "--visible 옵션으로 다시 실행해 창을 확인해 보세요."
        )
    log_path = write_log(results, log_dir)
    print(f"\n인쇄 작업 완료. 결과 로그: {log_path}")
    if _diag["entries"] and _diag["path"]:
        print(f"인쇄 실패 원인 상세 기록(분석용): {_diag['path']}")
        print("  → 자동화(COM) 인쇄가 왜 실패했는지 이 파일에 남아 있습니다.")


if __name__ == "__main__":
    main()