# -*- coding: utf-8 -*-
r"""
한글 자동화 보안모듈 설치 스크립트 (0.setup_security_module.py)
=============================================================

한글(HWP)을 자동화로 제어할 때 문서마다 뜨는
"위 파일에 접근하려는 시도(파일의 손상 또는 유출의 위험 등)가 있습니다"
보안 승인 창을 없애기 위한 설치 도구입니다.

한컴이 공식 배포하는 보안 승인 모듈(DLL)을 내려받아 사용자 컴퓨터에
등록합니다. 관리자 권한은 필요 없습니다 (현재 사용자 계정에만 적용).

사용법 (1회만 실행하면 됩니다):
    python 0.setup_security_module.py

하는 일:
  1. 한컴 공식 GitHub 저장소에서 보안모듈 압축파일 다운로드
     (직접 다운로드가 안 되면, 안내에 따라 수동으로 받은 파일을 사용)
  2. DLL을 사용자 폴더(%APPDATA%\hwp_print_tool)에 복사
  3. 레지스트리 HKEY_CURRENT_USER\SOFTWARE\HNC\HwpAutomation\Modules 에
     "FilePathCheckerModule" 이름으로 DLL 경로 등록

설치 후 2.hwp_print.py를 실행하면 보안 창 없이 진행됩니다.
"""

import os
import shutil
import sys
import zipfile
from pathlib import Path

# 한컴 공식 개발자센터 아카이브 (hancom-io GitHub)의 보안모듈 배포 주소
OFFICIAL_ZIP_URL = (
    "https://github.com/hancom-io/devcenter-archive/raw/main/"
    "hwp-automation/%EB%B3%B4%EC%95%88%EB%AA%A8%EB%93%88(Automation).zip"
)

MODULE_NAME = "FilePathCheckerModule"  # 레지스트리 값 이름 (코드의 RegisterModule과 일치해야 함)
REG_KEY_PATH = r"SOFTWARE\HNC\HwpAutomation\Modules"


def find_dll_nearby(folder):
    """스크립트 폴더 또는 지정 폴더에서 보안모듈 DLL/ZIP을 찾음"""
    folder = Path(folder)
    # 1) DLL이 이미 풀려있는 경우
    for p in folder.glob("*.dll"):
        if "filepathchecker" in p.name.lower():
            return p
    # 2) 수동으로 받아둔 zip이 있는 경우
    for p in folder.glob("*.zip"):
        if "보안모듈" in p.name or "automation" in p.name.lower():
            extracted = extract_dll_from_zip(p, folder / "_보안모듈_압축해제")
            if extracted:
                return extracted
    return None


def extract_dll_from_zip(zip_path, out_dir):
    """zip 안에서 보안모듈 DLL을 찾아 꺼냄"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        dll_names = [n for n in zf.namelist() if n.lower().endswith(".dll")]
        if not dll_names:
            return None
        # FilePathChecker 이름이 들어간 DLL 우선, 없으면 첫 번째 DLL
        dll_names.sort(key=lambda n: ("filepathchecker" not in n.lower(), n))
        target = dll_names[0]
        zf.extract(target, out_dir)
        return out_dir / target


def download_zip(dest_path):
    """한컴 공식 GitHub에서 보안모듈 zip 다운로드"""
    import urllib.request

    print(f"한컴 공식 저장소에서 보안모듈을 내려받는 중...\n  {OFFICIAL_ZIP_URL}")
    req = urllib.request.Request(OFFICIAL_ZIP_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as resp, open(dest_path, "wb") as f:
        shutil.copyfileobj(resp, f)
    return dest_path


def register_registry(dll_path):
    """HKCU 레지스트리에 보안모듈 경로 등록 (관리자 권한 불필요)"""
    import winreg

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY_PATH) as key:
        winreg.SetValueEx(key, MODULE_NAME, 0, winreg.REG_SZ, str(dll_path))

    # 등록 확인
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY_PATH) as key:
        value, _ = winreg.QueryValueEx(key, MODULE_NAME)
    return value


def main():
    if os.name != "nt":
        raise SystemExit("이 스크립트는 Windows에서만 실행할 수 있습니다.")

    script_dir = Path(__file__).resolve().parent
    install_dir = Path(os.environ.get("APPDATA", str(Path.home()))) / "hwp_print_tool"
    install_dir.mkdir(parents=True, exist_ok=True)

    installed_dll = install_dir / f"{MODULE_NAME}.dll"

    # 이미 설치되어 있으면 레지스트리만 다시 확인/등록
    if installed_dll.exists():
        print(f"이미 설치된 DLL 발견: {installed_dll}")
        dll_source = installed_dll
    else:
        # 1순위: 스크립트 옆에 미리 받아둔 DLL/zip
        dll_source = find_dll_nearby(script_dir)

        # 2순위: 공식 저장소에서 자동 다운로드
        if dll_source is None:
            try:
                zip_path = install_dir / "보안모듈(Automation).zip"
                download_zip(zip_path)
                dll_source = extract_dll_from_zip(zip_path, install_dir / "_압축해제")
                if dll_source is None:
                    raise RuntimeError("압축파일 안에서 DLL을 찾지 못했습니다.")
            except Exception as e:
                print("\n[자동 다운로드 실패]", e)
                print("\n수동 설치 방법:")
                print("  1. 브라우저에서 아래 주소를 열어 압축파일을 내려받으세요.")
                print(f"     {OFFICIAL_ZIP_URL}")
                print("     (한컴 공식 GitHub: hancom-io/devcenter-archive)")
                print(f"  2. 내려받은 zip 파일을 이 스크립트와 같은 폴더({script_dir})에 두세요.")
                print("  3. 이 스크립트를 다시 실행하세요.")
                sys.exit(1)

        # DLL을 설치 폴더로 복사 (이름 통일)
        if Path(dll_source).resolve() != installed_dll.resolve():
            shutil.copy2(dll_source, installed_dll)
        print(f"DLL 설치 완료: {installed_dll}")

    # 레지스트리 등록
    try:
        registered = register_registry(installed_dll)
    except Exception as e:
        raise SystemExit(f"레지스트리 등록 실패: {e}")

    print(f"레지스트리 등록 완료:")
    print(f"  위치: HKEY_CURRENT_USER\\{REG_KEY_PATH}")
    print(f"  이름: {MODULE_NAME}")
    print(f"  값:   {registered}")
    print()
    print("설치가 끝났습니다. 이제 인쇄 스크립트를 실행하면 보안 승인 창이 뜨지 않습니다.")
    print("(만약 그래도 창이 뜨면, 한글을 완전히 종료한 뒤 다시 시도해 보세요)")


if __name__ == "__main__":
    main()