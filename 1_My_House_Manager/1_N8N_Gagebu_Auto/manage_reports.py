#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
manage_reports.py
통합 실행 스크립트
- GAGEBU.py (병합 -> merged_gagebu.parquet)
- cont.py (템플릿 생성 -> 건물별_거주현황_명부_template.xlsx)
- report_cont.py (정산 리포트 생성 -> 건물별_거주현황_명부.xlsx)
- report_cont_2.py (호실별 입금내역 생성)
- report_cont_3.py (미확인 입금내역 생성)

각 단계 사이에 대기와 딜레이를 두어 파일 충돌/잠금 문제를 줄입니다.
이 스크립트는 현재 사용 중인 Python 인터프리터(`sys.executable`)로 하위 스크립트를 실행합니다.
"""

from pathlib import Path
import subprocess
import sys
import time
import shutil
import argparse
import re

BASE_DIR = Path(__file__).resolve().parent

SCRIPTS = {
    'merge': {
        'script': 'GAGEBU.py',
        'outputs': ['merged_gagebu.parquet']
    },
    'template': {
        'script': 'cont.py',
        'outputs': ['건물별_거주현황_명부.xlsx'],
        # template 단계는 원본 파일을 보존하기 위해 이름을 바꿉니다
        'rename': {'건물별_거주현황_명부.xlsx': '건물별_거주현황_명부_template.xlsx'}
    },
    'report': {
        'script': 'report_cont.py',
        'outputs': ['건물별_거주현황_명부.xlsx']
    },
    'per-room': {
        'script': 'report_cont_2.py',
        'outputs': ['봉명동_입금내역.xlsx', '신부동_입금내역.xlsx', '쌍용동_입금내역.xlsx']
    },
    'unidentified': {
        'script': 'report_cont_3.py',
        'outputs': ['미확인_입금내역.xlsx']
    }
}


def build_tree_text(root_dir):
    root = Path(root_dir)

    def sort_key(path_obj):
        return (path_obj.is_file(), path_obj.name.lower())

    def walk(current, prefix=""):
        items = sorted(list(current.iterdir()), key=sort_key)
        lines = []
        for index, item in enumerate(items):
            is_last = index == len(items) - 1
            branch = "└── " if is_last else "├── "
            if item.is_dir():
                lines.append(f"{prefix}{branch}{item.name}/")
                child_prefix = prefix + ("    " if is_last else "│   ")
                lines.extend(walk(item, child_prefix))
            else:
                lines.append(f"{prefix}{branch}{item.name}")
        return lines

    tree_lines = [f"{root.name}/"]
    tree_lines.extend(walk(root))
    return "\n".join(tree_lines)


def update_readme_tree(root_dir, readme_path):
    readme = Path(readme_path)
    section = f"# 파일 구조\n\n```text\n{build_tree_text(root_dir)}\n```\n"

    if readme.exists():
        content = readme.read_text(encoding='utf-8')
    else:
        content = ""

    pattern = r"(?ms)^# 파일 구조\s*\n```text\n.*?\n```\s*"
    if re.search(pattern, content):
        updated = re.sub(pattern, section + "\n", content, count=1)
    else:
        updated = section if not content.strip() else section + "\n" + content

    readme.write_text(updated, encoding='utf-8')
    print(f"[OK] README 트리 갱신 완료: {readme}")


def run_script(step_key, python_exe, cwd, retries=1, continue_on_error=False):
    entry = SCRIPTS[step_key]
    script_path = Path(cwd) / entry['script']
    if not script_path.exists():
        msg = f"스크립트 없음: {script_path}"
        print("[WARN]", msg)
        if continue_on_error:
            return False
        raise FileNotFoundError(msg)

    cmd = [python_exe, str(script_path)]
    for attempt in range(1, retries + 1):
        try:
            print(f"[RUN] {script_path.name} (시도 {attempt}/{retries})")
            subprocess.run(cmd, cwd=str(cwd), check=True)
            return True
        except subprocess.CalledProcessError as e:
            print(f"[ERROR] {script_path.name} 실행 실패 (exit {e.returncode})")
            if attempt == retries:
                if continue_on_error:
                    return False
                raise
            time.sleep(2)
        except Exception as e:
            print(f"[ERROR] {script_path.name} 실행 중 예외: {e}")
            if continue_on_error:
                return False
            raise


def wait_for_outputs(outputs, cwd, timeout, poll_interval=1.0):
    paths = [Path(cwd) / o for o in outputs]
    print(f"[WAIT] 출력 대기: {', '.join(p.name for p in paths)} (타임아웃 {timeout}s)")
    start = time.time()
    last_missing = None
    while time.time() - start < timeout:
        missing = []
        for p in paths:
            if not p.exists():
                missing.append(p)
                continue
            # 파일이 생성 중일 수 있으므로 간단히 읽기 테스트를 합니다.
            try:
                with p.open('rb') as f:
                    f.read(1)
            except Exception:
                missing.append(p)

        if not missing:
            print(f"[OK] 출력 준비됨: {', '.join(p.name for p in paths)}")
            return True

        if last_missing != [m.name for m in missing]:
            print(f"[WAIT] 아직 준비 안됨: {', '.join(m.name for m in missing)}")
            last_missing = [m.name for m in missing]

        time.sleep(poll_interval)

    print(f"[TIMEOUT] 출력 대기 타임아웃({timeout}s). 누락: {', '.join(m.name for m in missing)}")
    return False


def _cleanup_old_baks(dst_path, keep=3):
    """백업 파일(.bak.*.xlsx)을 최신 keep개만 남기고 삭제합니다."""
    stem = dst_path.stem  # 예: 건물별_거주현황_명부_template
    parent = dst_path.parent
    bak_files = sorted(
        parent.glob(f"{stem}.bak.*.{dst_path.suffix.lstrip('.')}"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    for old in bak_files[keep:]:
        old.unlink()
        print(f"[CLEAN] 오래된 백업 삭제: {old.name}")


def rename_outputs(rename_map, cwd):
    for src, dst in rename_map.items():
        src_path = Path(cwd) / src
        dst_path = Path(cwd) / dst
        if src_path.exists():
            # 만약 대상이 이미 존재하면 백업 처리
            if dst_path.exists():
                bak = dst_path.with_name(dst_path.stem + f".bak.{int(time.time())}" + dst_path.suffix)
                shutil.move(str(dst_path), str(bak))
                print(f"[MOVE] 기존 {dst} -> {bak.name}")
                # 최근 3개만 유지하고 나머지 삭제
                _cleanup_old_baks(dst_path, keep=3)
            shutil.move(str(src_path), str(dst_path))
            print(f"[RENAME] {src} -> {dst}")
        else:
            print(f"[WARN] rename 대상 없음: {src}")


def parse_args():
    p = argparse.ArgumentParser(description="통합 리포트 실행기 — 순차 실행/대기/딜레이 포함")
    p.add_argument('--all', action='store_true', help='모든 단계 실행 (merge, template, report, per-room, unidentified)')
    p.add_argument('--steps', nargs='+', choices=list(SCRIPTS.keys()), help='실행할 단계 지정')
    p.add_argument('--delay', type=float, default=2.0, help='단계 간 대기(초)')
    p.add_argument('--timeout', type=int, default=60, help='각 단계 출력 대기 타임아웃(초)')
    p.add_argument('--poll', type=float, default=1.0, help='출력 준비 폴링 간격(초)')
    p.add_argument('--retries', type=int, default=1, help='스크립트 실행 재시도 횟수')
    p.add_argument('--continue-on-error', action='store_true', help='에러 발생시 다음 단계로 계속 진행')
    p.add_argument('--dry-run', action='store_true', help='실행 예정 단계만 출력')
    p.add_argument('--update-readme-tree', action='store_true', help='readme.md의 파일 구조 트리를 자동 갱신')
    p.add_argument('--readme-path', type=str, default=str(BASE_DIR / 'readme.md'), help='트리를 갱신할 README 파일 경로')
    return p.parse_args()


def main():
    args = parse_args()

    if args.update_readme_tree and not args.all and not args.steps:
        update_readme_tree(BASE_DIR, Path(args.readme_path))
        return

    if args.all or not args.steps:
        steps = ['merge', 'template', 'report', 'per-room', 'unidentified']
    else:
        steps = args.steps

    print(f"[INFO] 작업 디렉터리: {BASE_DIR}")
    for step in steps:
        entry = SCRIPTS[step]
        script_name = entry['script']

        if args.dry_run:
            print(f"[DRY] {step}: {script_name} -> 기대 출력: {entry.get('outputs')}")
            continue

        try:
            ok = run_script(step, sys.executable, BASE_DIR, retries=args.retries, continue_on_error=args.continue_on_error)
        except Exception as e:
            print(f"[ERROR] {script_name} 실행 중 예외: {e}")
            if not args.continue_on_error:
                raise
            ok = False

        if not ok and not args.continue_on_error:
            print(f"[FAIL] {step} 실패로 중단합니다.")
            return

        outputs = entry.get('outputs', [])
        if outputs:
            ready = wait_for_outputs(outputs, BASE_DIR, timeout=args.timeout, poll_interval=args.poll)
            if not ready and not args.continue_on_error:
                print(f"[FAIL] {step} 출력 준비 실패로 중단합니다.")
                return

        # template 단계에 한해 원본 파일을 템플릿 이름으로 보존
        if 'rename' in entry and entry['rename']:
            rename_outputs(entry['rename'], BASE_DIR)

        print(f"[SLEEP] {args.delay}s 대기...")
        time.sleep(args.delay)

    if args.update_readme_tree:
        update_readme_tree(BASE_DIR, Path(args.readme_path))

    print('[DONE] 선택된 모든 단계 완료')


if __name__ == '__main__':
    main()
