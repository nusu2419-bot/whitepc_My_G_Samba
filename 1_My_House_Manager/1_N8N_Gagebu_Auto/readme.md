# 2026-3-17 최종
crontab -e : 자동실행 설정 매일 01 시
로그 확인: 같은 폴더에 생긴 auto_log.log 파일. 실행 시간과 결과가 기록

- 업데이트 예정:  더 보기 좋게 입금내역을 Html 형식으로 변
- 메모장, 스케쥴 , 문자발송과 연동 




# 파일 구조

```text
1_N8N_Gagebu_Auto/
├── Excel_Gagebu/
│   ├── 2015-07-15_2026-01-17.xlsx
│   ├── 2026-01-01_2026-01-31.xlsx
│   ├── 2026-02-01_2026-02-28.xlsx
│   └── 2026-03-01_2026-03-31.xlsx
├── A_CONTRACT.xlsx
├── GAGEBU.py
├── cont.py
├── manage_reports.py
├── n8n 가계부_주택관리자동화.gdoc
├── readme.md
├── report_cont.py
├── report_cont_2.py
└── report_cont_3.py
```


최종 실행 파일
/mnt/photos/My_G_Samba/1_My_House_Manager/1_N8N_Gagebu_Auto/manage_reports.py

---

# 프로젝트 개요

주택 임대 관리를 위한 입금내역 자동 정리 자동화 시스템입니다.
은행 가계부 엑셀 파일과 임차인 계약 정보를 결합하여, 건물·호실별 정산 리포트와 미확인 입금 내역을 자동으로 생성합니다.

>> 사용자 입력내용 : Excel_Gagebu 폴더안에 매일 입금내역 파일을 업로드한다. , /mnt/photos/My_G_Samba/1_My_House_Manager/1_N8N_Gagebu_Auto/A_CONTRACT.xlsx 계약내용 수정 또는 추가 될경우 변경내용 입력한다. 이 두가지를 입력하면 계산 하고 보고서형식으로 사용자가 보기좋게 만드는것이 목표임

---

# 관리 대상 건물

| 건물 | 호실 구성 | 총 호실 수 |
|------|----------|-----------|
| 봉명동 | 101~106, 201~206, 301~306 | 18호실 |
| 신부동 | 101, 201~206, 301~306, 401~406, 501~506 | 25호실 |
| 쌍용동 | 101~103, 201~203 | 6호실 |

---

# 입력 파일 설명

### `A_CONTRACT.xlsx` — 임차인 계약 정보
임차인의 계약 전체 이력을 관리하는 마스터 파일입니다.

| 열 | 설명 |
|----|------|
| 건물명 | 봉명동 / 신부동 / 쌍용동 |
| 호실 | 호실 번호 (예: 101) |
| 임차인 | 임차인 이름 (가계부 내용 열과 매칭 기준) |
| Phone | 연락처 |
| 보증금 | 만원 단위 |
| 월세 | 만원 단위 |
| 관리비 | 만원 단위 |
| 부가세 | 만원 단위 |
| 입주일 | 날짜 (거주월 계산 기준) |
| 상태 | `거주중` / `퇴실` 등 |

### `Excel_Gagebu/` — 은행 가계부 엑셀 파일 모음
은행에서 내보낸 입출금 내역 파일들을 기간별로 저장하는 폴더입니다.
파일명 형식: `YYYY-MM-DD_YYYY-MM-DD.xlsx`

주요 열 구조 (열 인덱스 기준):

| 열 인덱스 | 설명 |
|----------|------|
| 0 | 날짜 |
| 4 | 내용 (임차인 이름 포함, 예: `홍길동-월세`) |
| 5 | 금액 |
| 7 | 거래 고유값 (중복 체크 기준) |
| 분류 열 | 건물명 포함 여부로 해당 건물 입금 분류 |

---

# 스크립트별 역할

### `GAGEBU.py` — 1단계: 가계부 병합
- `Excel_Gagebu/` 폴더의 모든 `.xlsx` 파일을 하나로 병합
- 출력: `merged_gagebu.parquet` (이후 모든 단계의 기반 데이터)
- 부가: H열(인덱스 7) 기준 중복 행을 `duplicates_list.xlsx`로 추출

### `cont.py` — 2단계: 거주현황 템플릿 생성
- `A_CONTRACT.xlsx`에서 `상태 == '거주중'` 인 임차인만 필터링
- 건물별 전체 호실 틀에 맞춰 Left Join → 빈 호실은 공란 유지
- 출력: `건물별_거주현황_명부_template.xlsx` (정산 전 기본 명부)

### `report_cont.py` — 3단계: 정산 리포트 생성
- `A_CONTRACT.xlsx` + `merged_gagebu.parquet` 결합
- 임차인 이름으로 가계부 입금액 합산 (VLOOKUP 방식)
- 거주월 계산: 입주일 기준 선불 원칙 적용 (엑셀 수식 동일 로직)
- 출력: `건물별_거주현황_명부.xlsx`

| 생성 열 | 계산 방식 |
|--------|---------|
| 거주월 | 입주일 ~ 오늘 기준, 선불 원칙 |
| 받아야할금액 | `(보증금 + (월세+관리비+부가세) × 거주월) × 10000` |
| 받은금액 | 가계부에서 임차인 이름으로 합산 |
| 미수금 | `받은금액 - 받아야할금액` (음수 = 미납) |

### `report_cont_2.py` — 4단계: 호실별 입금내역 생성
- 건물별로 파일을 나누고, 각 호실을 시트로 구성
- 가계부 내용에서 이름 분리 (`홍길동-월세` → `홍길동`) 후 정확히 매칭
- 역대 임차인 전원의 입금 이력을 날짜 최신순으로 정리
- 출력: `봉명동_입금내역.xlsx`, `신부동_입금내역.xlsx`, `쌍용동_입금내역.xlsx`

### `report_cont_3.py` — 5단계: 미확인 입금내역 추출
- 가계부의 건물 관련 입금 중 `A_CONTRACT.xlsx` 임차인 이름과 **매칭되지 않는** 항목 추출
- 오입금, 이름 오기재, 신규 임차인 등 수동 확인이 필요한 항목 파악용
- 출력: `미확인_입금내역.xlsx`

### `manage_reports.py` — 통합 실행기
위 5개 스크립트를 순서대로 실행하며, 각 단계 완료 후 출력 파일 생성 여부를 확인하고 다음 단계로 진행합니다.

---

# 실행 흐름

```
Excel_Gagebu/*.xlsx
        │
        ▼
[1] GAGEBU.py         →  merged_gagebu.parquet
                                  │
A_CONTRACT.xlsx ─────────────────┤
        │                        │
        ▼                        ▼
[2] cont.py           →  건물별_거주현황_명부_template.xlsx
[3] report_cont.py    →  건물별_거주현황_명부.xlsx  (미수금 포함 정산표)
[4] report_cont_2.py  →  {건물명}_입금내역.xlsx     (호실별 입금 이력)
[5] report_cont_3.py  →  미확인_입금내역.xlsx       (매칭 불가 입금)
```

---

# 실행 방법

```bash
# 전체 실행 (인자 없이 실행하면 모든 단계 자동 실행)
python manage_reports.py

# 특정 단계만 실행
python manage_reports.py --steps merge
python manage_reports.py --steps template
python manage_reports.py --steps report
python manage_reports.py --steps per-room
python manage_reports.py --steps unidentified

# 실행 예정 단계 확인만 (실제 실행 안 함)
python manage_reports.py --dry-run

# README 파일 트리 자동 갱신
python manage_reports.py --update-readme-tree
```

---

# 실행 환경

- Python 환경: `conda activate Gagebu_Auto`
- 필수 패키지: `pandas`, `pyarrow`, `openpyxl`
- 설치 명령: `pip install pandas pyarrow openpyxl`


현재 모드에서는 파일을 직접 수정할 수 없어서, 아래 내용을 readme.md 하단에 붙여 넣으면 됩니다.

점검 결과 (2026-03-23)
점검 범위: 1_My_House_Manager 하위 Python 코드 전체

Windows C:\... 형식 경로: 발견되지 않음

경로 이슈 후보 1건:

backup_gagebu.py:8
base_path = "/mnt/photos/My_G_Samba/1_My_House_Manager/1_N8N_Gagebu_Auto" (절대경로 하드코딩)
나머지 주요 실행 스크립트는 __file__ 기준 상대경로 사용으로 이식성 양호:

report_cont.py:7-10
cont.py:5-7
GAGEBU.py:5-8
report_cont_2.py:5-7
report_cont_3.py:5-8
권장 조치: backup_gagebu.py도 현재 파일 위치 기준 상대경로 방식으로 변경 권장.

---

# 업데이트 내역 (2026-03-28)

### 1. GitHub 저장소 연동
- `/mnt/photos/My_G_Samba` 폴더를 Git 저장소로 초기화
- `.gitignore` 설정: `.py` 파일만 추적 (엑셀, 로그 등 제외)
- 원격 저장소: https://github.com/nusu2419-bot/whitepc_My_G_Samba

### 2. 엑셀 서식 보호 (값만 갱신)
- 기존 코드는 `to_excel()`로 시트를 통째로 덮어써 사용자 서식이 초기화되는 문제 있었음
- `excel_value_only_writer.py` 공통 모듈 신규 추가
- 기존 파일이 있으면 셀 **값만** 갱신하고 폰트·색상·열너비 등 서식은 유지
- 적용 파일: `GAGEBU.py`, `cont.py`, `report_cont.py`, `report_cont_2.py`, `report_cont_3.py`

### 3. 입력 경로 변경
- 가계부 엑셀 입력 폴더 변경
  - 변경 전: `1_N8N_Gagebu_Auto/Excel_Gagebu/`
  - 변경 후: `1_My_House_Manager/3_GageBu_Input/`
- 수정 파일: `GAGEBU.py`

### 4. 결과물 출력 경로 변경
- 모든 결과 엑셀 파일 출력 위치 변경
  - 변경 전: `1_N8N_Gagebu_Auto/` (스크립트 폴더 내)
  - 변경 후: `1_My_House_Manager/4_Report/`
- 해당 파일: `건물별_거주현황_명부.xlsx`, `봉명동/신부동/쌍용동_입금내역.xlsx`, `미확인_입금내역.xlsx`
- 수정 파일: `cont.py`, `report_cont.py`, `report_cont_2.py`, `report_cont_3.py`
- `manage_reports.py`의 `SCRIPTS` 출력 감시 경로도 `4_Report` 기준으로 수정

### 5. 템플릿 백업 파일 자동 정리
- `manage_reports.py`에 `_cleanup_old_baks()` 함수 추가
- `.bak.*.xlsx` 파일을 최신 **3개**만 유지하고 초과분 자동 삭제
- 기존에 쌓인 6개 → 3개로 즉시 정리

### 6. 로그 파일 타임스탬프 및 자동 정리
- `manage_reports.py` 실행 시 `auto_log.log` 에 날짜/시간 헤더 자동 기록
  ```
  ==================================================
  [START] 2026-03-28 08:53:31
  ==================================================
  ...
  [END]   2026-03-28 08:53:44
  ```
- `trim_log()` 함수 추가: 로그가 **200줄** 초과 시 오래된 항목 자동 삭제
- `try/finally`로 감싸 오류 발생 시에도 `[END]` 기록 및 trim 반드시 실행