import pandas as pd
import os
from excel_value_only_writer import write_sheets_value_only

# 1. 경로 설정
base_dir = os.path.dirname(os.path.abspath(__file__))
input_path = os.path.join(base_dir, "A_CONTRACT.xlsx")
parquet_path = os.path.join(base_dir, "merged_gagebu.parquet")
_report_dir = "/mnt/photos/My_G_Samba/1_My_House_Manager/4_Report"
output_unidentified = os.path.join(_report_dir, "미확인_입금내역.xlsx")

# 제외할 키워드
exclude_keywords = ["퇴실자", "미확인", "입금자", "비어있음"]

def export_unidentified_deposits():
    try:
        # 데이터 로드
        contract_df = pd.read_excel(input_path)
        gagebu_df = pd.read_parquet(parquet_path)
        
        content_col = gagebu_df.columns[4]  # '내용' 열
        date_col = gagebu_df.columns[0]     # '날짜' 열
        gagebu_df[date_col] = pd.to_datetime(gagebu_df[date_col])

        # 1. 계약서의 모든 유효한 임차인 이름 집합 생성
        all_tenants = contract_df['임차인'].dropna().unique()
        valid_tenant_names = {str(n).strip() for n in all_tenants if not any(ex in str(n) for ex in exclude_keywords)}

        # 2. 가계부 내용에서 이름만 추출하여 임시 열 생성
        gagebu_df['temp_name'] = gagebu_df[content_col].str.split('-').str[0].str.strip()

        # 3. 임차인 목록에 이름이 없는 데이터만 필터링 (정확히 일치하지 않는 경우)
        # 건물 관련 분류(봉명동, 신부동, 쌍용동) 중에서 임차인 이름이 확인되지 않은 것들
        buildings_pattern = "봉명동|신부동|쌍용동"
        unidentified_df = gagebu_df[
            (gagebu_df['분류'].str.contains(buildings_pattern, na=False)) &
            (~gagebu_df['temp_name'].isin(valid_tenant_names))
        ].copy()

        # 4. 날짜 내림차순 정렬 (최신순)
        unidentified_df = unidentified_df.sort_values(by=date_col, ascending=False)

        # 5. 임시 열 삭제 후 엑셀 저장 (기존 파일이 있으면 서식은 유지하고 값만 갱신)
        write_sheets_value_only(output_unidentified, {"Sheet1": unidentified_df.drop(columns=['temp_name'])})
        
        print(f"미확인 입금 내역 추출 완료: {output_unidentified}")
        print(f"총 {len(unidentified_df)}건의 미확인 데이터가 발견되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    export_unidentified_deposits()