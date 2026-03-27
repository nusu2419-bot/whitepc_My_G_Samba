import pandas as pd
import os
import calendar
from datetime import datetime

# 1. 경로 설정
base_dir = os.path.dirname(os.path.abspath(__file__))
input_path = os.path.join(base_dir, "A_CONTRACT.xlsx")
parquet_path = os.path.join(base_dir, "merged_gagebu.parquet")
output_path = os.path.join(base_dir, "건물별_거주현황_명부.xlsx")

# 2. 건물별 호실 설정
room_config = {
    "봉명동": [101, 102, 103, 104, 105, 106, 201, 202, 203, 204, 205, 206, 301, 302, 303, 304, 305, 306],
    "신부동": [101, 201, 202, 203, 204, 205, 206, 301, 302, 303, 304, 305, 306, 401, 402, 403, 404, 405, 406, 501, 502, 503, 504, 505, 506],
    "쌍용동": [101, 102, 103, 201, 202, 203],
}

def calc_excel_logic_months(start_date, today):
    """엑셀 수식 분리 계산 방식 적용 (선불 원칙)"""
    if pd.isnull(start_date) or start_date > today:
        return 0
    
    # 년/월 차이 계산
    months_diff = (today.year - start_date.year) * 12 + (today.month - start_date.month)
    
    # 이번 달의 마지막 날짜 확인
    last_day_of_month = calendar.monthrange(today.year, today.month)[1]
    # 입주일의 '일'과 이번 달 '말일' 중 작은 값으로 비교
    compare_day = min(start_date.day, last_day_of_month)
    
    # 오늘 날짜가 비교일보다 크거나 같으면 이번 달치 발생(+1)
    add_month = 1 if today.day >= compare_day else 0
    
    return months_diff + add_month

def create_management_sheet():
    try:
        df = pd.read_excel(input_path)
        gagebu_df = pd.read_parquet(parquet_path)
        
        # 가계부 열 설정 (E열:내용, F열:금액)
        name_col = gagebu_df.columns[4]
        amt_col = gagebu_df.columns[5]
        
        active_residents = df[df['상태'] == '거주중'].copy()
        today = datetime.now()

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for building, rooms in room_config.items():
                base_df = pd.DataFrame({'건물명': building, '호실': rooms})
                building_data = active_residents[active_residents['건물명'] == building]
                merged_df = pd.merge(base_df, building_data, on=['건물명', '호실'], how='left')

                # 1. 거주월 계산
                merged_df['입주일'] = pd.to_datetime(merged_df['입주일'], errors='coerce')
                merged_df['거주월'] = merged_df['입주일'].apply(lambda x: calc_excel_logic_months(x, today))

                # 2. 받아야할 금액 (보증금 + (월세+관리비+부가세)*거주월) * 10000
                for col in ['보증금', '월세', '관리비', '부가세']:
                    merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce').fillna(0)
                
                merged_df['받아야할금액'] = (merged_df['보증금'] + 
                                       (merged_df['월세'] + merged_df['관리비'] + merged_df['부가세']) * merged_df['거주월']) * 10000

                # 3. 받은 금액 (가계부 이름 매칭)
                def sum_paid(name):
                    if pd.isnull(name) or str(name).strip() == "": return 0
                    # 이름이 포함된 행만 필터링하여 합산
                    return gagebu_df[gagebu_df[name_col].str.contains(str(name), na=False)][amt_col].sum()

                merged_df['받은금액'] = merged_df['임차인'].apply(sum_paid)

                # 4. 미수금 계산 (받은금액 L - 받아야할금액 K)
                merged_df['미수금'] = merged_df['받은금액'] - merged_df['받아야할금액']

                # 결과 정리
                merged_df['입주일'] = merged_df['입주일'].dt.strftime('%Y-%m-%d')
                cols = ['건물명', '호실', '임차인', 'Phone', '보증금', '월세', '관리비', '부가세', '입주일', '거주월', '받아야할금액', '받은금액', '미수금']
                
                final_df = merged_df[cols].sort_values(by='호실')
                final_df.to_excel(writer, sheet_name=building, index=False)
                print(f"[{building}] 정산 데이터 생성 완료")

        print(f"\n작업 완료: {output_path}")

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    create_management_sheet()