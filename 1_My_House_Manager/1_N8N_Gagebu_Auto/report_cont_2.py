import pandas as pd
import os

# 1. 경로 설정
base_dir = os.path.dirname(os.path.abspath(__file__))
input_path = os.path.join(base_dir, "A_CONTRACT.xlsx")
parquet_path = os.path.join(base_dir, "merged_gagebu.parquet")

# 제외할 가짜 이름 키워드
exclude_keywords = ["퇴실자", "미확인", "입금자", "비어있음"]

# 2. 건물 및 호실 설정
room_config = {
    "봉명동": [101, 102, 103, 104, 105, 106, 201, 202, 203, 204, 205, 206, 301, 302, 303, 304, 305, 306],
    "신부동": [101, 201, 202, 203, 204, 205, 206, 301, 302, 303, 304, 305, 306, 401, 402, 403, 404, 405, 406, 501, 502, 503, 504, 505, 506],
    "쌍용동": [101, 102, 103, 201, 202, 203],
}

def create_individual_room_reports():
    try:
        # 데이터 로드
        contract_df = pd.read_excel(input_path)
        gagebu_df = pd.read_parquet(parquet_path)
        
        content_col = gagebu_df.columns[4]  # '내용' 열
        date_col = gagebu_df.columns[0]     # '날짜' 열
        gagebu_df[date_col] = pd.to_datetime(gagebu_df[date_col])

        # [핵심] 정확한 매칭을 위해 가계부 '내용'에서 이름만 미리 분리 (예: '홍길동-월세' -> '홍길동')
        # 하이픈(-)이 없으면 전체 텍스트를 이름으로 인식합니다.
        gagebu_df['temp_name'] = gagebu_df[content_col].str.split('-').str[0].str.strip()

        for building, rooms in room_config.items():
            building_file = os.path.join(base_dir, f"{building}_입금내역.xlsx")
            
            with pd.ExcelWriter(building_file, engine='openpyxl') as writer:
                for room in rooms:
                    # 1. 해당 호실의 모든 역대 임차인 이름 리스트업
                    tenants = contract_df[
                        (contract_df['건물명'] == building) & 
                        (contract_df['호실'] == room)
                    ]['임차인'].dropna().unique()
                    
                    # 2. 제외 키워드 필터링 및 공백 제거
                    clean_names = [str(n).strip() for n in tenants if not any(ex in str(n) for ex in exclude_keywords)]
                    
                    if not clean_names:
                        pd.DataFrame(columns=gagebu_df.columns[:-1]).to_excel(writer, sheet_name=f"{room}호", index=False)
                        continue

                    # 3. [VLOOKUP 방식] 'temp_name'이 'clean_names' 리스트에 정확히 포함된 행만 추출
                    room_history = gagebu_df[
                        (gagebu_df['분류'].str.contains(building, na=False)) &
                        (gagebu_df['temp_name'].isin(clean_names))
                    ].copy()

                    # 4. 날짜 내림차순 정렬 및 불필요한 임시 열 삭제 후 저장
                    room_history = room_history.sort_values(by=date_col, ascending=False)
                    room_history.drop(columns=['temp_name']).to_excel(writer, sheet_name=f"{room}호", index=False)
                
            print(f"[{building}] 정확한 이름 매칭으로 파일 생성 완료")

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    create_individual_room_reports()