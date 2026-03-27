import os
import zipfile
import datetime
import time

def backup_gagebu_auto():
    # 대상 경로 설정
    base_path = "/mnt/photos/My_G_Samba/1_My_House_Manager/1_N8N_Gagebu_Auto"
    backup_dir = os.path.join(base_path, "Backup")
    
    # 1. Backup 폴더 생성
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
        print(f"Backup 폴더 생성됨: {backup_dir}")

    # 2. 3개월(90일) 이상 된 구버전 백업 삭제
    now_time = time.time()
    retention_sec = 90 * 86400
    if os.path.exists(backup_dir):
        for file in os.listdir(backup_dir):
            file_path = os.path.join(backup_dir, file)
            if os.path.isfile(file_path) and file.endswith(".zip"):
                if (now_time - os.path.getmtime(file_path)) > retention_sec:
                    os.remove(file_path)
                    print(f"삭제된 노후 백업: {file}")

    # 3. 날짜/시간을 맨 앞으로 배치 (예: 20260323_1630_backup.zip)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    zip_name = f"{timestamp}_backup.zip"  # 이 부분을 수정했습니다!
    zip_path = os.path.join(backup_dir, zip_name)

    # 4. 압축 실행
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(base_path):
            # Backup 폴더 자체는 압축 대상에서 제외 (무한 루프 방지)
            if "Backup" in root.split(os.sep):
                continue
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, base_path)
                zipf.write(full_path, rel_path)

    print(f"백업 완료: {zip_path}")

if __name__ == "__main__":
    backup_gagebu_auto()