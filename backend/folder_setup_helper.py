import os

folders = [
    r"C:\Users\sailo\Desktop\face_attendance_system\backend\embeddings",
    r"C:\Users\sailo\Desktop\face_attendance_system\backend\reports\monthly",
    r"C:\Users\sailo\Desktop\face_attendance_system\backend\reports\yearly",
]

for folder in folders:
    try:
        os.makedirs(folder, exist_ok=True)
        print(f"[INFO] Folder ready: {folder}")
    except PermissionError:
        print(f"[ERROR] Permission denied creating folder: {folder}")
    except Exception as e:
        print(f"[ERROR] Failed to create folder {folder}: {e}")
