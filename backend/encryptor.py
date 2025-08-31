import os
import pickle
from cryptography.fernet import Fernet

KEY_FILE = r"C:\Users\sailo\Desktop\face_attendance_system\backend\embeddings\secret.key"

def generate_key():
    key_dir = os.path.dirname(KEY_FILE)
    os.makedirs(key_dir, exist_ok=True)
    key = Fernet.generate_key()
    with open(KEY_FILE, "wb") as f:
        f.write(key)
    print("[INFO] Fernet key generated and saved.")

def load_key():
    if not os.path.exists(KEY_FILE):
        print("[ERROR] Fernet key not found. You must copy secret.key from admin PC.")
        raise PermissionError("You must copy secret.key from admin PC before running this.")
    with open(KEY_FILE, "rb") as f:
        return f.read()

def encrypt_embedding(data):
    key = load_key()
    f = Fernet(key)
    pickled = pickle.dumps(data)
    return f.encrypt(pickled)

def decrypt_embedding(encrypted):
    key = load_key()
    f = Fernet(key)
    return pickle.loads(f.decrypt(encrypted))

# Only run on PC1 (admin) to generate key initially
if __name__ == "__main__":
    if not os.path.exists(KEY_FILE):
        generate_key()
    else:
        print("[INFO] Key already exists. No action taken.")
