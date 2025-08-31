import face_recognition
import cv2
import os
from encryptor import encrypt_embedding

BASE_DIR = r"C:\Users\sailo\Desktop\face_attendance_system\backend\embeddings"

def register_face(name, role, classroom):
    folder = os.path.join(BASE_DIR, classroom)
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, f"{name}_{role}.bin")

    if os.path.exists(file_path):
        print(f"[INFO] {name} ({role}) is already registered in classroom {classroom}.")
        return

    cap = cv2.VideoCapture(0)
    print(f"[INFO] Please position your face in front of the camera for registration.")

    registered = False
    while True:
        ret, frame = cap.read()
        if not ret:
            print("[ERROR] Failed to grab frame from camera.")
            break

        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        boxes = face_recognition.face_locations(rgb_frame)

        if not registered and len(boxes) == 1:
            encodings = face_recognition.face_encodings(rgb_frame, boxes)
            if encodings:
                data = {
                    "name": name,
                    "role": role,
                    "classroom": classroom,
                    "encoding": encodings[0]
                }
                try:
                    encrypted = encrypt_embedding(data)
                except PermissionError as e:
                    print("[ERROR]", str(e))
                    print("[INSTRUCTION] Copy secret.key file from admin PC to this PC's embeddings folder.")
                    cap.release()
                    cv2.destroyAllWindows()
                    return
                with open(file_path, "wb") as f:
                    f.write(encrypted)
                print(f"[SUCCESS] Face registered and saved to {file_path}")
                registered = True
                cv2.imshow("Register Face", frame)
                cv2.waitKey(1000)
                break
            else:
                cv2.putText(frame, "Failed to encode face, please try again", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
        elif not registered:
            if len(boxes) > 1:
                cv2.putText(frame, "Multiple faces detected - show only your face!", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
            elif len(boxes) == 0:
                cv2.putText(frame, "No face detected - please position your face", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)
            else:
                cv2.putText(frame, "Detecting face...", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 0), 2)

        cv2.imshow("Register Face", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            print("[INFO] Registration cancelled by user.")
            break

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    name = input("Enter Name: ").strip()
    role = input("Enter Role (student/faculty): ").lower().strip()
    classroom = input("Enter Classroom (5EP2/5EP3): ").upper().strip()
    register_face(name, role, classroom)
