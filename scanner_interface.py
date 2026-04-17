import cv2
from pyzbar.pyzbar import decode

def activate_scanner():
    # Windows: CAP_DSHOW usually opens faster than default backend
    try:
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
    except Exception:
        cap = cv2.VideoCapture(0)

    # Reduce startup latency and decoding workload
    try:
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
        cap.set(cv2.CAP_PROP_FPS, 30)
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
    except Exception:
        pass

    # Warm-up a few frames so first decode isn't delayed
    for _ in range(10):
        try:
            cap.read()
        except Exception:
            break

    print("📷 Scanner Active. Align QR code. Press 'q' to cancel.")
    
    while True:
        ret, frame = cap.read()
        if not ret:
            continue

        for barcode in decode(frame):
            # 1. Get raw text from QR
            raw_data = barcode.data.decode('utf-8').strip()
            print(f"DEBUG: Scanned raw text -> '{raw_data}'") 
            
            cap.release()
            cv2.destroyAllWindows()
            
            # 2. Logic to clean the ID
            try:
                # Multi-line payload (Box/Unit QR): find ART_CODE line
                if "ART_CODE:" in raw_data:
                    for line in raw_data.splitlines():
                        if line.strip().startswith("ART_CODE:"):
                            clean_id = line.split(":", 1)[1].strip()
                            return clean_id, "Linked_Part"

                if "ART_NO:" in raw_data:
                    # If QR is 'ART_NO:001', this gives us '001'
                    clean_id = raw_data.split(":")[1].strip()
                    print(f"✅ Cleaned Article ID: {clean_id}")
                    return clean_id, "Linked_Part"
                
                elif "|" in raw_data:
                    # Logic for old format 'ID:123|Name:panel'
                    parts = raw_data.split("|")
                    clean_id = parts[0].split(":")[1].strip()
                    return clean_id, "Linked_Part"
                
                else:
                    # If QR is just raw text like '001'
                    return raw_data, "Direct_ID"
                    
            except Exception as e:
                print(f"❌ Error decoding QR format: {e}")
                return None, None
            
        cv2.imshow('resQ Smart Scanner', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
            
    cap.release()
    cv2.destroyAllWindows()
    return None, None