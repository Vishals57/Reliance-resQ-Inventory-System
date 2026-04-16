import cv2
from pyzbar.pyzbar import decode

def activate_scanner():
    cap = cv2.VideoCapture(0)
    print("📷 Scanner Active. Align QR code. Press 'q' to cancel.")
    
    while True:
        ret, frame = cap.read()
        for barcode in decode(frame):
            raw_data = barcode.data.decode('utf-8').strip()
            print(f"DEBUG: Scanned raw text -> '{raw_data}'") # This tells us exactly what's in the QR
            
            cap.release()
            cv2.destroyAllWindows()
            
            # If your QR is "ID:1|Name:panel"
            if "|" in raw_data and ":" in raw_data:
                try:
                    parts = raw_data.split("|")
                    p_id = parts[0].split(":")[1].strip()
                    p_name = parts[1].split(":")[1].strip()
                    return p_id, p_name
                except Exception as e:
                    print(f"❌ Formatting Error: {e}")
                    return None, None
            else:
                # If your QR is JUST the ID (e.g., "1")
                # We will use the raw data as ID and ask for a generic name
                return raw_data, "Unidentified Part"
            
        cv2.imshow('resQ Smart Scanner', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
            
    cap.release()
    cv2.destroyAllWindows()
    return None, None