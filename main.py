import inventory_engine as engine
import scanner_interface as scanner
import qr_generator as gen # Assuming File 1 is named this

def show_menu():
    print("\n--- Reliance resQ SmartTrack System ---")
    print("1. Register New Part (Generate QR)")
    print("2. Scan Part IN (Stock Entry)")
    print("3. Scan Part OUT (Dispatch)")
    print("4. Exit")
    return input("Select Option: ")

def main():
    engine.initialize_db()
    
    while True:
        choice = show_menu()
        
        if choice == '1':
            p_id = input("Enter Part ID: ")
            p_name = input("Enter Part Name: ")
            gen.generate_part_qr(p_id, p_name) # Uses code from File 1
            
        elif choice in ['2', '3']:
            mode = "IN" if choice == '2' else "OUT"
            p_id, p_name = scanner.activate_scanner()
            
            if p_id:
                success, msg = engine.update_stock(p_id, p_name, mode)
                print(msg)
            else:
                print("⚠️ Scan Cancelled.")
                
        elif choice == '4':
            print("Shutting down resQ SmartTrack. Goodbye!")
            break
        else:
            print("Invalid Choice!")

if __name__ == "__main__":
    main()