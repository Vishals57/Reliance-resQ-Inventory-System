import inventory_engine as engine
import scanner_interface as scanner
import qr_generator as gen 

def main():
    engine.initialize_db()
    
    while True:
        print("\n--- resQ Store Management System ---")
        print("1. Register New Article (Master Data)")
        print("2. Inward Purchase (GRN - Local/Company)")
        print("3. Outward to Engineer (Job Issue)")
        print("4. Exit")
        
        # THIS LINE MUST BE HERE - Defining 'choice' before the IF statements
        choice = input("Select Option (1-4): ")

        if choice == '1':
            art_no = input("Article No: ")
            name = input("Part Name: ")
            cp = float(input("Cost Price: "))
            sp = float(input("Selling Price: "))
            cat = input("Category (OG/Local): ")
            
            # Step A: Save to Excel
            success, msg = engine.register_new_article(art_no, name, cp, sp, cat)
            print(msg)
            
            # Step B: Generate the Physical QR Code (single QR per Article_No)
            if success:
                gen.generate_box_qr(art_no, name, "-", sp, 1, sr_list=None)

        elif choice == '2':
            # Scan the QR to get the Article Number
            art_no, _ = scanner.activate_scanner()
            if art_no:
                print(f"📦 Scanned Article: {art_no}")
                p_type = input("Purchase Type (1 for Company / 2 for Local): ")
                p_label = "Company" if p_type == '1' else "Local"
                _, msg = engine.process_movement(art_no, "IN", purchase_type=p_label)
                print(msg)
            else:
                print("⚠️ Scan Cancelled.")

        elif choice == '3':
            # Scan the QR to see which part is going out
            art_no, _ = scanner.activate_scanner()
            if art_no:
                print(f"📦 Scanned Article: {art_no}")
                print("\nSelect Engineer:")
                for i, eng in enumerate(engine.ENGINEERS):
                    print(f"{i+1}. {eng}")
                
                try:
                    eng_idx = int(input("Choice: ")) - 1
                    engineer_name = engine.ENGINEERS[eng_idx]
                    _, msg = engine.process_movement(art_no, "OUT", engineer=engineer_name)
                    print(msg)
                except (ValueError, IndexError):
                    print("❌ Invalid Engineer Selection!")
            else:
                print("⚠️ Scan Cancelled.")

        elif choice == '4':
            print("Shutting down resQ SmartTrack. Goodbye!")
            break
        
        else:
            print("❌ Invalid Choice! Please enter 1, 2, 3, or 4.")

if __name__ == "__main__":
    main()