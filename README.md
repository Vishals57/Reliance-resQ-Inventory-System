📦 Reliance resQ Enterprise Inventory System v5.7
A professional-grade, dark-mode Desktop ERP application designed for high-pressure service center environments. This system automates the tracking of serialized spare parts, manages engineer assignments, and generates physical QR assets for warehouse management.

🚀 Key Features
🖥️ Modern Admin UI
CustomTkinter Interface: Built with a Windows 11-style dark theme for a premium user experience.

Master-Detail Dashboard: Real-time synchronization between Master Stock and Live Transactions.

Scrollable Engineer Panel: Optimized for large teams (supports 21+ engineers) with a seamless selection grid.

🔍 Operational Intelligence
Live QR Scanning: Integrated webcam scanning using OpenCV for instant article identification.

Dual-Mode Movement: Fast toggling between Company (OG) and Local purchase types to track stock sources.

Shelf-Logic Tracking: Automatic "Inward/Outward" status management with timestamped logging.

🏷️ Asset Studio
QR Generator: Instant generation of high-fidelity QR codes for every registered article.

Live Preview & Print: View generated QR codes directly in the app and open them with a single click for physical printing.

Financial Tracking: Native support for Cost Price (CP) and Selling Price (SP) to monitor revenue and stock valuation.

🛠️ Tech Stack
Language: Python 3.11+

GUI Framework: CustomTkinter (Modernized Tkinter)

Database: Excel-based ERP Logic (pandas, openpyxl)

Computer Vision: OpenCV & pyzbar (for QR detection)

Asset Generation: qrcode & Pillow (PIL)

📁 Project Structure
Plaintext
├── gui_app.py            # Main Executive UI (v5.7)
├── inventory_engine.py   # Backend logic & Excel database handler
├── scanner_interface.py  # Webcam & QR scanning logic
├── qr_generator.py       # QR asset creation engine
├── Article_QRs/          # Auto-generated QR image storage
└── resQ_Enterprise_Inventory.xlsx  # Centralized Data Ledger
⚙️ Installation & Setup
Clone the repository:

Bash
git clone https://github.com/Vishals57/Reliance-resQ-Inventory-System.git
cd Reliance-resQ-Inventory-System
Install Dependencies:

Bash
pip install customtkinter pandas openpyxl opencv-python pillow qrcode pyzbar
Launch the Application:

Bash
python gui_app.py
📖 How to Use
Register: Go to the Asset Manager tab, enter the article details (ID, Name, CP, SP), and hit "Save."

Print: Click the Open to Print button to get the QR sticker for the product box.

Inward: Scan the QR under the Operations tab and select "Inward" to add it to your live stock.

Issue (Outward): Scan the part, select the Engineer from the scrollable list, and hit "Outward."

Track: Monitor all movements and charges in the Dashboard tab.

👨‍💻 Developed By
Vishal Sunil Shinde | Software Developer & Data Analyst at resQ

## Why this project matters
This isn't just a school project; it was developed to solve real-world inventory bottlenecks at a high-volume service center. It demonstrates proficiency in Full-Stack Desktop Development, Data Management, and UX Design.
