🚜 Seed & Fertilizer Calibration Tool
A professional-grade desktop application designed for the precise calibration of agricultural seeding and fertilization machinery. Developed as part of an academic thesis at Akdeniz University, Faculty of Agriculture, this tool integrates mechanical engineering formulas with a modern GUI to optimize field efficiency and planting accuracy.

🔬 Scientific & Engineering Background
The application digitizes traditional calibration methods used in Agricultural Mechanization. It provides high-precision calculations for:

1. Seeding Calibration (Grain & Precision)
Static Seeding Test: Calculates the required seed mass for 20 wheel rotations based on the decare norm (kg/da).

Formula: Q 
20
​
 =Q×0.063×D×B

Precision Seeding Quality (Singulation): Analyzes row-spacing distribution to determine Multiples, Skips, and Acceptable seeding rates based on theoretical target distance.

Quality Thresholds: Multiples (x<0.5a), Skips (x>1.5a).

2. Fertilization Norms
Volumetric Flow Analysis: Calculates the hourly capacity (Q 
h
​
 ) of centrifugal or box spreaders by integrating feeding speed, gate height, and fertilizer bulk density (λ).

Application Norm (q): Converts hourly flow into field application rates (kg/da) based on working width and tractor ground speed.

3. Field Performance & Logistics
Work Performance: Calculates effective field capacity (da/h) considering field efficiency coefficients.

Marker Lengths: Geometric calculation for tractor alignment to ensure zero-overlap and zero-gap seeding.

💻 Technical Stack & Features
Architecture
GUI Framework: ttkbootstrap for a modern, responsive, and theme-able interface.

Data Management: Persistent storage using JSON-based configuration for user preferences and calculation history.

Data Validation: Real-time input sanitization using Tkinter's validatecommand to ensure numerical integrity.

OOP Principles: Utilization of dataclasses and modular class structures for scalable computation logic.

Reporting & Export
Professional PDF Export: Powered by reportlab, featuring custom UTF-8 (DejaVuSans) font integration to support multi-language characters and academic formatting.

Excel Integration: Built with openpyxl and pandas, exporting styled spreadsheets with descriptive metadata for field records.

History Tracking: A built-in SQLite-style history log (stored in JSON) to review the last 5 calculations per module.

🛠 Installation & Usage
Prerequisites
Python 3.8+

DejaVuSans.ttf font file (included in root for PDF rendering)

Setup
Clone the repository:

Bash
git clone https://github.com/yourusername/seed-calibration-tool.git
cd seed-calibration-tool
Install dependencies:

Bash
pip install ttkbootstrap reportlab pandas openpyxl matplotlib
Run the application:

Bash
python main.py
📊 Application Modules
Grain Seeder: 20-rotation test and field-norm verification.

Precision Seeder: Row spacing analysis and singulation quality check.

Fertilizer Spreader: Flow rate and distribution norm calibration.

Agricultural Logistics: Marker length, germination rate, and work success calculations.

🎓 Academic Credit
Author: Hasan Dural

Advisor: Prof. Dr. Davut Karayel

Institution: Akdeniz University, Faculty of Agriculture, Department of Agricultural Machinery and Technologies Engineering.

📜 License
This project is licensed under the MIT License. See the LICENSE file for details.
