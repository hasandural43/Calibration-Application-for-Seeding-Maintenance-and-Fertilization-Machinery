<div align="center">

# 🚜 Agricultural Machinery Calibration & Engineering Suite

**A Professional Desktop Application for High-Precision Seeding and Fertilization Optimization**

[![Python Version](https://img.shields.io/badge/Python-3.8%2B-blue.svg?style=for-the-badge&logo=python)](#)
[![UI Framework](https://img.shields.io/badge/GUI-ttkbootstrap-2ea44f.svg?style=for-the-badge)](#)
[![License](https://img.shields.io/badge/License-MIT-purple.svg?style=for-the-badge)](#)
[![Institution](https://img.shields.io/badge/Akdeniz_University-Agriculture-orange.svg?style=for-the-badge)](#)

*Developed for rigorous agricultural mechanization analysis, bridging theoretical agronomic formulas with real-world field application.*

</div>

---

## 📑 Table of Contents
1. [Executive Summary](#1-executive-summary)
2. [Mathematical Models & Core Modules](#2-mathematical-models--core-modules)
3. [Precision, Accuracy, and Agronomic Standards](#3-precision-accuracy-and-agronomic-standards)
4. [Software Architecture & Tech Stack](#4-software-architecture--tech-stack)
5. [Step-by-Step User Guide](#5-step-by-step-user-guide)
6. [Installation & Deployment](#6-installation--deployment)
7. [Reporting & Data Export](#7-reporting--data-export)
8. [Academic Citation](#8-academic-citation)

---

## 1. Executive Summary

The **Agricultural Machinery Calibration Suite** is a desktop software engineered to optimize the performance of grain seeders, precision planters, and centrifugal fertilizer spreaders. Developed as part of an academic thesis at **Akdeniz University, Faculty of Agriculture**, the tool digitizes complex mechanization formulas to eliminate input waste, ensure optimal plant spacing, and generate professional compliance reports.

---

## 2. Mathematical Models & Core Modules

The software engine computes real-time agronomic data using the following deterministic engineering models:

### 2.1. Grain Seeder Calibration (Static & Field)
Calculates the exact mechanical delivery rate required to achieve a target decare norm.
* **Static 20-Rotation Formulation:**
  $$Q_{20} = Q \times 0.063 \times D \times B$$
  Where $Q_{20}$ is the seed mass for 20 rotations (kg), $Q$ is the target rate (kg/da), $D$ is wheel diameter (m), and $B$ is working width (m).
* **Field Consumption Rate:**
  $$Q = \frac{1000 \times q}{2 \times L \times B}$$
  Where $q$ is total seeds consumed (kg) over distance $L$ (m).

### 2.2. Precision Seeder Kinematics
Determines the theoretical intra-row spacing based on mechanical transmission.
* **Row Spacing ($a$):**
  $$a = \frac{\pi \times D}{i \times n}$$
  Where $i$ is the transmission ratio and $n$ is the number of holes on the vacuum disk.

### 2.3. Volumetric Fertilizer Spreaders
Synchronizes material discharge with tractor ground speed to prevent toxic over-application.
* **Hourly Discharge Capacity ($Q_h$):**
  $$Q_h = 0.06 \times V_b \times h \times b \times \lambda$$
  Where $V_b$ is feed rate (m/min), $h$ is gate height (m), $b$ is gate width (m), and $\lambda$ is bulk density (kg/m³).
* **Application Norm ($q$):**
  $$q = \left( \frac{Q_h}{W \times V} \right) \times 1000$$
  Where $W$ is working width (m) and $V$ is forward velocity (km/h).

### 2.4. Logistics & Agronomy
* **Work Performance:** Computes effective field capacity ($C = \frac{W \times V \times \eta}{10}$ da/h) utilizing field efficiency coefficients ($\eta$).
* **Marker Alignment:** Geometric calculation to prevent overlapping: $x_{right/left} = b \mp \frac{L}{2}$.
* **Biological Seed Rate:** Adjusts required seed mass utilizing Thousand Grain Weight (TGW), Purity (%), and Germination Rate (%).

---

## 3. Precision, Accuracy, and Agronomic Standards

This software operates with **64-bit floating-point precision** for all internal mathematical computations, ensuring zero arithmetic degradation.

### Seeding Quality (Singulation) Standards
The **Ekim Kalitesi (Seeding Quality)** module algorithmically parses arrays of field-measured seed spacings and evaluates them against global agronomic standards:
* **Multiples (İkizleme):** Distance $< 0.5a$ (Causes resource competition).
* **Skips (Boşluk):** Distance $> 1.5a$ (Causes yield gaps).
* **Acceptable (Kabul Edilebilir):** $0.5a \leq x \leq 1.5a$.

**Systematic Compliance Thresholds:**
The software utilizes strict boolean logic to issue a `SUITABLE ✅` or `NOT SUITABLE ❌` verdict based on selected crop tolerances:
1. **Row Crops (Corn, Cotton, Sunflower):** Requires **$\geq 90\%$** acceptable spacing.
2. **Vegetable Crops:** Requires **$\geq 85\%$** acceptable spacing.

---

## 4. Software Architecture & Tech Stack

* **Front-End:** `ttkbootstrap` (Modernized Tkinter implementation with Dark/Light theme support).
* **Validation Layer:** Real-time keystroke validation (`validatecommand`) blocking non-float/NaN injections.
* **Data Persistence:** Local `JSON` serialization storing the last 5 chronological calculation states per module.
* **Reporting Engines:** `ReportLab` (Vector-based PDF generation with dynamic tables) and `openpyxl` (Structured Excel data formatting).

---

## 5. Step-by-Step User Guide

### Step 1: Initialization & Module Selection
Launch the application. You will be greeted by the Main Dashboard. Select your target machinery category: **Tahıl (Grain)**, **Hassas (Precision)**, or **Gübre (Fertilizer)**.

> **Note:** Use the unit toggle at the top left to switch globally between Metric (m) and Centimetric (cm) inputs.

### Step 2: Parameter Injection
Navigate to the desired tab. Input your mechanical constraints (e.g., wheel diameter, transmission ratio). The system automatically sanitizes inputs.

### Step 3: Computation & Verification
Click **"Hesapla" (Calculate)**. The software processes the algorithmic matrix and outputs the exact operational norm. For precision quality, it calculates the statistical distribution of your inputted field measurements.

### Step 4: Export & Auditing
* Click **"PDF'e Yazdır"** to generate an academic-grade report containing all inputs, outputs, and the specific mathematical formula utilized.
* Click **"Excel'e Aktar"** to push the data into a structured `.xlsx` file for long-term farm management tracking.
* Click **"Geçmiş" (History)** to view your previous calculations in a robust `Treeview` data grid.

---

## 6. Installation & Deployment

### Prerequisites
* **Python 3.8+**
* Ensure `DejaVuSans.ttf` is present in the root directory for UTF-8 PDF rendering.

### Quick Start
```bash
# 1. Clone the repository
git clone [https://github.com/yourusername/seed-calibration-tool.git](https://github.com/yourusername/seed-calibration-tool.git)
cd seed-calibration-tool

# 2. Set up a virtual environment (Recommended)
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Execute the application
python main.py

## 7. Reporting & Data Export

The application utilizes a custom `PDFGenerator` class to bypass standard Tkinter limitations, ensuring high-fidelity outputs for academic and field use.
* **PDF Reports:** Powered by `ReportLab`, rendering vectorized tables, custom headers, and italicized mathematical formulas. It includes UTF-8 encoding support via `DejaVuSans` to ensure complete accuracy for localized characters.
* **Excel Reports:** Built with `openpyxl`, generating stylized spreadsheets with distinct color-coded headers (`#2E86AB` for parameters, `#C73E1D` for results). It automatically includes metadata (Timestamp, Author) and executes automated column width adjustments for immediate print-readiness.

---

## 8. Academic Citation & License

**Author:** Hasan Dural  
**Academic Supervisor:** Prof. Dr. Davut Karayel  
**Institution:** Akdeniz University, Faculty of Agriculture, Department of Agricultural Machinery and Technologies Engineering.

This software is released under the **MIT License**. Permission is granted to use, modify, and distribute this software. For academic and professional use, please provide appropriate attribution to the original author and institution.
