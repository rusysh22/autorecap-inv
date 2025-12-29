# Nefi AutoRecap File Invoice - JNT Express 🚛📊

**Consolidate Vendor J&T Express Invoices into System Format Instantly.**

A modern, web-based tool designed to automate the consolidation of multiple Vendor Invoice Excel files into a single, standardized System Import Format. Built with **Python (Flask)** and **Pandas**, featuring a beautiful **Glassmorphism UI**.

![App Screenshot](https://via.placeholder.com/800x400?text=Application+UI+Preview)

## ✨ Key Features

*   **Drag & Drop Interface**: Easily upload multiple Excel files (`.xlsx`, `.xls`) at once.
*   **Smart Data Processing**:
    *   Automatically merges data from multiple source files.
    *   Cleans and formats "Nama Tugas" (Route Names).
    *   Maps columns intelligently based on J&T Express requirements.
    *   Calculates **PPN 1.1%** and **PPh 2%** automatically.
*   **Visual Breakdown**:
    *   **Color-Coded Tracks**: Each file gets a unique color for easy visual tracking.
    *   **Payment Breakdown Table**: View row counts, tax details, and total amounts per file.
    *   **Consolidated Preview**: See the merged data before downloading.
*   **Customizable Output**:
    *   Mandatory Filename Prefix: `陆运数据核对_`
    *   Customizable Suffix (e.g., Date or Batch Code).
    *   **Strict Excel Styling**: Output file matches the exact header format (SimSun font, Red text, Grey background, Height 30).
*   **Auto-Cleanup**: Automatically deletes files older than 1 hour to save disk space.

## 🛠️ Tech Stack

*   **Backend**: Python 3.x, Flask
*   **Data Processing**: Pandas, OpenPyXL
*   **Frontend**: HTML5, Vanilla JavaScript, Tailwind CSS (via CDN)
*   **Design**: Modern Glassmorphism Aesthetic

## 🚀 Installation & Usage

1.  **Clone the Repository**
    ```bash
    git clone https://github.com/rusysh22/autorecap-inv.git
    cd autorecap-inv
    ```

2.  **Install Dependencies**
    Ensure you have Python installed, then run:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the Application**
    ```bash
    python app.py
    ```

4.  **Open in Browser**
    Visit `http://localhost:5000` in your web browser.

## 📝 Column Mapping Rule

The application maps specific columns from the Vendor Invoice to the System Format:

| System Column | Source Column | Note |
| :--- | :--- | :--- |
| **Agen Operasional** | Col B | |
| **Kode Tugas** | Col D | |
| **Nama Tugas** | Col G | Cleaned (Split by '-', first 3 chars) |
| **Plat Mobil** | Col H | |
| **Jenis Kendaraan** | Col I | |
| **Mode Operasi** | Col J | |
| **Metode Perhitungan** | Col O | |
| **Berat** | - | Default: 0 |
| **Tarif Pengiriman per kg** | - | Default: 0 |
| **Tarif Pengiriman Sistem** | Col P | |
| **PPN** | Col U | |
| **PPH** | Col V | |
| **Total pembayaran aktual** | Col W | |

## © Copyright

**Nefi AutoRecap** is created by **Nefi Yunilistya**.
Copyright © 2025. All Rights Reserved.
