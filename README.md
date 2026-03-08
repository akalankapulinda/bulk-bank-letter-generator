# bulk-bank-letter-generator
A desktop application built with PyQt5 and Pandas that automates bulk Word document generation from Excel data, featuring a custom Cyberpunk UI.


# Bulk Bank Confirmation Engine 🏦⚡

A high-performance desktop application built with Python and PyQt5 designed to automate the generation of bank confirmation letters. It reads customer data from an Excel file, merges it with a Microsoft Word template, and generates individual customized documents in bulk. 

It features a custom, responsive Cyberpunk-themed graphical user interface (GUI) and utilizes multithreading to ensure the app remains responsive during heavy batch processing.

## ✨ Features
* **Drag-and-Drop Interface:** Easily load customer data by dropping an `.xlsx` or `.xls` file directly into the app.
* **Cyberpunk UI:** A visually striking, dark-mode interface with custom gradients, hover effects, and real-time progress tracking.
* **Bulk Document Generation:** Uses `docxtpl` to rapidly generate personalized `.docx` files based on a master template.
* **Direct Printing:** Built-in OS-level print routing to automatically send generated files to your default printer.
* **Multithreaded Processing:** Background workers prevent the UI from freezing while generating hundreds of documents.
* **Data Validation:** Automatically parses and formats date columns and currency values before document generation. 
## 🛠️ Tech Stack
* **Python 3.x**
* **PyQt5:** For the graphical user interface.
* **Pandas:** For robust Excel data extraction and manipulation.
* **DocxTemplate:** For populating Microsoft Word template variables.

<img width="1920" height="1080" alt="Screenshot (128)" src="https://github.com/user-attachments/assets/9b4639b2-0cc3-439a-968e-462bbf72b887" />


## 📋 Prerequisites
Before running this project, ensure you have Python installed along with the required libraries. 

Create a `requirements.txt` file or install them directly:
```bash
pip install PyQt5 pandas docxtpl openpyxl


