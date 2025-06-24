# Excel To CSV

This is a simple WPF application that converts Excel files (XLS/XLSX) to CSV format.  
It was created as a personal project for learning and experimentation.

Most part of code was written by Chat GPT. <sub>_(Even this README file!)_</sub>

---

## 🔧 Features

- Drag and drop Excel files into the app
- Save path is configurable and persistent
- Supports multiple files at once
- Handles commas, quotes, and newlines in cell contents correctly for CSV

---

## 📦 Dependencies

- [.NET 8.0 (WPF)](https://dotnet.microsoft.com)
- [NPOI (2.7.3 or higher)](https://github.com/nissl-lab/npoi)  
  Used to read XLS and XLSX files  
  License: Apache License 2.0

> All packages are managed via NuGet and declared in `ExcelToCsvApp.csproj`.

---

## ⚖ License

This project is licensed under the **MIT License**.  
However, it depends on third-party libraries with their own licenses:

- ExcelToCsvApp: MIT License (c) 2025 YourNameHere  
- NPOI: Apache License 2.0

You are free to use, modify, and distribute the software, respecting the licenses above.

---

## ⚠ Notes

- This tool only extracts **textual data** from Excel files.
  - Formatting, images, charts, and macros are not preserved.
- Commas, quotes, and line breaks in cell content are properly escaped in CSV format.
- The output CSV files are saved in the folder you specify.
- If an Excel file is malformed or contains unreadable data, it may be skipped or partially converted.
- Save path is stored locally under `%AppData%\ExcelToCsvApp\config.txt`.

---

## 📁 Save Path Location

The chosen CSV output folder path is saved at: `%AppData%\ExcelToCsvApp\config.txt`
This path is persistent between sessions.

---

## 📝 Purpose

This project is a **non-commercial, educational tool** made to practice:

- WPF UI development
- Drag-and-drop file handling
- Working with open-source libraries
- File input/output and CSV formatting

Feel free to fork or reuse it for your own purposes.
