# 📊 CSV / Excel Viewer

A lightweight desktop application for viewing, searching, and navigating large CSV and Excel files efficiently.

Built with **Python**, **Tkinter**, and **Pandas**.

---

## 🚀 Features

- 📂 Open **CSV and Excel (.xlsx)** files  
- ⚡ Efficient handling of **large files** using chunk loading  
- 🔎 Full-text **search across entire file**  
- 📄 Pagination (Next / Previous chunks)  
- 🎨 Alternating row colors for better readability  
- 🌟 Highlight search results in the table  
- 📋 Copy selected rows (Right-click or Ctrl + C)  
- 📑 Excel **multi-sheet support**  
- 🧠 Works without loading the whole file into memory  

---

## ⚙️ Requirements

- Python 3.8+
- Required libraries:

```bash
pip install pandas openpyxl
```

---

## ▶️ How to Run

```bash
python your_file_name.py
```

---

## 📌 How It Works

### CSV Files
- Loaded in chunks (default: 5000 rows)
- Only one chunk is displayed at a time
- Improves performance on large files

### Excel Files
- Loads all sheets
- Displays one sheet at a time

### Search
- Scans the **entire file**
- Results are shown in a separate window
- Matching rows are highlighted in the main table

---

## ⌨️ Controls

| Action | Shortcut |
|------|--------|
| Copy selected rows | Ctrl + C |
| Context menu | Right click |
| Navigate chunks | Buttons |

---

## 📁 Project Structure

```
project/
│
├── main.py        # Main application file
└── README.md      # Documentation
```

---

## 🧩 Key Components

- Tkinter – GUI  
- ttk.Treeview – Table display  
- Pandas – Data processing  
- Chunk-based reading for performance  

---

## ⚡ Performance Notes

- Handles large CSV files without freezing  
- Does not load entire dataset into memory  
- Suitable for files **1GB+**  

---

## 📄 License

Free to use for personal and educational purposes.
