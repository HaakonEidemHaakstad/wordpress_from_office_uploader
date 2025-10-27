# 🧭 WordPress Uploader v1.0



**WordPress Uploader** is a lightweight, open-source desktop app built with **Python** and **PySide6** that lets you update WordPress pages directly from **Word**, **Excel**, or **HTML** files — without manually cleaning or pasting code.  



It converts `.docx` and `.xlsx` files into clean, self-contained HTML that preserves the original formatting, then uploads them to your WordPress site via the **REST API**.  



Perfect for teams or individuals who regularly update content such as tables, newsletters, or schedules from Office documents.



---



## ✨ Features



✅ **Drag-and-Drop Uploads**  

Drop a `.docx`, `.xlsx`, or `.html` file straight into the window — no file dialog required.



✅ **Accurate Office Conversion**  

Uses Microsoft Word and Excel (via COM automation) to export documents as HTML while keeping fonts, colors, borders, alignment, and spacing intact.



✅ **Smart HTML Cleaning**  

A built-in BeautifulSoup/LXML cleaner:

- Removes unnecessary Microsoft Office markup.  

- Inlines all CSS and images as `data:` URIs.  

- Produces WordPress-compatible HTML that looks exactly like your original document.



✅ **Safe Uploading with Undo / Regret**  

Before each upload, the program automatically downloads and backs up the current page’s HTML.  

If you upload to the wrong page or don’t like the result, click **Undo / Regret** to instantly restore the previous version.



✅ **Credential Auto-Fill**  

Optionally save credentials to `defaults.txt` (password encrypted using Windows DPAPI when available).



✅ **Progress Feedback**  

A progress bar and status log show every step of the process — conversion, cleaning, upload.



✅ **Cross-Version Compatible**  

Runs on **Python 3.10–3.14** with **Windows + Microsoft Office** installed.



---



## 🧩 Requirements



- **Windows 10/11**

- **Microsoft Office** (Word + Excel)

- **Python 3.10–3.14**



Install dependencies:



```bash

pip install -r requirements.txt

```



**requirements.txt**

```

PySide6

requests

beautifulsoup4

lxml

pywin32

chardet

```



---



 ⚙️ Setup & Usage



1. **Clone or download** this repository:

```bash

git clone https://github.com/<your-username>/wordpress-uploader.git

cd wordpress-uploader

```



2. **Create and activate a virtual environment (optional)**

```bash

python -m venv venv

venv\\Scripts\\activate

```



3. **Install dependencies**

```bash

pip install -r requirements.txt

```



4. **Run the application**

```bash

python main.py

```



5. **(Optional)** Build a standalone executable:

```bash

pyinstaller --onefile --windowed main.py

 ```

Output will appear in `dist/main.exe`.

---



## 🔑 Connecting to WordPress



1. Log into your WordPress admin panel.  

2. Go to `Users → Profile → Application Passwords`.  

3. Create a new application password (give it a name like **Uploader**).  

4. In the app, enter:

- **Site URL** — e.g. `https://example.com`

- **Username** — your WordPress username

- **Application Password**

5. Click **Log In** → the dropdown will populate with all editable pages.

6. Drag your file into the window, check “I have selected the correct page,” and click **Update**.



---



## 🩹 Undo / Regret



Each upload automatically saves the previous page content to a local file:



```

temp\_backup.html

```



If you need to revert, click **Undo / Regret**, and the app will re-upload the previous version to WordPress.



---



 🧰 Technical Overview
```

| Component            | Purpose                                        |

|----------------------|------------------------------------------------|

| PySide6              | GUI framework (Qt for Python)                  |

| pywin32              | Accesses Word/Excel via COM for conversion     |

| BeautifulSoup + lxml | Cleans and normalizes exported HTML            |

| requests             | Uploads content through the WordPress REST API |

| chardet              | Auto-detects file encodings                    |

| DPAPI (win32crypt)   | Optional encryption for stored credentials     |
```


**Workflow:**

```

Drag & Drop → Office COM export → HTML Cleaning → Inline assets → Upload → Undo backup

```



---



## 📂 Project Structure



```

wordpress-uploader/

│

├── main.py                # Main PySide6 application

├── requirements.txt       # Dependencies

├── README.md              # This file

└── LICENSE                # MIT License

```



---



## 🏷 Version History


```
| Version | Date       | Highlights                                                                 |

|---------|------------|----------------------------------------------------------------------------|

| 1.0.0   | 2025-10-27 | Initial public release — full Word/Excel support, Undo/Regret, PySide6 GUI |
```

---



## 💡 Example Use Case



Non-tech savvy individuals that know how to create Microsoft Word or Excel can easily produce websites hosted using WordPress by simply producing the web pages using Word or Excel and uploading them to WordPress, without having to navigate the WordPress user interface.

With this tool, you simply drag in the Word or Excel file, pick the page to update, and click **Update** — done.



---



## 🧑‍💻 Author



**Haakon Eidem Haakstad**  

[GitHub](https://github.com/<your-username>)



---



## 🪪 License



This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.



```text

MIT License



Copyright (c) 2025 Haakon Eidem Haakstad



Permission is hereby granted, free of charge, to any person obtaining a copy

of this software and associated documentation files (the "Software"), to deal

in the Software without restriction...

```



---



## 🙏 Acknowledgments



- Built with **Qt for Python (PySide6)**.  

- Uses **BeautifulSoup**, **lxml**, and **pywin32**.  

- Inspired by real needs at Fontenehuset Asker for efficient content publishing.  

- Open-sourced to help others simplify the same workflow.



