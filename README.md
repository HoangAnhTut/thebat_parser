# 📧 EML to DOCX Email Parser

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)
[![BeautifulSoup4](https://img.shields.io/badge/bs4-used-success)](https://pypi.org/project/beautifulsoup4/)
[![pandas](https://img.shields.io/badge/pandas-not_used-lightgrey)](https://pandas.pydata.org/)
[![python-docx](https://img.shields.io/badge/python--docx-used-success)](https://pypi.org/project/python-docx/)

## 📌 Purpose

This script parses `.eml` email files stored in the `data/` directory and extracts key information such as:

- Date and time of the message
- Sender and recipient
- Subject and message body (converted from HTML if needed)
- Names of any attachments

It then generates a well-structured `.docx` file (`emails.docx`) with all the extracted data organized in a table format.

## ⚙️ Used Technologies

- **Python 3.8+**
- `beautifulsoup4` — for parsing and cleaning up HTML email content.
- `python-docx` — for generating Word documents with email tables.
- `email` (standard library) — for reading and decoding `.eml` files.

> 📌 `pandas` is listed in the `requirements.txt` but **not currently used** in the script.

## 🗂 Folder Structure

```
project/
│
├── data/               # Folder containing .eml files
├── main.py             # Main script to run the parser
├── requirements.txt    # Python dependencies
├── start.bat           # Windows batch file to launch the script
└── emails.docx         # Output Word document (generated)
```

## 🚀 How to Run

1. Install the dependencies:

```bash
pip install -r requirements.txt
```

2. Place your `.eml` email files into the `data/` folder.

3. Run the script:

```bash
python main.py
```

4. The resulting `emails.docx` file will be created in the project root.

## 📄 Output Format

The output Word file will contain a table with the following columns:

- Date/Time  
- Sender  
- Recipient  
- Email Subject and Body  
- Attachment Names  


## 📜 License

[MIT License](LICENSE) — free to use, adapt, and improve 🤘

---

## 🤝 Contact
[![Telegram Badge](https://img.shields.io/badge/Contact-blue?style=flat&logo=telegram&logoColor=white)](https://t.me/spystars777)
