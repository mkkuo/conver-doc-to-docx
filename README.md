# 📝 DOC to DOCX Converter (Windows Only)

本工具是一個自動化 Python 腳本，可用來將目錄下所有 `.doc` 檔案轉換成 `.docx` 格式，支援遞迴處理資料夾結構，適用於 Windows 系統。

---

## 🔧 功能特色

- 自動掃描目錄與子目錄內所有 `.doc` 檔案
- 使用 Microsoft Word COM 元件進行轉換
- 每個 `.doc` 檔案將生成對應 `.docx` 檔案
- 可選擇是否刪除原始 `.doc` 檔

---

## 📂 專案結構

```
.
├── convert doc to docx.py    # 主程式
├── your_folder/              # 放置 Word 檔案的目錄
└── README.md
```

---

## ⚙️ 使用方式

1. 安裝 Python 套件（需安裝 `pywin32`）：

```bash
pip install pywin32
```

2. 修改腳本中的目錄路徑：

```python
file_path = r"D:\your\path\to\docs"
```

3. 執行程式：

```bash
python "convert doc to docx.py"
```

---

## 📌 注意事項

- 僅適用於 Windows 系統（依賴 Word 的 COM 自動化）
- 請確認電腦已安裝 Microsoft Word
- 若 Word 視窗未關閉，程式將會自動結束 Word 應用

---

## 📄 授權

本專案採用 MIT License 授權。
