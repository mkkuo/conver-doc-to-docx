import os
import win32com.client as win32
import traceback

# 設定目標資料夾
file_path = r"{your_path}"

# 啟動 Word 應用程式 (只啟動一次)
word_app = win32.DispatchEx("Word.Application")
word_app.Visible = False  # 讓 Word 在背景執行，不顯示視窗

try:
    for root, dirs, files in os.walk(file_path):
        for doc_file in files:
            if doc_file.endswith(".doc"):
                doc_file_path = os.path.join(root, doc_file)
                docx_file_name = os.path.splitext(doc_file)[0] + ".docx"
                docx_file_path = os.path.join(root, docx_file_name)

                print(f'正在處理：{doc_file_path}')
                print('----------------------')

                try:
                    doc = word_app.Documents.Open(doc_file_path)
                    doc.SaveAs(docx_file_path, FileFormat=16)
                    doc.Close()
                    print(f"已成功轉換：{doc_file_path} -> {docx_file_path}")

                    # 移除doc檔案（非必要）
                    # os.remove(doc_file_path)

                except Exception as inner_e:
                    print(f"轉換 {doc_file_path} 時發生內部錯誤：{inner_e}")
                    traceback.print_exc()

except Exception as outer_e:
    print(f"發生外部錯誤：{outer_e}")
    traceback.print_exc()

finally:
    word_app.Quit()  # 確保 Word 在任何情況下都會被關閉
    print('所有doc檔轉換docx檔：已完成')