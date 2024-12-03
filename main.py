import os
import comtypes.client
from PyPDF2 import PdfReader
from openpyxl import load_workbook
import pandas as pd


def try_open_docx(file_path, word_app):
    try:
        doc = word_app.Documents.Open(file_path)
        doc.Close()
        return True
    except Exception as e:
        log_error("word_hata.txt", file_path)
        return False


def try_open_doc(file_path, word_app):
    try:
        doc = word_app.Documents.Open(file_path)
        doc.Close()
        return True
    except Exception as e:
        log_error("word_hata.txt", file_path)
        return False


def try_open_pdf(file_path):
    try:
        reader = PdfReader(file_path)
        if len(reader.pages) > 0:
            return True
        else:
            raise ValueError("PDF has no pages")
    except Exception as e:
        log_error("pdf_hata.txt", file_path)
        return False


def try_open_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        if not df.empty:
            return True
        else:
            raise ValueError("Excel file is empty")
    except Exception as e:
        log_error("excel_hata.txt", file_path)
        return False


def log_error(file_name, file_path):
    """Hatalı dosyayı dosya ismiyle kaydeder."""
    with open(file_name, "a", encoding="utf-8") as error_file:
        error_file.write(f"{os.path.basename(file_path)}\n")


def scan_directory_for_files(directory):
    SUPPORTED_EXTENSIONS = {
        "word": [".docx", ".doc"],
        "pdf": [".pdf"],
        "excel": [".xlsx", ".xls"]
    }
    
    file_statistics = {key: {"total": 0, "successful": 0, "failed": 0} for key in SUPPORTED_EXTENSIONS}
    total_files = 0
    successful_files_count = 0
    failed_files_count = 0

    try:
        word_app = comtypes.client.CreateObject('Word.Application')
        word_app.Visible = False

        for file_name in os.listdir(directory):
            file_path = os.path.join(directory, file_name)
            success = False
            total_files += 1
            
            if any(file_name.endswith(ext) for ext in SUPPORTED_EXTENSIONS["word"]):
                file_statistics["word"]["total"] += 1
                success = try_open_docx(file_path, word_app) if file_name.endswith(".docx") else try_open_doc(file_path, word_app)
            elif file_name.endswith(".pdf"):
                file_statistics["pdf"]["total"] += 1
                success = try_open_pdf(file_path)
            elif any(file_name.endswith(ext) for ext in SUPPORTED_EXTENSIONS["excel"]):
                file_statistics["excel"]["total"] += 1
                success = try_open_excel(file_path)
            
            if success:
                successful_files_count += 1
                if file_name.endswith(tuple(SUPPORTED_EXTENSIONS["word"])):
                    file_statistics["word"]["successful"] += 1
                elif file_name.endswith(".pdf"):
                    file_statistics["pdf"]["successful"] += 1
                elif file_name.endswith(tuple(SUPPORTED_EXTENSIONS["excel"])):
                    file_statistics["excel"]["successful"] += 1
            else:
                failed_files_count += 1
                if file_name.endswith(tuple(SUPPORTED_EXTENSIONS["word"])):
                    file_statistics["word"]["failed"] += 1
                elif file_name.endswith(".pdf"):
                    file_statistics["pdf"]["failed"] += 1
                elif file_name.endswith(tuple(SUPPORTED_EXTENSIONS["excel"])):
                    file_statistics["excel"]["failed"] += 1

        word_app.Quit()
    except Exception as e:
        print(f"Error initializing Word application: {e}")
    
    return file_statistics, total_files, successful_files_count, failed_files_count


# Raporlama
current_directory = os.getcwd()
file_statistics, total_files, successful_files_count, failed_files_count = scan_directory_for_files(current_directory)

# Genel Rapor
print("\n*** GENEL RAPOR ***")
print(f"Toplam Dosya Sayısı: {total_files}")
print(f"Başarıyla Açılan Dosyalar: {successful_files_count}")
print(f"Açılamayan Dosyalar: {failed_files_count}\n")

print("*** Dosya Türü Bazında Detaylar ***")
for file_type, stats in file_statistics.items():
    print(f"{file_type.capitalize()} Dosyaları:")
    print(f"  Toplam: {stats['total']}")
    print(f"  Başarılı: {stats['successful']}")
    print(f"  Başarısız: {stats['failed']}\n")
