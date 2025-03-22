from win32com.client import Dispatch
from tkinter import filedialog
from os.path import dirname, abspath, join, basename
from os import getcwd
from psutil import process_iter

def process_files():
    try:
        o = Dispatch("Excel.Application")
        o.Visible = False

        # Permite seleção de múltiplos arquivos
        file_paths = filedialog.askopenfilenames(title="Selecione os arquivos", filetypes=[("Excel files", "*.xlsx")])

        if not file_paths:
            print("Nenhum arquivo selecionado.")
            return

        sheet_index = [1, 2]

        for main_path in file_paths:
            try:
                wb = o.Workbooks.Open(main_path)
                normalized_path = abspath(dirname(main_path))

                for i in sheet_index:
                    wb.WorkSheets(i).Select()
                    if i == 1:
                        cert_path = join(normalized_path, f"{basename(main_path).replace('.xlsx', '.pdf')}")
                        wb.ActiveSheet.ExportAsFixedFormat(0, cert_path)
                    if i == 2:
                        rmd_path = join(normalized_path, f"RMD_{basename(main_path).replace('.xlsx', '.pdf')}")
                        wb.ActiveSheet.ExportAsFixedFormat(0, rmd_path)

                    print(f"Arquivo {i} ({basename(main_path)}) exportado para {normalized_path}")

                wb.Close(False)
            except Exception as e:
                print(f"Erro ao processar {basename(main_path)}: {e}")

        o.Quit()

    except Exception as e:
        print(f"Erro geral: {e}")

    finally:
        if 'o' in locals():
            del o

        for proc in process_iter():
            if proc.name().lower() == "excel.exe":
                proc.kill()

process_files()
