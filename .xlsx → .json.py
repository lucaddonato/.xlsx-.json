'''
.xlsx → .json
'''





# Bibliotecas
import os
import tkinter as tk
from ctypes import windll
from tkinter import filedialog, font, messagebox, ttk

import pandas as pd





# Local
caminho = os.path.join(os.path.expanduser("~"), "Desktop", "JSON")
os.makedirs(caminho, exist_ok=True)





# Dicionário
xlsx = []
planilhas_por_arquivo = {}
comboboxes = {}





# Selcionar
def selecionar():
    arquivos = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xls")])


    for arquivo in arquivos:
        label = tk.Label(frame_planilhas, text=arquivo, font=root)
        label.pack(anchor="w")
    root.update_idletasks()
    root.minsize(root.winfo_width(), root.winfo_height())


    if arquivos:
        xlsx.clear()
        planilhas_por_arquivo.clear()
        for widget in frame_planilhas.winfo_children():
            widget.destroy()


        for arquivo in arquivos:
            xlsx.append(arquivo)
            try:
                planilhas = pd.ExcelFile(arquivo).sheet_names
                planilhas_por_arquivo[arquivo] = tk.StringVar()
                ttk.Label(frame_planilhas, text=os.path.basename(arquivo)).pack(anchor="w")
                cb = ttk.Combobox(frame_planilhas, textvariable=planilhas_por_arquivo[arquivo], values=planilhas, state="readonly")
                cb.pack(fill="x", pady=2)
                cb.current(0)
                comboboxes[arquivo] = cb
            except Exception as erro:
                messagebox.showerror("Erro", f"{os.path.basename(arquivo)}:\n{erro}")





# Converter
def converter():
    if not xlsx:
        return messagebox.showwarning("Aviso", "Selecione arquivos!")


    try:
        for arquivo in xlsx:
            nome_planilha = planilhas_por_arquivo[arquivo].get()
            if not nome_planilha:
                continue

            df = pd.read_excel(arquivo, sheet_name=nome_planilha, header=7)
            df.to_json(
                os.path.join(caminho, f"{os.path.splitext(os.path.basename(arquivo))[0]}_{nome_planilha}.json"),
                orient="records",
                force_ascii=False,
                indent=4
            )
        messagebox.showinfo("Sucesso", f"\n{caminho}")
    except Exception as erro:
        messagebox.showerror("Erro", f"\n{erro}")





# Janela
root = tk.Tk()
root.title("Converter")
windll.shcore.SetProcessDpiAwareness(2)  # Alta resolução

root.option_add("*Font", ("Calibri", 14))

tk.Label(root, text=".xlsx   →   .json", font=("Calibri", 22, "bold"))\
  .pack(pady=(10, 50))
tk.Button(root, text="Selecionar arquivos", command=selecionar)\
  .pack(pady=10)

frame_planilhas = tk.Frame(root)
frame_planilhas.pack(fill="both", expand=True, padx=10, pady=10)

tk.Button(root, text="Converter", command=converter)\
  .pack(pady=10)

root.update_idletasks()
root.minsize(root.winfo_width(), root.winfo_height())

root.mainloop()