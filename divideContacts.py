import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side

def dividir_contatos():
    arquivo_entrada, pasta_saida, nome_base, tamanho_lote = entrada_var.get(), saida_var.get(), nome_var.get().strip(), tamanho_lote_var.get().strip()

    if not all([arquivo_entrada, pasta_saida, nome_base, tamanho_lote.isdigit()]):
        return messagebox.showerror("Erro", "Todos os campos devem ser preenchidos corretamente!")

    pasta_saida = os.path.join(pasta_saida, nome_base)
    if os.path.exists(pasta_saida):
        return messagebox.showerror("Erro", "A pasta já existe. Escolha outro nome.")

    os.makedirs(pasta_saida)
    df = pd.read_excel(arquivo_entrada, usecols=["ID", "Contato", "Telefone"], engine="openpyxl").iloc[::-1].reset_index(drop=True)
    total_contatos, tamanho_lote = len(df), int(tamanho_lote)
    
    for i, inicio in enumerate(range(0, total_contatos, tamanho_lote), 1):
        nome_arquivo = os.path.join(pasta_saida, f"{nome_base} - {i}.xlsx")
        df.iloc[inicio:inicio + tamanho_lote].to_excel(nome_arquivo, index=False, engine="openpyxl")

        wb = load_workbook(nome_arquivo)
        for cell in wb.active[1]: cell.font, cell.border = Font(bold=False), Border()
        wb.save(nome_arquivo)

        progress_var.set(int((i / ((total_contatos // tamanho_lote) + 1)) * 100))
        root.update_idletasks()

    messagebox.showinfo("Concluído", f"Processo concluído! {i} arquivos gerados.")

def selecionar_arquivo(): entrada_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))
def selecionar_pasta(): saida_var.set(filedialog.askdirectory())

root = tk.Tk()
root.title("Divisor de Contatos Excel")
root.geometry("500x350")
frame = tk.Frame(root, padx=20, pady=10, bg="#f0f0f0")
frame.pack(expand=True)

for texto, var, row, func in [("Arquivo de entrada:", entrada_var := tk.StringVar(), 0, selecionar_arquivo),
                              ("Pasta de saída:", saida_var := tk.StringVar(), 2, selecionar_pasta),
                              ("Nome base dos arquivos:", nome_var := tk.StringVar(), 4, None),
                              ("Quantidade de contatos por arquivo:", tamanho_lote_var := tk.StringVar(), 6, None)]:
    tk.Label(frame, text=texto, bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(5, 0))
    e = tk.Entry(frame, textvariable=var, width=50)
    e.grid(row=row + 1, column=0, pady=2)
    if func: tk.Button(frame, text="Selecionar", command=func).grid(row=row + 1, column=1, padx=5)
    e.bind("<Return>", lambda _: dividir_contatos())

progress_var = tk.IntVar()
ttk.Progressbar(frame, variable=progress_var, maximum=100, length=400, mode="determinate").grid(row=8, column=0, pady=10, columnspan=2)
tk.Button(frame, text="Iniciar", command=dividir_contatos, bg="#4CAF50", fg="white", padx=10, pady=5).grid(row=9, column=0, pady=10, columnspan=2)

root.mainloop()
