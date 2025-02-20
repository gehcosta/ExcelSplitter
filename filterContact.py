import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class GerenciadorContatos:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador de Contatos Duplicados")
        self.df, self.duplicados, self.index_atual, self.selecionados, self.caminho_original = None, [], 0, {}, ""

        tk.Button(root, text="Selecionar Arquivo Excel", command=self.carregar_arquivo).pack(pady=10)
        self.tree = ttk.Treeview(root, columns=("A", "B", "C"), show="headings")
        for c, t in zip("ABC", ["Id", "Contato", "Telefone"]):
            self.tree.heading(c, text=t)
        self.tree.pack(pady=10)
        self.tree.bind("<ButtonRelease-1>", self.selecionar_linha)

        frame_botoes = tk.Frame(root)
        frame_botoes.pack(pady=10, fill=tk.X)
        tk.Button(frame_botoes, text="Voltar", command=self.voltar).pack(side=tk.LEFT, padx=20)
        tk.Button(frame_botoes, text="Próximo", command=self.proximo).pack(side=tk.RIGHT, padx=20)
        root.bind("<Return>", lambda _: self.proximo())

    def carregar_arquivo(self):
        self.caminho_original = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not self.caminho_original:
            return

        self.df = pd.read_excel(self.caminho_original, dtype=str, header=None, usecols="A:C").fillna("")
        self.df.columns = ["A", "B", "C"]
        self.duplicados = [g for _, g in self.df[self.df.duplicated(subset=['C'], keep=False)].groupby("C")]

        if not self.duplicados:
            return messagebox.showinfo("Info", "Nenhum número de telefone duplicado encontrado.")

        self.index_atual = 0
        self.exibir_grupo()

    def exibir_grupo(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        if self.index_atual < len(self.duplicados):
            for _, row in self.duplicados[self.index_atual].iterrows():
                self.tree.insert("", tk.END, values=(row["A"], row["B"], row["C"]))

    def selecionar_linha(self, _):
        if (sel := self.tree.selection()):
            self.selecionados[self.duplicados[self.index_atual]["C"].iloc[0]] = self.tree.item(sel[0], "values")

    def proximo(self):
        if self.index_atual < len(self.duplicados) - 1:
            self.index_atual += 1
            self.exibir_grupo()
        else:
            self.salvar_resultado()

    def voltar(self):
        if self.index_atual > 0:
            self.index_atual -= 1
            self.exibir_grupo()

    def salvar_resultado(self):
        if not self.selecionados:
            return messagebox.showwarning("Aviso", "Nenhum contato foi selecionado para manter.")

        df_filtrado = self.df[~self.df['C'].isin(self.selecionados.keys())]
        df_filtrado = pd.concat([df_filtrado, pd.DataFrame(self.selecionados.values(), columns=["A", "B", "C"])], ignore_index=True)
        df_filtrado = df_filtrado.sort_values(by=["A"], ascending=True)
        
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="contatos_filtrado.xlsx")
        
        if not caminho_saida:
            caminho_saida = os.path.join(os.path.dirname(self.caminho_original), "contatos_filtrado.xlsx")
        
        df_filtrado.to_excel(caminho_saida, index=False, header=False)
        messagebox.showinfo("Finalizado", f"Arquivo salvo: {caminho_saida}")
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    GerenciadorContatos(root)
    root.mainloop()
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class GerenciadorContatos:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador de Contatos Duplicados")
        self.df, self.duplicados, self.index_atual, self.selecionados, self.caminho_original = None, [], 0, {}, ""

        tk.Button(root, text="Selecionar Arquivo Excel", command=self.carregar_arquivo).pack(pady=10)
        self.tree = ttk.Treeview(root, columns=("A", "B", "C"), show="headings")
        for c, t in zip("ABC", ["Id", "Contato", "Telefone"]):
            self.tree.heading(c, text=t)
        self.tree.pack(pady=10)
        self.tree.bind("<ButtonRelease-1>", self.selecionar_linha)

        frame_botoes = tk.Frame(root)
        frame_botoes.pack(pady=10, fill=tk.X)
        tk.Button(frame_botoes, text="Voltar", command=self.voltar).pack(side=tk.LEFT, padx=20)
        tk.Button(frame_botoes, text="Próximo", command=self.proximo).pack(side=tk.RIGHT, padx=20)
        root.bind("<Return>", lambda _: self.proximo())

    def carregar_arquivo(self):
        self.caminho_original = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not self.caminho_original:
            return

        self.df = pd.read_excel(self.caminho_original, dtype=str, header=None, usecols="A:C").fillna("")
        self.df.columns = ["A", "B", "C"]
        self.duplicados = [g for _, g in self.df[self.df.duplicated(subset=['C'], keep=False)].groupby("C")]

        if not self.duplicados:
            return messagebox.showinfo("Info", "Nenhum número de telefone duplicado encontrado.")

        self.index_atual = 0
        self.exibir_grupo()

    def exibir_grupo(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        if self.index_atual < len(self.duplicados):
            for _, row in self.duplicados[self.index_atual].iterrows():
                self.tree.insert("", tk.END, values=(row["A"], row["B"], row["C"]))

    def selecionar_linha(self, _):
        if (sel := self.tree.selection()):
            self.selecionados[self.duplicados[self.index_atual]["C"].iloc[0]] = self.tree.item(sel[0], "values")

    def proximo(self):
        if self.index_atual < len(self.duplicados) - 1:
            self.index_atual += 1
            self.exibir_grupo()
        else:
            self.salvar_resultado()

    def voltar(self):
        if self.index_atual > 0:
            self.index_atual -= 1
            self.exibir_grupo()

    def salvar_resultado(self):
        if not self.selecionados:
            return messagebox.showwarning("Aviso", "Nenhum contato foi selecionado para manter.")

        df_filtrado = self.df[~self.df['C'].isin(self.selecionados.keys())]
        df_filtrado = pd.concat([df_filtrado, pd.DataFrame(self.selecionados.values(), columns=["A", "B", "C"])], ignore_index=True)
        df_filtrado = df_filtrado.sort_values(by=["A"], ascending=True)
        
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="contatos_filtrado.xlsx")
        
        if not caminho_saida:
            caminho_saida = os.path.join(os.path.dirname(self.caminho_original), "contatos_filtrado.xlsx")
        
        df_filtrado.to_excel(caminho_saida, index=False, header=False)
        messagebox.showinfo("Finalizado", f"Arquivo salvo: {caminho_saida}")
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    GerenciadorContatos(root)
    root.mainloop()