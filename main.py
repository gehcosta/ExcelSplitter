import os
import tkinter as tk
from tkinter import messagebox

def abrir_gerenciador_contatos():
    os.system("python filterContact.py")

def abrir_divisor_contatos():
    os.system("python divideContacts.py")

root = tk.Tk()
root.title("Menu Principal")
root.geometry("300x200")

tk.Label(root, text="Selecione o programa:", font=("Arial", 12, "bold")).pack(pady=10)

btn_contatos = tk.Button(root, text="Gerenciador de Contatos", command=abrir_gerenciador_contatos)
btn_contatos.pack(pady=5)

btn_divisor = tk.Button(root, text="Divisor de Contatos", command=abrir_divisor_contatos)
btn_divisor.pack(pady=5)

btn_sair = tk.Button(root, text="Sair", command=root.quit)
btn_sair.pack(pady=10)

root.mainloop()
