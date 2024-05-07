import pandas as pd
from openpyxl import Workbook
import tkinter as tk
from tkinter import messagebox

atividades = []
def cadastrar_atividade():
    atividade = entry_atividade.get()
    data = entry_data.get()
    if atividade and data:
        id_atividade = len(atividades) + 1
        atividades.append({"ID": id_atividade, "Atividade": atividade, "Data": data})
        messagebox.showinfo("Sucesso", "Atividade cadastrada com sucesso!")
        entry_atividade.delete(0, tk.END)
        entry_data.delete(0, tk.END)
    else:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos.")


def exibir_atividades():
    if not atividades:
        messagebox.showinfo("Informação", "Nenhuma atividade cadastrada ainda.")
    else:
        atividades_str = ""
        for atividade in atividades:
            atividades_str += f"ID: {atividade['ID']}, Atividade: {atividade['Atividade']}, Data: {atividade['Data']}\n"
        messagebox.showinfo("Atividades Cadastradas", atividades_str)


def exportar_para_csv():
    if not atividades:
        messagebox.showinfo("Informação", "Nenhuma atividade cadastrada para exportar.")
        return

    df = pd.DataFrame(atividades)
    df.to_csv("atividades.csv", index=False)
    messagebox.showinfo("Sucesso", "Atividades exportadas para 'atividades.csv' com sucesso!")


def exportar_para_excel():
    if not atividades:
        messagebox.showinfo("Informação", "Nenhuma atividade cadastrada para exportar.")
        return

    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Atividade", "Data"])
    for atividade in atividades:
        ws.append([atividade["ID"], atividade["Atividade"], atividade["Data"]])

    wb.save("atividades.xlsx")
    messagebox.showinfo("Sucesso", "Atividades exportadas para 'atividades.xlsx' com sucesso!")


# Configuração da interface gráfica
root = tk.Tk()
root.title("Cadastro de Atividades")

label_atividade = tk.Label(root, text="Atividade:")
label_atividade.pack()
entry_atividade = tk.Entry(root, width=50)
entry_atividade.pack()

label_data = tk.Label(root, text="Data (dd/mm/aaaa):")
label_data.pack()
entry_data = tk.Entry(root, width=20)
entry_data.pack()

btn_cadastrar = tk.Button(root, text="Cadastrar", command=cadastrar_atividade)
btn_cadastrar.pack()

btn_exibir = tk.Button(root, text="Exibir Atividades", command=exibir_atividades)
btn_exibir.pack()

btn_exportar_csv = tk.Button(root, text="Exportar para CSV", command=exportar_para_csv)
btn_exportar_csv.pack()

btn_exportar_excel = tk.Button(root, text="Exportar para Excel", command=exportar_para_excel)
btn_exportar_excel.pack()

root.mainloop()
