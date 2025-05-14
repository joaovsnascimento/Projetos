import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import ttkbootstrap as ttk
from ttkbootstrap.style import Style
from ttkbootstrap.constants import *

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r'O:\Shared drives\credencial.json', scope)
client = gspread.authorize(creds)
sheet = client.open_by_key('base_sheets').sheet1

historico_operacoes = []

def ler_estoque():
    dados = sheet.get_all_records()
    estoque_text.delete(1.0, tk.END)
    for item in dados:
        estoque_text.insert(tk.END, f"SKU: {item['SKU']}    Quantidade: {item['Quantidade']}\n")

def adicionar_item():
    nome = entry_nome.get()
    quantidade = entry_quantidade.get()
    if nome and quantidade:
        try:
            quantidade = int(quantidade)
            novo_item = [nome, quantidade]
            sheet.append_row(novo_item)
            historico_operacoes.append(('adicionar_item', nome, quantidade))
            messagebox.showinfo("Sucesso", f"Item {nome} adicionado com sucesso!")
            ler_estoque()
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira uma quantidade válida.")
    else:
        messagebox.showerror("Erro", "Por favor, insira todos os campos.")

def adicionar_quantidade():
    nome = entry_nome.get()
    quantidade_adicional = entry_quantidade.get()
    try:
        quantidade_adicional = int(quantidade_adicional)
        if quantidade_adicional < 0:
            raise ValueError("A quantidade para adicionar deve ser maior que zero.")
        cell = sheet.find(nome)
        if cell:
            quantidade_atual = int(sheet.cell(cell.row, cell.col + 1).value)
            nova_quantidade = quantidade_atual + quantidade_adicional
            sheet.update_cell(cell.row, cell.col + 1, nova_quantidade)
            historico_operacoes.append(('adicionar_quantidade', nome, quantidade_adicional))
            messagebox.showinfo("Sucesso", f"{quantidade_adicional} unidades adicionadas a {nome}. Nova quantidade: {nova_quantidade}.")
            ler_estoque()
        else:
            messagebox.showerror("Erro", f"Item {nome} não encontrado.")
    except ValueError as e:
        messagebox.showerror("Erro", str(e))

def subtrair_quantidade():
    nome = entry_nome.get()
    quantidade_subtrair = entry_quantidade.get()
    try:
        quantidade_subtrair = int(quantidade_subtrair)
        if quantidade_subtrair < 0:
            raise ValueError("A quantidade a ser retirada deve ser positiva.")
        cell = sheet.find(nome)
        if cell:
            quantidade_atual = int(sheet.cell(cell.row, cell.col + 1).value)
            if quantidade_atual >= quantidade_subtrair:
                nova_quantidade = quantidade_atual - quantidade_subtrair
                sheet.update_cell(cell.row, cell.col + 1, nova_quantidade)
                historico_operacoes.append(('subtrair_quantidade', nome, quantidade_subtrair))
                messagebox.showinfo("Sucesso", f"{quantidade_subtrair} unidades subtraídas de {nome}. Nova quantidade: {nova_quantidade}.")
                ler_estoque()
            else:
                messagebox.showerror("Erro", f"A quantidade para subtrair excede o estoque atual de {nome}.")
        else:
            messagebox.showerror("Erro", f"Item {nome} não encontrado.")
    except ValueError as e:
        messagebox.showerror("Erro", str(e))

def remover_item():
    nome = entry_nome.get()
    cell = sheet.find(nome)
    if cell:
        quantidade_atual = int(sheet.cell(cell.row, cell.col + 1).value)
        sheet.delete_rows(cell.row)
        historico_operacoes.append(('remover_item', nome, quantidade_atual))
        messagebox.showinfo("Sucesso", f"Item {nome} removido com sucesso!")
        ler_estoque()
    else:
        messagebox.showerror("Erro", f"Item {nome} não encontrado.")

def procurar_item():
    nome = entry_nome.get()
    cell = sheet.find(nome)
    if cell:
        quantidade = sheet.cell(cell.row, cell.col + 1).value
        messagebox.showinfo("Item Encontrado", f"Item: {nome}\nQuantidade: {quantidade}")
    else:
        messagebox.showerror("Erro", f"Item {nome} não encontrado.")

def desfazer_ultima_operacao():
    if not historico_operacoes:
        messagebox.showinfo("Erro", "Nenhuma operação para desfazer.")
        return

    ultima_operacao = historico_operacoes.pop()
    tipo_operacao, nome, quantidade = ultima_operacao

    if tipo_operacao == 'adicionar_item':
        cell = sheet.find(nome)
        if cell:
            sheet.delete_rows(cell.row)
            messagebox.showinfo("Sucesso", f"Adição do item {nome} desfeita com sucesso!")
    elif tipo_operacao == 'adicionar_quantidade':
        cell = sheet.find(nome)
        if cell:
            quantidade_atual = int(sheet.cell(cell.row, cell.col + 1).value)
            nova_quantidade = quantidade_atual - quantidade
            sheet.update_cell(cell.row, cell.col + 1, nova_quantidade)
            messagebox.showinfo("Sucesso", f"Adição de quantidade para {nome} desfeita com sucesso!")
    elif tipo_operacao == 'subtrair_quantidade':
        cell = sheet.find(nome)
        if cell:
            quantidade_atual = int(sheet.cell(cell.row, cell.col + 1).value)
            nova_quantidade = quantidade_atual + quantidade
            sheet.update_cell(cell.row, cell.col + 1, nova_quantidade)
            messagebox.showinfo("Sucesso", f"Subtração de quantidade para {nome} desfeita com sucesso!")
    elif tipo_operacao == 'remover_item':
        sheet.append_row([nome, quantidade])
        messagebox.showinfo("Sucesso", f"Remoção do item {nome} desfeita com sucesso!")

    ler_estoque()

def salvar_estoque():
    try:
        dados = sheet.get_all_records()
        nova_aba = client.open_by_key('base_sheets').add_worksheet(title="Estoque Backup", rows="100", cols="20")
        nova_aba.append_row(["SKU", "Quantidade"])
        for item in dados:
            nova_aba.append_row([item['SKU'], item['Quantidade']])
        messagebox.showinfo("Sucesso", "Estoque salvo com sucesso na nova aba!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar estoque: {str(e)}")

app = tk.Tk()
app.title("Gerenciador de Estoque")
app.geometry("550x350")
style = Style(theme = "cyborg")

tk.Label(app, text="SKU").grid(row=0, column=0)
tk.Label(app, text="Quantidade").grid(row=1, column=0)

entry_nome = tk.Entry(app)
entry_nome.grid(row=0, column=1)
entry_quantidade = tk.Entry(app)
entry_quantidade.grid(row=1, column=1)

tk.Button(app, text="Ver Estoque", command=ler_estoque).grid(row=2, column=0, pady=4)
tk.Button(app, text="Adicionar Item", command=adicionar_item).grid(row=4, column=0, pady=4)
tk.Button(app, text="Adicionar Quantidade", command=adicionar_quantidade).grid(row=3, column=0, pady=4)
tk.Button(app, text="Subtrair Quantidade", command=subtrair_quantidade).grid(row=3, column=1, pady=4)
tk.Button(app, text="Remover Item", command=remover_item).grid(row=4, column=1, pady=4)
tk.Button(app, text="Procurar Item", command=procurar_item).grid(row=2, column=1, pady=4)
tk.Button(app, text="Desfazer", command=desfazer_ultima_operacao).grid(row=5, column=0, pady=4)
tk.Button(app, text="Salvar Estoque", command=salvar_estoque).grid(row=5, column=1, pady=4)

estoque_text = tk.Text(app, height=20, width=40)
estoque_text.grid(row=0, column=2, rowspan=6, padx=10, pady=10)

app.mainloop()
