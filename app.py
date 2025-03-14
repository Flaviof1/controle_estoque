import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from PIL import Image, ImageTk
import locale
from openpyxl import Workbook

# Configuração de localidade para formato de moeda
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Variável global para armazenar o total vendido
total_vendido = 0.0

# Função para conectar ao banco de dados SQLite
def conectar():
    try:
        conn = sqlite3.connect('banco_de_dados.db')
        return conn
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro de conexão com o banco de dados: {e}")
        return None

# Função para criar a tabela de produtos no banco de dados
def criar_tabela():
    conn = conectar()
    if conn is None:
        return
    
    try:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS produtos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                quantidade INTEGER NOT NULL,
                preco REAL NOT NULL
            )
        ''')
        conn.commit()
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao criar a tabela: {e}")
    finally:
        conn.close()

# Função para adicionar um novo produto ao banco de dados
def adicionar_produto():
    nome = entry_nome.get()
    quantidade = entry_quantidade.get()
    preco = entry_preco.get()

    if not nome or not quantidade or not preco:
        messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
        return

    if not quantidade.isdigit() or not preco.replace('.', '', 1).isdigit():
        messagebox.showerror("Erro", "Quantidade deve ser um número inteiro e preço deve ser um número válido!")
        return

    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO produtos (nome, quantidade, preco)
            VALUES (?, ?, ?)
        """, (nome, int(quantidade), float(preco)))
        conn.commit()
        messagebox.showinfo("Sucesso", "Produto adicionado com sucesso!")
        
        entry_nome.delete(0, tk.END)
        entry_quantidade.delete(0, tk.END)
        entry_preco.delete(0, tk.END)
        
        listar_produtos()
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao adicionar produto: {e}")
    finally:
        conn.close()

# Função para listar os produtos do banco de dados (apenas nome e quantidade)
def listar_produtos():
    for item in tree.get_children():
        tree.delete(item)
    
    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("SELECT nome, quantidade FROM produtos")
        
        produtos = cursor.fetchall()
        if not produtos:
            messagebox.showinfo("Informação", "Nenhum produto encontrado!")
        
        for row in produtos:
            tree.insert("", "end", values=(row[0], row[1]))
    
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao listar produtos: {e}")
    
    finally:
        conn.close()

# Função para pesquisar produtos
def pesquisar_produto():
    termo = entry_pesquisa.get().strip().lower()
    if not termo:
        messagebox.showwarning("Aviso", "Digite um termo para pesquisar!")
        return
    
    for item in tree.get_children():
        tree.delete(item)
    
    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("SELECT nome, quantidade FROM produtos WHERE LOWER(nome) LIKE ?", (f"%{termo}%",))
        
        produtos = cursor.fetchall()
        if not produtos:
            messagebox.showinfo("Informação", "Nenhum produto encontrado!")
        
        for row in produtos:
            tree.insert("", "end", values=(row[0], row[1]))
    
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao pesquisar produtos: {e}")
    
    finally:
        conn.close()

# Função para excluir um produto do banco de dados
def excluir_produto():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Erro", "Selecione um produto para excluir!")
        return
    
    produto_nome = tree.item(selected_item, "values")[0]
    
    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("DELETE FROM produtos WHERE nome = ?", (produto_nome,))
        conn.commit()
        messagebox.showinfo("Sucesso", "Produto excluído com sucesso!")
        listar_produtos()
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao excluir produto: {e}")
    finally:
        conn.close()

# Função para vender (retirar unidades) de um produto
def vender_produto():
    global total_vendido

    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Erro", "Selecione um produto para vender!")
        return
    
    produto_nome = tree.item(selected_item, "values")[0]
    quantidade_vender = entry_venda_quantidade.get()

    if not quantidade_vender.isdigit() or int(quantidade_vender) <= 0:
        messagebox.showerror("Erro", "Digite uma quantidade válida para venda!")
        return
    
    quantidade_vender = int(quantidade_vender)

    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("SELECT quantidade, preco FROM produtos WHERE nome = ?", (produto_nome,))
        resultado = cursor.fetchone()
        
        if not resultado:
            messagebox.showerror("Erro", "Produto não encontrado!")
            return
        
        quantidade_atual = resultado[0]
        preco_unitario = resultado[1]
        
        if quantidade_atual >= quantidade_vender:
            nova_quantidade = quantidade_atual - quantidade_vender
            cursor.execute("UPDATE produtos SET quantidade = ? WHERE nome = ?", (nova_quantidade, produto_nome))
            conn.commit()
            
            total_vendido += preco_unitario * quantidade_vender
            label_total_vendido.config(text=f"Total Vendido: {locale.currency(total_vendido, grouping=True)}")
            
            messagebox.showinfo("Sucesso", f"{quantidade_vender} unidades vendidas!")
            listar_produtos()
        else:
            messagebox.showwarning("Aviso", "Não há unidades suficientes para venda!")
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao vender produto: {e}")
    finally:
        conn.close()

# Função para atualizar um produto no banco de dados
def atualizar_produto():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Erro", "Selecione um produto para atualizar!")
        return
    
    produto_nome = tree.item(selected_item, "values")[0]
    nome = entry_nome.get()
    quantidade = entry_quantidade.get()
    preco = entry_preco.get()

    if not nome or not quantidade or not preco:
        messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
        return

    if not quantidade.isdigit() or not preco.replace('.', '', 1).isdigit():
        messagebox.showerror("Erro", "Quantidade deve ser um número inteiro e preço deve ser um número válido!")
        return
    
    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE produtos
            SET nome = ?, quantidade = ?, preco = ?
            WHERE nome = ?
        """, (nome, int(quantidade), float(preco), produto_nome))
        conn.commit()
        messagebox.showinfo("Sucesso", "Produto atualizado com sucesso!")
        listar_produtos()
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao atualizar produto: {e}")
    finally:
        conn.close()

# Função para exportar produtos para Excel
def exportar_excel():
    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, quantidade, preco FROM produtos")
        produtos = cursor.fetchall()
        
        if not produtos:
            messagebox.showinfo("Informação", "Nenhum produto para exportar!")
            return
        
        # Abrir diálogo para salvar o arquivo
        arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not arquivo:
            return
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Produtos"
        
        # Adicionar cabeçalhos
        ws.append(["ID", "Nome", "Quantidade", "Preço"])
        
        # Adicionar dados
        for produto in produtos:
            ws.append(produto)
        
        wb.save(arquivo)
        messagebox.showinfo("Sucesso", f"Dados exportados para {arquivo}!")
    
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao exportar dados: {e}")
    finally:
        conn.close()

# Função para abrir a janela de estoque
def abrir_janela_estoque():
    janela_estoque = tk.Toplevel(root)
    janela_estoque.title("Estoque de Produtos")
    janela_estoque.geometry("800x600")
    
    # Frame para conter a Treeview e a barra de rolagem
    frame_tree = tk.Frame(janela_estoque)
    frame_tree.pack(fill="both", expand=True, padx=20, pady=20)
    
    # Treeview para exibir o estoque
    tree_estoque = ttk.Treeview(frame_tree, columns=("ID", "Nome", "Quantidade", "Preço"), show="headings")
    tree_estoque.pack(side="left", fill="both", expand=True)
    
    # Barra de rolagem
    scrollbar = ttk.Scrollbar(frame_tree, orient="vertical", command=tree_estoque.yview)
    scrollbar.pack(side="right", fill="y")
    tree_estoque.configure(yscrollcommand=scrollbar.set)
    
    tree_estoque.heading("ID", text="ID")
    tree_estoque.heading("Nome", text="Nome")
    tree_estoque.heading("Quantidade", text="Quantidade")
    tree_estoque.heading("Preço", text="Preço")
    
    try:
        conn = conectar()
        if conn is None:
            return
        
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, quantidade, preco FROM produtos")
        produtos = cursor.fetchall()
        
        # Calcular o capital investido
        capital_investido = sum(quantidade * preco for _, _, quantidade, preco in produtos)
        
        for produto in produtos:
            tree_estoque.insert("", "end", values=produto)
        
        # Exibir o capital investido e o total vendido
        label_capital_investido = tk.Label(janela_estoque, text=f"Capital Investido: {locale.currency(capital_investido, grouping=True)}", font=("Arial", 14))
        label_capital_investido.pack(pady=10)
        
        label_total_vendido_estoque = tk.Label(janela_estoque, text=f"Total Vendido: {locale.currency(total_vendido, grouping=True)}", font=("Arial", 14))
        label_total_vendido_estoque.pack(pady=10)
    
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao carregar estoque: {e}")
    finally:
        conn.close()

# Criando a janela principal
root = tk.Tk()
root.title("Controle de Estoque")
root.geometry("1200x800")  # Aumentei o tamanho da janela
root.config(bg="#f0f0f0")

# Configurar o grid da janela principal para centralizar o conteúdo
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=0)  # Logo
root.grid_rowconfigure(1, weight=0)  # Frames de adicionar/vender
root.grid_rowconfigure(2, weight=0)  # Total vendido
root.grid_rowconfigure(3, weight=0)  # Botões de exportação
root.grid_rowconfigure(4, weight=0)  # Botão "Ver Estoque"
root.grid_rowconfigure(5, weight=1)  # Treeview (expande apenas o necessário)

# Carregar e exibir logo
try:
    logo_img = Image.open("assets/logo.png")
    logo_img = logo_img.resize((300, 300))  # Aumentei o tamanho da logo
    logo = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(root, image=logo, bg="#f0f0f0")
    logo_label.grid(row=0, column=0, columnspan=2, pady=20, sticky="n")
except FileNotFoundError:
    messagebox.showwarning("Aviso", "Logo não encontrada!")

# Frame para adicionar/atualizar produtos
frame_adicionar = tk.Frame(root, bg="#f0f0f0")
frame_adicionar.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")

tk.Label(frame_adicionar, text="Nome:", bg="#f0f0f0", fg="#333333", font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=10)
entry_nome = tk.Entry(frame_adicionar, width=40, bg="#e0e0e0", fg="#333333", font=("Arial", 12))
entry_nome.grid(row=0, column=1, pady=10, sticky="ew")

tk.Label(frame_adicionar, text="Quantidade:", bg="#f0f0f0", fg="#333333", font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=10)
entry_quantidade = tk.Entry(frame_adicionar, width=40, bg="#e0e0e0", fg="#333333", font=("Arial", 12))
entry_quantidade.grid(row=1, column=1, pady=10, sticky="ew")

tk.Label(frame_adicionar, text="Preço:", bg="#f0f0f0", fg="#333333", font=("Arial", 12)).grid(row=2, column=0, sticky="w", pady=10)
entry_preco = tk.Entry(frame_adicionar, width=40, bg="#e0e0e0", fg="#333333", font=("Arial", 12))
entry_preco.grid(row=2, column=1, pady=10, sticky="ew")

tk.Button(frame_adicionar, text="Adicionar Produto", command=adicionar_produto, bg="#4CAF50", fg="white", font=("Arial", 12)).grid(row=3, columnspan=2, pady=20, sticky="ew")
tk.Button(frame_adicionar, text="Atualizar Produto", command=atualizar_produto, bg="#2196F3", fg="white", font=("Arial", 12)).grid(row=4, columnspan=2, pady=20, sticky="ew")
tk.Button(frame_adicionar, text="Excluir Produto", command=excluir_produto, bg="#f44336", fg="white", font=("Arial", 12)).grid(row=5, columnspan=2, pady=20, sticky="ew")

# Frame para vender produtos
frame_venda = tk.Frame(root, bg="#f0f0f0")
frame_venda.grid(row=1, column=1, padx=20, pady=20, sticky="nsew")

tk.Label(frame_venda, text="Quantidade para Vender:", bg="#f0f0f0", fg="#333333", font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=10)
entry_venda_quantidade = tk.Entry(frame_venda, width=40, bg="#e0e0e0", fg="#333333", font=("Arial", 12))
entry_venda_quantidade.grid(row=0, column=1, pady=10, sticky="ew")

tk.Button(frame_venda, text="Vender Produto", command=vender_produto, bg="#FFC107", fg="white", font=("Arial", 12)).grid(row=1, columnspan=2, pady=20, sticky="ew")

# Exibir o total vendido
label_total_vendido = tk.Label(root, text="Total Vendido: R$ 0,00", bg="#f0f0f0", fg="#333333", font=("Arial", 14))
label_total_vendido.grid(row=2, column=0, columnspan=2, pady=20, sticky="n")

# Botão para exportar Excel
tk.Button(root, text="Exportar para Excel", command=exportar_excel, bg="#4CAF50", fg="white", font=("Arial", 12)).grid(row=3, column=0, columnspan=2, pady=20, sticky="n")

# Botão para abrir a janela de estoque
tk.Button(root, text="Ver Estoque", command=abrir_janela_estoque, bg="#2196F3", fg="white", font=("Arial", 12)).grid(row=4, column=0, columnspan=2, pady=20, sticky="n")

# Frame para pesquisa
frame_pesquisa = tk.Frame(root, bg="#f0f0f0")
frame_pesquisa.grid(row=5, column=0, columnspan=2, padx=20, pady=20, sticky="ew")

tk.Label(frame_pesquisa, text="Pesquisar Produto:", bg="#f0f0f0", fg="#333333", font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=10)
entry_pesquisa = tk.Entry(frame_pesquisa, width=40, bg="#e0e0e0", fg="#333333", font=("Arial", 12))
entry_pesquisa.grid(row=0, column=1, pady=10, sticky="ew")

tk.Button(frame_pesquisa, text="Pesquisar", command=pesquisar_produto, bg="#9C27B0", fg="white", font=("Arial", 12)).grid(row=0, column=2, padx=10, sticky="ew")

# Frame para conter a Treeview e a barra de rolagem
frame_tree = tk.Frame(root)
frame_tree.grid(row=6, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")

# Treeview para listar produtos (apenas nome e quantidade)
tree = ttk.Treeview(frame_tree, columns=("Nome", "Quantidade"), show="headings")
tree.pack(side="left", fill="both", expand=True)

# Barra de rolagem
scrollbar = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
scrollbar.pack(side="right", fill="y")
tree.configure(yscrollcommand=scrollbar.set)

tree.heading("Nome", text="Nome")
tree.heading("Quantidade", text="Quantidade")

# Configurar o grid da Treeview para expandir
root.grid_rowconfigure(6, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Carregar lista de produtos
listar_produtos()

# Criar a tabela se não existir
criar_tabela()

# Rodar o aplicativo
root.mainloop()