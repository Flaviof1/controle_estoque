import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from PIL import Image, ImageTk
import locale
from openpyxl import Workbook
from contextlib import contextmanager
import os
from datetime import datetime

# Configuração de localidade para formato de moeda
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class EstoqueApp:
    def __init__(self, root):
        self.root = root
        self.setup_ui()
        self.criar_banco_dados()
        
    @contextmanager
    def conectar(self):
        """Gerenciador de contexto para conexão com o banco de dados."""
        conn = None
        try:
            conn = sqlite3.connect('estoque.db')
            yield conn
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro de conexão com o banco de dados: {e}")
        finally:
            if conn:
                conn.close()
    
    def criar_banco_dados(self):
        """Cria as tabelas necessárias no banco de dados."""
        with self.conectar() as conn:
            cursor = conn.cursor()
            
            # Tabela de produtos
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS produtos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    quantidade INTEGER NOT NULL,
                    preco_compra REAL NOT NULL
                )
            ''')
            
            # Tabela de vendas (histórico)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS vendas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    produto_id INTEGER NOT NULL,
                    produto_nome TEXT NOT NULL,
                    quantidade INTEGER NOT NULL,
                    preco_venda REAL NOT NULL,
                    preco_compra REAL NOT NULL,
                    data_venda TEXT NOT NULL,
                    FOREIGN KEY (produto_id) REFERENCES produtos (id)
                )
            ''')
            
            conn.commit()

    def adicionar_produto(self):
        """Adiciona um novo produto ao estoque."""
        nome = self.entry_nome.get().strip()
        quantidade = self.entry_quantidade.get().strip()
        preco_compra = self.entry_preco_compra.get().strip()

        if not nome or not quantidade or not preco_compra:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return
            
        if not quantidade.isdigit() or int(quantidade) <= 0:
            messagebox.showerror("Erro", "Quantidade deve ser um número positivo!")
            return
            
        if not preco_compra.replace('.', '', 1).isdigit() or float(preco_compra) <= 0:
            messagebox.showerror("Erro", "Preço de compra inválido!")
            return

        with self.conectar() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute(
                    "INSERT INTO produtos (nome, quantidade, preco_compra) VALUES (?, ?, ?)",
                    (nome, int(quantidade), float(preco_compra)))
                conn.commit()
                messagebox.showinfo("Sucesso", "Produto adicionado com sucesso!")
                self.limpar_campos()
                self.listar_produtos()
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao adicionar produto: {e}")

    def vender_produto(self):
        """Realiza a venda de um produto."""
        produto = self.obter_produto_selecionado()
        if not produto:
            messagebox.showerror("Erro", "Selecione um produto para vender!")
            return
            
        id_produto, nome, quantidade, preco_compra = produto
        quantidade_venda = self.entry_venda_quantidade.get().strip()
        preco_venda = self.entry_preco_venda.get().strip()

        if not quantidade_venda.isdigit() or int(quantidade_venda) <= 0:
            messagebox.showerror("Erro", "Quantidade inválida para venda!")
            return
            
        quantidade_venda = int(quantidade_venda)
        
        if quantidade_venda > quantidade:
            messagebox.showerror("Erro", "Quantidade em estoque insuficiente!")
            return
            
        if not preco_venda.replace('.', '', 1).isdigit() or float(preco_venda) <= 0:
            messagebox.showerror("Erro", "Preço de venda inválido!")
            return
            
        preco_venda = float(preco_venda)

        with self.conectar() as conn:
            cursor = conn.cursor()
            try:
                # Atualiza o estoque
                nova_quantidade = quantidade - quantidade_venda
                cursor.execute(
                    "UPDATE produtos SET quantidade = ? WHERE id = ?",
                    (nova_quantidade, id_produto))
                
                # Registra a venda
                data_venda = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute(
                    '''INSERT INTO vendas 
                    (produto_id, produto_nome, quantidade, preco_venda, preco_compra, data_venda)
                    VALUES (?, ?, ?, ?, ?, ?)''',
                    (id_produto, nome, quantidade_venda, preco_venda, preco_compra, data_venda))
                
                conn.commit()
                messagebox.showinfo("Sucesso", 
                    f"Venda registrada:\n"
                    f"{quantidade_venda} x {nome}\n"
                    f"Total: {locale.currency(quantidade_venda * preco_venda, grouping=True)}")
                
                self.listar_produtos()
                self.limpar_campos_venda()
            except sqlite3.Error as e:
                conn.rollback()
                messagebox.showerror("Erro", f"Erro ao registrar venda: {e}")

    def obter_produto_selecionado(self):
        """Retorna os dados do produto selecionado na Treeview."""
        selected = self.tree.selection()
        if not selected:
            return None
            
        item = self.tree.item(selected[0])
        produto_id = item['values'][0]
        nome = item['values'][1]
        quantidade = item['values'][2]
        preco_compra = item['values'][3]
        
        return (produto_id, nome, quantidade, preco_compra)

    def listar_produtos(self, termo=None):
        """Lista os produtos na Treeview."""
        self.tree.delete(*self.tree.get_children())
        
        with self.conectar() as conn:
            cursor = conn.cursor()
            
            if termo:
                cursor.execute(
                    '''SELECT id, nome, quantidade, preco_compra 
                    FROM produtos 
                    WHERE nome LIKE ? 
                    ORDER BY nome''',
                    (f"%{termo}%",))
            else:
                cursor.execute(
                    '''SELECT id, nome, quantidade, preco_compra 
                    FROM produtos 
                    ORDER BY nome''')
            
            for row in cursor.fetchall():
                self.tree.insert("", "end", values=row)

    def abrir_relatorio_vendas(self):
        """Abre a janela com o relatório de vendas e lucros."""
        relatorio_window = tk.Toplevel(self.root)
        relatorio_window.title("Relatório de Vendas e Lucros")
        relatorio_window.geometry("1100x600")
        
        # Frame principal
        main_frame = tk.Frame(relatorio_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview para vendas
        columns = ("ID", "Produto", "Quantidade", "Preço Compra", "Preço Venda", 
                 "Total Compra", "Total Venda", "Lucro", "Data")
        tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=20)
        
        # Configurar colunas
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)
        
        # Barras de rolagem
        scroll_y = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
        scroll_x = ttk.Scrollbar(main_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Frame para totais
        total_frame = tk.Frame(relatorio_window)
        total_frame.pack(fill=tk.X, pady=10)
        
        # Carregar dados
        with self.conectar() as conn:
            cursor = conn.cursor()
            
            # Total de vendas
            cursor.execute('''
                SELECT 
                    SUM(quantidade * preco_venda) as total_vendas,
                    SUM(quantidade * preco_compra) as total_custo,
                    SUM(quantidade * (preco_venda - preco_compra)) as total_lucro
                FROM vendas
            ''')
            totais = cursor.fetchone()
            total_vendas = totais[0] or 0
            total_custo = totais[1] or 0
            total_lucro = totais[2] or 0
            
            # Listar vendas
            cursor.execute('''
                SELECT 
                    id, produto_nome, quantidade, preco_compra, preco_venda,
                    (quantidade * preco_compra) as total_compra,
                    (quantidade * preco_venda) as total_venda,
                    (quantidade * (preco_venda - preco_compra)) as lucro,
                    data_venda
                FROM vendas
                ORDER BY data_venda DESC
            ''')
            
            for venda in cursor.fetchall():
                tree.insert("", "end", values=(
                    venda[0],
                    venda[1],
                    venda[2],
                    locale.currency(venda[3], grouping=True),
                    locale.currency(venda[4], grouping=True),
                    locale.currency(venda[5], grouping=True),
                    locale.currency(venda[6], grouping=True),
                    locale.currency(venda[7], grouping=True),
                    venda[8]
                ))
        
        # Exibir totais
        tk.Label(total_frame, 
                text=f"Total em Vendas: {locale.currency(total_vendas, grouping=True)}",
                font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        
        tk.Label(total_frame, 
                text=f"Total em Custos: {locale.currency(total_custo, grouping=True)}",
                font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        
        tk.Label(total_frame, 
                text=f"Lucro Total: {locale.currency(total_lucro, grouping=True)}",
                font=("Arial", 10, "bold"), fg="green").pack(side=tk.LEFT, padx=10)

    def setup_ui(self):
        """Configura a interface gráfica principal."""
        self.root.title("Sistema de Estoque e Vendas")
        self.root.geometry("1000x700")
        
        # Carregar logo de fundo
        self.carregar_logo_fundo()
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame de cadastro
        cadastro_frame = tk.LabelFrame(main_frame, text="Cadastro de Produtos", 
                                     padx=10, pady=10, bg='white')
        cadastro_frame.pack(fill=tk.X, pady=5)
        
        # Campos de cadastro
        tk.Label(cadastro_frame, text="Nome:", bg='white').grid(row=0, column=0, sticky="w")
        self.entry_nome = tk.Entry(cadastro_frame, width=40)
        self.entry_nome.grid(row=0, column=1, padx=5, pady=2)
        
        tk.Label(cadastro_frame, text="Quantidade:", bg='white').grid(row=1, column=0, sticky="w")
        self.entry_quantidade = tk.Entry(cadastro_frame, width=10)
        self.entry_quantidade.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        tk.Label(cadastro_frame, text="Preço de Compra:", bg='white').grid(row=2, column=0, sticky="w")
        self.entry_preco_compra = tk.Entry(cadastro_frame, width=10)
        self.entry_preco_compra.grid(row=2, column=1, padx=5, pady=2, sticky="w")
        
        # Botões de cadastro
        btn_frame = tk.Frame(cadastro_frame, bg='white')
        btn_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        tk.Button(btn_frame, text="Adicionar", command=self.adicionar_produto, 
                 bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Atualizar", command=self.atualizar_produto, 
                 bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Excluir", command=self.excluir_produto, 
                 bg="#f44336", fg="white").pack(side=tk.LEFT, padx=5)
        
        # Frame de venda
        venda_frame = tk.LabelFrame(main_frame, text="Registrar Venda", 
                                  padx=10, pady=10, bg='white')
        venda_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(venda_frame, text="Quantidade:", bg='white').grid(row=0, column=0, sticky="w")
        self.entry_venda_quantidade = tk.Entry(venda_frame, width=10)
        self.entry_venda_quantidade.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        tk.Label(venda_frame, text="Preço de Venda:", bg='white').grid(row=1, column=0, sticky="w")
        self.entry_preco_venda = tk.Entry(venda_frame, width=10)
        self.entry_preco_venda.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        tk.Button(venda_frame, text="Registrar Venda", command=self.vender_produto, 
                 bg="#FF9800", fg="white").grid(row=2, column=0, columnspan=2, pady=5)
        
        # Frame de pesquisa
        pesquisa_frame = tk.Frame(main_frame, bg='white')
        pesquisa_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(pesquisa_frame, text="Pesquisar:", bg='white').pack(side=tk.LEFT)
        self.entry_pesquisa = tk.Entry(pesquisa_frame, width=30)
        self.entry_pesquisa.pack(side=tk.LEFT, padx=5)
        tk.Button(pesquisa_frame, text="Buscar", 
                 command=lambda: self.listar_produtos(self.entry_pesquisa.get()),
                 bg="#9C27B0", fg="white").pack(side=tk.LEFT)
        tk.Button(pesquisa_frame, text="Limpar", 
                 command=lambda: [self.entry_pesquisa.delete(0, tk.END), self.listar_produtos()],
                 bg="#607D8B", fg="white").pack(side=tk.LEFT, padx=5)
        
        # Treeview de produtos
        tree_frame = tk.Frame(main_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("ID", "Nome", "Quantidade", "Preço Compra")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Botões inferiores
        bottom_frame = tk.Frame(main_frame, bg='white')
        bottom_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(bottom_frame, text="Relatório de Vendas", command=self.abrir_relatorio_vendas,
                 bg="#673AB7", fg="white", width=20).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_frame, text="Exportar para Excel", command=self.exportar_excel,
                 bg="#009688", fg="white", width=20).pack(side=tk.LEFT, padx=5)
        
        # Carregar dados iniciais
        self.listar_produtos()

    def carregar_logo_fundo(self):
        """Carrega a logo como imagem de fundo."""
        try:
            # Caminhos possíveis para a logo
            caminhos = [
                os.path.join("assets", "logo.png"),
                "logo.png",
                os.path.join(os.path.dirname(__file__), "logo.png")
            ]
            
            logo_path = None
            for caminho in caminhos:
                if os.path.exists(caminho):
                    logo_path = caminho
                    break
            
            if not logo_path:
                print("Logo não encontrada nos seguintes locais:")
                for caminho in caminhos:
                    print(f"- {caminho}")
                return
            
            # Carregar e redimensionar a imagem
            img = Image.open(logo_path)
            img.thumbnail((600, 600))  # Tamanho grande para fundo
            
            # Converter para PhotoImage
            self.logo_bg = ImageTk.PhotoImage(img)
            
            # Criar label para a logo
            self.bg_label = tk.Label(self.root, image=self.logo_bg, bg='white')
            self.bg_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            
            # Configurar para ficar atrás de tudo
            self.bg_label.lower()
            
        except Exception as e:
            print(f"Erro ao carregar logo de fundo: {str(e)}")

    def limpar_campos(self):
        """Limpa os campos de cadastro."""
        self.entry_nome.delete(0, tk.END)
        self.entry_quantidade.delete(0, tk.END)
        self.entry_preco_compra.delete(0, tk.END)
        
    def limpar_campos_venda(self):
        """Limpa os campos de venda."""
        self.entry_venda_quantidade.delete(0, tk.END)
        self.entry_preco_venda.delete(0, tk.END)

    def atualizar_produto(self):
        """Atualiza um produto existente."""
        produto = self.obter_produto_selecionado()
        if not produto:
            messagebox.showerror("Erro", "Selecione um produto para atualizar!")
            return
            
        id_produto = produto[0]
        nome = self.entry_nome.get().strip()
        quantidade = self.entry_quantidade.get().strip()
        preco_compra = self.entry_preco_compra.get().strip()

        if not nome or not quantidade or not preco_compra:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return
            
        if not quantidade.isdigit() or int(quantidade) <= 0:
            messagebox.showerror("Erro", "Quantidade deve ser um número positivo!")
            return
            
        if not preco_compra.replace('.', '', 1).isdigit() or float(preco_compra) <= 0:
            messagebox.showerror("Erro", "Preço de compra inválido!")
            return

        with self.conectar() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute(
                    "UPDATE produtos SET nome = ?, quantidade = ?, preco_compra = ? WHERE id = ?",
                    (nome, int(quantidade), float(preco_compra), id_produto))
                conn.commit()
                messagebox.showinfo("Sucesso", "Produto atualizado com sucesso!")
                self.limpar_campos()
                self.listar_produtos()
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao atualizar produto: {e}")

    def excluir_produto(self):
        """Exclui um produto do estoque."""
        produto = self.obter_produto_selecionado()
        if not produto:
            messagebox.showerror("Erro", "Selecione um produto para excluir!")
            return
            
        if not messagebox.askyesno("Confirmar", f"Tem certeza que deseja excluir '{produto[1]}'?"):
            return

        with self.conectar() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute("DELETE FROM produtos WHERE id = ?", (produto[0],))
                conn.commit()
                messagebox.showinfo("Sucesso", "Produto excluído com sucesso!")
                self.listar_produtos()
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao excluir produto: {e}")

    def exportar_excel(self):
        """Exporta os dados para um arquivo Excel."""
        with self.conectar() as conn:
            cursor = conn.cursor()
            
            # Criar arquivo Excel
            wb = Workbook()
            ws_produtos = wb.active
            ws_produtos.title = "Produtos"
            
            # Cabeçalhos produtos
            ws_produtos.append(["ID", "Nome", "Quantidade", "Preço Compra", "Valor Total"])
            
            # Dados produtos
            cursor.execute("SELECT id, nome, quantidade, preco_compra FROM produtos ORDER BY nome")
            for produto in cursor.fetchall():
                ws_produtos.append([
                    produto[0],
                    produto[1],
                    produto[2],
                    produto[3],
                    produto[2] * produto[3]  # Valor total
                ])
            
            # Planilha de vendas
            ws_vendas = wb.create_sheet("Vendas")
            ws_vendas.append(["ID", "Produto", "Quantidade", "Preço Compra", "Preço Venda",
                             "Total Compra", "Total Venda", "Lucro", "Data"])
            
            # Dados vendas
            cursor.execute('''
                SELECT id, produto_nome, quantidade, preco_compra, preco_venda,
                       (quantidade * preco_compra), (quantidade * preco_venda),
                       (quantidade * (preco_venda - preco_compra)), data_venda
                FROM vendas
                ORDER BY data_venda DESC
            ''')
            for venda in cursor.fetchall():
                ws_vendas.append(list(venda))
            
            # Salvar arquivo
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Salvar relatório como"
            )
            
            if file_path:
                wb.save(file_path)
                messagebox.showinfo("Sucesso", f"Dados exportados para:\n{file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = EstoqueApp(root)
    root.mainloop()