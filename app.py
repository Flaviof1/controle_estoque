import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import locale
from openpyxl import Workbook
import os
from datetime import datetime
from contextlib import contextmanager

# Configuração de localidade para formato de moeda
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class SistemaEstoque:
    def __init__(self, root):
        self.root = root
        self.carregar_logo()
        self.configurar_interface()
        self.criar_banco_dados()
        
    def carregar_logo(self):
        """Carrega e exibe a logo na interface."""
        try:
            logo_path = os.path.join("assets", "logo.png")
            
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((200, 100), Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                
                logo_frame = ttk.Frame(self.root)
                logo_frame.pack(fill=tk.X, pady=10)
                
                logo_label = ttk.Label(logo_frame, image=self.logo_img)
                logo_label.pack()
            else:
                print(f"Logo não encontrada em: {logo_path}")
        except Exception as e:
            print(f"Erro ao carregar logo: {e}")

    @contextmanager
    def conectar_banco(self):
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
        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS produtos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    quantidade INTEGER NOT NULL,
                    preco_custo REAL NOT NULL
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS vendas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    produto_id INTEGER NOT NULL,
                    nome_produto TEXT NOT NULL,
                    quantidade INTEGER NOT NULL,
                    preco_venda REAL NOT NULL,
                    preco_custo REAL NOT NULL,
                    data_venda TEXT NOT NULL,
                    FOREIGN KEY (produto_id) REFERENCES produtos (id)
                )
            ''')
            
            conn.commit()

    def configurar_interface(self):
        """Configura a interface gráfica principal."""
        self.root.title("Sistema de Controle de Estoque")
        self.root.geometry("1200x800")
        
        # Notebook (abas)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Abas
        self.configurar_aba_operacoes()
        self.configurar_aba_visualizacao()
        self.configurar_aba_relatorios()
        
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.atualizar_abas())

    def configurar_aba_operacoes(self):
        """Configura a aba de operações."""
        self.aba_operacoes = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_operacoes, text="Operações")
        
        # Frame de cadastro
        cadastro_frame = ttk.LabelFrame(self.aba_operacoes, text="Cadastro de Produtos", padding=15)
        cadastro_frame.pack(fill=tk.X, pady=10, padx=10)
        
        # Campos do formulário
        ttk.Label(cadastro_frame, text="Nome:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_nome = ttk.Entry(cadastro_frame, width=40)
        self.entry_nome.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(cadastro_frame, text="Quantidade:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_quantidade = ttk.Entry(cadastro_frame, width=10)
        self.entry_quantidade.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(cadastro_frame, text="Preço de Custo:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.entry_preco_custo = ttk.Entry(cadastro_frame, width=10)
        self.entry_preco_custo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # Botões de ação
        btn_frame = ttk.Frame(cadastro_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        ttk.Button(btn_frame, text="Adicionar", command=self.adicionar_produto, style='Success.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Atualizar", command=self.atualizar_produto, style='Primary.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Excluir", command=self.excluir_produto, style='Danger.TButton').pack(side=tk.LEFT, padx=5)
        
        # Frame de vendas
        venda_frame = ttk.LabelFrame(self.aba_operacoes, text="Registrar Venda", padding=15)
        venda_frame.pack(fill=tk.X, pady=10, padx=10)
        
        ttk.Label(venda_frame, text="Quantidade:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_qtde_venda = ttk.Entry(venda_frame, width=10)
        self.entry_qtde_venda.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(venda_frame, text="Preço de Venda:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_preco_venda = ttk.Entry(venda_frame, width=10)
        self.entry_preco_venda.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Button(venda_frame, text="Registrar Venda", command=self.registrar_venda, style='Primary.TButton').grid(row=2, column=0, columnspan=2, pady=10)
        
        # Frame de pesquisa
        pesquisa_frame = ttk.Frame(self.aba_operacoes)
        pesquisa_frame.pack(fill=tk.X, pady=10, padx=10)
        
        ttk.Label(pesquisa_frame, text="Pesquisar:").pack(side=tk.LEFT)
        self.entry_pesquisa = ttk.Entry(pesquisa_frame, width=30)
        self.entry_pesquisa.pack(side=tk.LEFT, padx=5)
        ttk.Button(pesquisa_frame, text="Buscar", command=self.buscar_produtos, style='Accent.TButton').pack(side=tk.LEFT)
        ttk.Button(pesquisa_frame, text="Limpar", command=self.limpar_pesquisa, style='TButton').pack(side=tk.LEFT, padx=5)
        
        # Tabela de produtos
        tree_frame = ttk.Frame(self.aba_operacoes)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        colunas = ("ID", "Nome", "Quantidade", "Preço Custo")
        self.tree_produtos = ttk.Treeview(tree_frame, columns=colunas, show="headings", height=15)
        
        for col in colunas:
            self.tree_produtos.heading(col, text=col)
            self.tree_produtos.column(col, width=120, anchor=tk.CENTER)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_produtos.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree_produtos.xview)
        self.tree_produtos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree_produtos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Carregar dados iniciais
        self.listar_produtos()

    def configurar_aba_visualizacao(self):
        """Configura a aba de visualização."""
        self.aba_visualizacao = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_visualizacao, text="Visualização")
        
        main_frame = ttk.Frame(self.aba_visualizacao)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tabela de estoque
        colunas = ("ID", "Nome", "Quantidade", "Preço Custo", "Valor Total")
        self.tree_estoque = ttk.Treeview(main_frame, columns=colunas, show="headings", height=15)
        
        for col in colunas:
            self.tree_estoque.heading(col, text=col)
            self.tree_estoque.column(col, width=120, anchor=tk.CENTER)
        
        # Barras de rolagem
        scroll_y = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree_estoque.yview)
        scroll_x = ttk.Scrollbar(main_frame, orient=tk.HORIZONTAL, command=self.tree_estoque.xview)
        self.tree_estoque.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree_estoque.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Frame de resumo
        resumo_frame = ttk.LabelFrame(self.aba_visualizacao, text="Resumo Financeiro", padding=15)
        resumo_frame.pack(fill=tk.X, pady=10, padx=10)
        
        # Rótulos de totais
        ttk.Label(resumo_frame, text="Totais:", style='Title.TLabel').pack(anchor=tk.W, pady=5)
        
        self.lbl_total_estoque = ttk.Label(resumo_frame, 
                                         text="Valor Total em Estoque: R$ 0,00", 
                                         style='Info.TLabel')
        self.lbl_total_estoque.pack(anchor=tk.W, pady=2)
        
        self.lbl_total_vendas = ttk.Label(resumo_frame, 
                                        text="Total em Vendas: R$ 0,00", 
                                        style='Info.TLabel')
        self.lbl_total_vendas.pack(anchor=tk.W, pady=2)
        
        self.lbl_lucro_total = ttk.Label(resumo_frame, 
                                       text="Lucro Total: R$ 0,00", 
                                       style='Success.TLabel')
        self.lbl_lucro_total.pack(anchor=tk.W, pady=2)
        
        # Botão de exportação
        ttk.Button(resumo_frame, 
                 text="Exportar Relatório Completo", 
                 command=self.exportar_relatorio, 
                 style='Primary.TButton').pack(pady=10)

    def configurar_aba_relatorios(self):
        """Configura a aba de relatórios."""
        self.aba_relatorios = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_relatorios, text="Relatórios")
        
        main_frame = ttk.Frame(self.aba_relatorios)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tabela de vendas
        colunas = ("ID", "Produto", "Qtd", "Custo", "Venda", 
                 "Total Custo", "Total Venda", "Lucro", "Data")
        self.tree_vendas = ttk.Treeview(main_frame, columns=colunas, show="headings", height=20)
        
        # Configurar colunas
        for col in colunas:
            self.tree_vendas.heading(col, text=col)
            self.tree_vendas.column(col, width=100, anchor=tk.CENTER)
        
        # Barras de rolagem
        scroll_y = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree_vendas.yview)
        scroll_x = ttk.Scrollbar(main_frame, orient=tk.HORIZONTAL, command=self.tree_vendas.xview)
        self.tree_vendas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree_vendas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Frame de totais
        totais_frame = ttk.LabelFrame(self.aba_relatorios, text="Totais de Vendas", padding=15)
        totais_frame.pack(fill=tk.X, pady=10, padx=10)
        
        self.lbl_total_vendas = ttk.Label(totais_frame, 
                                        text="Total em Vendas: R$ 0,00",
                                        style='Info.TLabel')
        self.lbl_total_vendas.pack(side=tk.LEFT, padx=10)
        
        self.lbl_total_custo = ttk.Label(totais_frame, 
                                       text="Total em Custos: R$ 0,00",
                                       style='Info.TLabel')
        self.lbl_total_custo.pack(side=tk.LEFT, padx=10)
        
        self.lbl_total_lucro = ttk.Label(totais_frame, 
                                       text="Lucro Total: R$ 0,00",
                                       style='Success.TLabel')
        self.lbl_total_lucro.pack(side=tk.LEFT, padx=10)

    def atualizar_abas(self):
        """Atualiza todas as abas quando há mudança."""
        aba_atual = self.notebook.index(self.notebook.select())
        
        if aba_atual == 1:  # Aba de Visualização
            self.atualizar_aba_visualizacao()
        elif aba_atual == 2:  # Aba de Relatórios
            self.atualizar_aba_relatorios()

    def atualizar_aba_visualizacao(self):
        """Atualiza a aba de visualização com dados recentes."""
        # Limpar dados existentes
        for item in self.tree_estoque.get_children():
            self.tree_estoque.delete(item)
        
        # Inicializar totais
        total_estoque = 0
        total_vendas = 0
        lucro_total = 0
        
        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            
            # Carregar produtos
            cursor.execute("SELECT id, nome, quantidade, preco_custo FROM produtos")
            for produto in cursor.fetchall():
                id_prod, nome, qtd, preco = produto
                valor_total = qtd * preco
                total_estoque += valor_total
                
                self.tree_estoque.insert("", "end", values=(
                    id_prod, 
                    nome, 
                    qtd, 
                    locale.currency(preco, grouping=True),
                    locale.currency(valor_total, grouping=True)
                ))
            
            # Calcular totais de vendas
            cursor.execute('''
                SELECT 
                    SUM(quantidade * preco_venda),
                    SUM(quantidade * preco_custo),
                    SUM(quantidade * (preco_venda - preco_custo))
                FROM vendas
            ''')
            dados_vendas = cursor.fetchone()
            
            if dados_vendas and dados_vendas[0]:
                total_vendas = dados_vendas[0]
                custo_total = dados_vendas[1]
                lucro_total = dados_vendas[2]
        
        # Atualizar rótulos
        self.lbl_total_estoque.config(
            text=f"Valor Total em Estoque: {locale.currency(total_estoque, grouping=True)}")
        
        self.lbl_total_vendas.config(
            text=f"Total em Vendas: {locale.currency(total_vendas, grouping=True)}")
        
        self.lbl_lucro_total.config(
            text=f"Lucro Total: {locale.currency(lucro_total, grouping=True)}")

    def atualizar_aba_relatorios(self):
        """Atualiza a aba de relatórios com dados recentes."""
        # Limpar dados existentes
        for item in self.tree_vendas.get_children():
            self.tree_vendas.delete(item)
        
        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            
            # Obter totais de vendas
            cursor.execute('''
                SELECT 
                    SUM(quantidade * preco_venda),
                    SUM(quantidade * preco_custo),
                    SUM(quantidade * (preco_venda - preco_custo))
                FROM vendas
            ''')
            totais = cursor.fetchone()
            
            total_vendas = totais[0] or 0 if totais else 0
            total_custo = totais[1] or 0 if totais else 0
            lucro_total = totais[2] or 0 if totais else 0
            
            # Carregar dados de vendas
            cursor.execute('''
                SELECT 
                    id, nome_produto, quantidade, preco_custo, preco_venda,
                    (quantidade * preco_custo),
                    (quantidade * preco_venda),
                    (quantidade * (preco_venda - preco_custo)),
                    data_venda
                FROM vendas
                ORDER BY data_venda DESC
            ''')
            
            for venda in cursor.fetchall():
                self.tree_vendas.insert("", "end", values=(
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
        
        # Atualizar totais
        self.lbl_total_vendas.config(
            text=f"Total em Vendas: {locale.currency(total_vendas, grouping=True)}")
        
        self.lbl_total_custo.config(
            text=f"Total em Custos: {locale.currency(total_custo, grouping=True)}")
        
        self.lbl_total_lucro.config(
            text=f"Lucro Total: {locale.currency(lucro_total, grouping=True)}")

    def adicionar_produto(self):
        """Adiciona um novo produto ao estoque."""
        nome = self.entry_nome.get().strip()
        quantidade = self.entry_quantidade.get().strip()
        preco_custo = self.entry_preco_custo.get().strip()

        if not nome or not quantidade or not preco_custo:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return
            
        if not quantidade.isdigit() or int(quantidade) <= 0:
            messagebox.showerror("Erro", "Quantidade deve ser um número positivo!")
            return
            
        try:
            preco_custo = float(preco_custo)
            if preco_custo <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Preço de custo inválido!")
            return

        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute(
                    "INSERT INTO produtos (nome, quantidade, preco_custo) VALUES (?, ?, ?)",
                    (nome, int(quantidade), preco_custo))
                conn.commit()
                messagebox.showinfo("Sucesso", "Produto adicionado com sucesso!")
                self.limpar_campos()
                self.listar_produtos()
                self.atualizar_abas()
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao adicionar produto: {e}")

    def atualizar_produto(self):
        """Atualiza um produto existente."""
        produto = self.obter_produto_selecionado()
        if not produto:
            messagebox.showerror("Erro", "Selecione um produto para atualizar!")
            return
            
        produto_id = produto[0]
        nome = self.entry_nome.get().strip()
        quantidade = self.entry_quantidade.get().strip()
        preco_custo = self.entry_preco_custo.get().strip()

        if not nome or not quantidade or not preco_custo:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return
            
        if not quantidade.isdigit() or int(quantidade) <= 0:
            messagebox.showerror("Erro", "Quantidade deve ser um número positivo!")
            return
            
        try:
            preco_custo = float(preco_custo)
            if preco_custo <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Preço de custo inválido!")
            return

        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute(
                    "UPDATE produtos SET nome = ?, quantidade = ?, preco_custo = ? WHERE id = ?",
                    (nome, int(quantidade), preco_custo, produto_id))
                conn.commit()
                messagebox.showinfo("Sucesso", "Produto atualizado com sucesso!")
                self.limpar_campos()
                self.listar_produtos()
                self.atualizar_abas()
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao atualizar produto: {e}")

    def excluir_produto(self):
        """Remove um produto do estoque."""
        produto = self.obter_produto_selecionado()
        if not produto:
            messagebox.showerror("Erro", "Selecione um produto para excluir!")
            return
            
        if not messagebox.askyesno("Confirmar", f"Tem certeza que deseja excluir '{produto[1]}'?"):
            return

        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute("DELETE FROM produtos WHERE id = ?", (produto[0],))
                conn.commit()
                messagebox.showinfo("Sucesso", "Produto excluído com sucesso!")
                self.listar_produtos()
                self.atualizar_abas()
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao excluir produto: {e}")

    def registrar_venda(self):
        """Registra uma venda de produto."""
        produto = self.obter_produto_selecionado()
        if not produto:
            messagebox.showerror("Erro", "Selecione um produto para vender!")
            return
            
        produto_id, nome, quantidade, preco_custo = produto
        qtde_venda = self.entry_qtde_venda.get().strip()
        preco_venda = self.entry_preco_venda.get().strip()

        if not qtde_venda.isdigit() or int(qtde_venda) <= 0:
            messagebox.showerror("Erro", "Quantidade inválida para venda!")
            return
            
        qtde_venda = int(qtde_venda)
        
        if qtde_venda > quantidade:
            messagebox.showerror("Erro", "Quantidade em estoque insuficiente!")
            return
            
        try:
            preco_venda = float(preco_venda)
            if preco_venda <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Preço de venda inválido!")
            return

        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            try:
                # Atualizar estoque
                nova_quantidade = quantidade - qtde_venda
                cursor.execute(
                    "UPDATE produtos SET quantidade = ? WHERE id = ?",
                    (nova_quantidade, produto_id))
                
                # Registrar venda
                data_venda = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute(
                    '''INSERT INTO vendas 
                    (produto_id, nome_produto, quantidade, preco_venda, preco_custo, data_venda)
                    VALUES (?, ?, ?, ?, ?, ?)''',
                    (produto_id, nome, qtde_venda, preco_venda, preco_custo, data_venda))
                
                conn.commit()
                messagebox.showinfo("Sucesso", 
                    f"Venda registrada:\n"
                    f"{qtde_venda} x {nome}\n"
                    f"Total: {locale.currency(qtde_venda * preco_venda, grouping=True)}")
                
                self.listar_produtos()
                self.limpar_campos_venda()
                self.atualizar_abas()
            except sqlite3.Error as e:
                conn.rollback()
                messagebox.showerror("Erro", f"Erro ao registrar venda: {e}")

    def obter_produto_selecionado(self):
        """Retorna o produto selecionado na tabela."""
        selecionado = self.tree_produtos.selection()
        if not selecionado:
            return None
            
        item = self.tree_produtos.item(selecionado[0])
        produto_id = item['values'][0]
        nome = item['values'][1]
        quantidade = item['values'][2]
        preco_custo = item['values'][3]
        
        return (produto_id, nome, quantidade, preco_custo)

    def listar_produtos(self, termo=None):
        """Lista os produtos na tabela."""
        self.tree_produtos.delete(*self.tree_produtos.get_children())
        
        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            
            if termo:
                cursor.execute(
                    '''SELECT id, nome, quantidade, preco_custo 
                    FROM produtos 
                    WHERE nome LIKE ? 
                    ORDER BY nome''',
                    (f"%{termo}%",))
            else:
                cursor.execute(
                    '''SELECT id, nome, quantidade, preco_custo 
                    FROM produtos 
                    ORDER BY nome''')
            
            for linha in cursor.fetchall():
                self.tree_produtos.insert("", "end", values=linha)

    def buscar_produtos(self):
        """Busca produtos pelo nome."""
        termo = self.entry_pesquisa.get().strip()
        self.listar_produtos(termo)

    def limpar_pesquisa(self):
        """Limpa a pesquisa e lista todos os produtos."""
        self.entry_pesquisa.delete(0, tk.END)
        self.listar_produtos()

    def limpar_campos(self):
        """Limpa os campos de cadastro."""
        self.entry_nome.delete(0, tk.END)
        self.entry_quantidade.delete(0, tk.END)
        self.entry_preco_custo.delete(0, tk.END)
        
    def limpar_campos_venda(self):
        """Limpa os campos de venda."""
        self.entry_qtde_venda.delete(0, tk.END)
        self.entry_preco_venda.delete(0, tk.END)

    def exportar_relatorio(self):
        """Exporta os dados para um arquivo Excel."""
        with self.conectar_banco() as conn:
            cursor = conn.cursor()
            
            # Criar arquivo Excel
            wb = Workbook()
            ws_produtos = wb.active
            ws_produtos.title = "Produtos"
            
            # Cabeçalhos produtos
            ws_produtos.append(["ID", "Nome", "Quantidade", "Preço Custo", "Valor Total"])
            
            # Dados produtos
            cursor.execute("SELECT id, nome, quantidade, preco_custo FROM produtos ORDER BY nome")
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
            ws_vendas.append(["ID", "Produto", "Quantidade", "Preço Custo", "Preço Venda",
                             "Total Custo", "Total Venda", "Lucro", "Data"])
            
            # Dados vendas
            cursor.execute('''
                SELECT id, nome_produto, quantidade, preco_custo, preco_venda,
                       (quantidade * preco_custo), (quantidade * preco_venda),
                       (quantidade * (preco_venda - preco_custo)), data_venda
                FROM vendas
                ORDER BY data_venda DESC
            ''')
            for venda in cursor.fetchall():
                ws_vendas.append(list(venda))
            
            # Salvar arquivo
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Arquivos Excel", "*.xlsx")],
                title="Salvar relatório como"
            )
            
            if caminho_arquivo:
                wb.save(caminho_arquivo)
                messagebox.showinfo("Sucesso", f"Dados exportados para:\n{caminho_arquivo}")

if __name__ == "__main__":
    root = tk.Tk()
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configurações de estilo
    style.configure('.', background='#f8f9fa', foreground='#495057', font=('Segoe UI', 10))
    style.configure('TButton', background='#4267b2', foreground='white', borderwidth=1, font=('Segoe UI', 10, 'bold'), padding=8)
    style.map('TButton', background=[('active', '#3b5998')], foreground=[('active', 'white')])
    style.configure('Primary.TButton', background='#4267b2', foreground='white')
    style.configure('Success.TButton', background='#4CAF50', foreground='white')
    style.configure('Danger.TButton', background='#f44336', foreground='white')
    style.configure('Accent.TButton', background='#8b9dc3', foreground='white')
    style.configure('Treeview', background='white', foreground='#333333', fieldbackground='white', rowheight=28, font=('Segoe UI', 9))
    style.configure('Treeview.Heading', background='#4267b2', foreground='white', font=('Segoe UI', 10, 'bold'))
    style.map('Treeview', background=[('selected', '#8b9dc3')], foreground=[('selected', 'white')])
    
    # Verificar se a pasta assets existe
    if not os.path.exists("assets"):
        os.makedirs("assets")
        print("Pasta assets criada - coloque sua logo.png nela")
    
    app = SistemaEstoque(root)
    root.mainloop()