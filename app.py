import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk
import tkinter as tk
import sqlite3
from datetime import datetime, timedelta
from tkcalendar import DateEntry
import pandas as pd
import csv

# ---------------------------- Main Application Class ---------------------------- #
class ControleEstoqueApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSUPAB - GERÊNCIA DE SUPRIMENTOS")
        self.root.geometry("1200x800")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.database_path = 'estoque.db'
        self.setup_database()
        self.items_dict = {}
        self.item_list = []
        self.load_items_from_database()
        self.create_widgets()
        self.load_data()

    def setup_database(self):
        """Creates the necessary tables and extra columns and sets up indexes for performance.
           Also creates a new 'comentarios' table for storing multiple comments per licitação."""
        self.conn = sqlite3.connect(self.database_path)
        self.cursor = self.conn.cursor()

        with self.conn:
            # Licitações table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS licitacoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    numero_processo TEXT NOT NULL,
                    nome_empresa TEXT NOT NULL,
                    dados_empresa TEXT,
                    tipo_produto TEXT,
                    item_solicitado TEXT NOT NULL,
                    saldo_ata INTEGER NOT NULL,
                    saldo_ata_inicial INTEGER,
                    vencimento_ata TEXT,
                    status_licitacao TEXT NOT NULL,
                    ultima_atualizacao TEXT,
                    estoque_disponivel REAL DEFAULT 0
                )
            ''')
            # OCs table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS ocs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    numero_oc TEXT NOT NULL,
                    data_aceite TEXT,
                    empresa_marca TEXT,
                    item_pi TEXT NOT NULL,
                    quantidade_oc INTEGER NOT NULL,
                    quantidade_recebida INTEGER DEFAULT 0,
                    quantidade_pendente INTEGER,
                    status_pedido TEXT NOT NULL,
                    status_arrecadacao TEXT
                )
            ''')
            # Items table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_name TEXT NOT NULL,
                    uf TEXT,
                    tipo_produto TEXT,
                    cmm REAL DEFAULT 0
                )
            ''')
            # Transações table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS transacoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_transacao TEXT,
                    item_name TEXT,
                    quantidade INTEGER
                )
            ''')
            # New Comentários table (for multiple comments per licitação)
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS comentarios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    licitacao_id INTEGER,
                    comentario TEXT,
                    data TEXT
                )
            ''')

        # Add extra columns to licitacoes if needed
        self.cursor.execute("PRAGMA table_info(licitacoes)")
        licitacoes_columns = set(col[1] for col in self.cursor.fetchall())
        if 'saldo_ata_inicial' not in licitacoes_columns:
            self.cursor.execute('ALTER TABLE licitacoes ADD COLUMN saldo_ata_inicial INTEGER')
            self.cursor.execute('UPDATE licitacoes SET saldo_ata_inicial=saldo_ata WHERE saldo_ata_inicial IS NULL')
        if 'ultima_atualizacao' not in licitacoes_columns:
            self.cursor.execute('ALTER TABLE licitacoes ADD COLUMN ultima_atualizacao TEXT')
        if 'estoque_disponivel' not in licitacoes_columns:
            self.cursor.execute('ALTER TABLE licitacoes ADD COLUMN estoque_disponivel REAL DEFAULT 0')
        if 'disp_manual' not in licitacoes_columns:
            self.cursor.execute('ALTER TABLE licitacoes ADD COLUMN disp_manual REAL')
        # (The old "comentario" column is no longer used; comments are in the new table.)
        if 'tipo_produto' not in licitacoes_columns:
            self.cursor.execute('ALTER TABLE licitacoes ADD COLUMN tipo_produto TEXT')

        self.cursor.execute("PRAGMA table_info(items)")
        items_columns = set(col[1] for col in self.cursor.fetchall())
        if 'cmm' not in items_columns:
            self.cursor.execute('ALTER TABLE items ADD COLUMN cmm REAL DEFAULT 0')

        # Create indexes to improve performance
        self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_item_solicitado ON licitacoes(item_solicitado)")
        self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_item_pi ON ocs(item_pi)")

        self.conn.commit()

    def format_number_br(self, number):
        return f"{number:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def create_widgets(self):
        self.tab_control = ctk.CTkTabview(self.root)
        self.tab_control.pack(expand=1, fill='both')

        self.painel_licitacoes = self.tab_control.add("Licitações")
        self.painel_ocs = self.tab_control.add("Ordens de Compra")
        self.painel_dashboard = self.tab_control.add("Dashboard")
        self.painel_insercao_rm = self.tab_control.add("Inserir Estoque Atualizado")
        # Archive tab removed

        self.create_painel_licitacoes()
        self.create_painel_ocs()
        self.create_painel_dashboard()
        self.create_painel_inserir_estoque()

    def validate_int(self, value_if_allowed):
        if value_if_allowed == "":
            return True
        try:
            value = int(value_if_allowed)
            return value >= 0
        except ValueError:
            return False

    def load_data(self):
        self.carregar_licitacoes()
        self.carregar_ocs()
        self.carregar_dashboard()

    def load_items_from_database(self):
        self.items_dict = {}
        self.item_list = []
        self.cursor.execute('SELECT item_name, uf, tipo_produto, cmm FROM items')
        for item_name, uf, tipo_produto, cmm in self.cursor.fetchall():
            self.item_list.append(item_name)
            self.items_dict[item_name] = {'uf': uf, 'tipo_produto': tipo_produto, 'cmm': cmm}
        self.item_list.sort()  # Alphabetical order

    # ---------------------------- Licitações Panel ---------------------------- #
    def create_painel_licitacoes(self):
        self.painel_licitacoes.grid_columnconfigure(0, weight=1)
        self.painel_licitacoes.grid_columnconfigure(1, weight=1)

        search_frame = ctk.CTkFrame(self.painel_licitacoes)
        search_frame.grid(row=0, column=0, columnspan=2, pady=5)
        ctk.CTkLabel(search_frame, text="Buscar:").pack(side='left', padx=5)
        self.entry_search_licitacoes = ctk.CTkEntry(search_frame)
        self.entry_search_licitacoes.pack(side='left', padx=5)
        ctk.CTkButton(search_frame, text="Buscar", command=self.search_licitacoes).pack(side='left', padx=5)
        ctk.CTkButton(search_frame, text="Limpar", command=self.clear_search_licitacoes).pack(side='left', padx=5)
        ctk.CTkButton(search_frame, text="Carregar Itens", command=self.load_item_database).pack(side='left', padx=5)

        ctk.CTkButton(
            self.painel_licitacoes,
            text="Editar CMM dos Itens",
            command=self.open_edit_cmm_window
        ).grid(row=13, column=0, columnspan=2, pady=5)

        labels = [
            "Número do Processo:", "Nome da Empresa:", "Dados da Empresa:",
            "Tipo de Produto:", "Item Solicitado:", "Saldo da Ata:",
            "Vencimento da Ata:", "Status da Licitação:"
        ]
        self.entries_licitacoes = {}

        for idx, text in enumerate(labels):
            ctk.CTkLabel(self.painel_licitacoes, text=text).grid(row=idx+1, column=0, padx=5, pady=5, sticky='e')
            if text == "Tipo de Produto:":
                widget = ctk.CTkComboBox(self.painel_licitacoes, values=["FRIGORIFICADOS", "SECOS"])
                widget.set("FRIGORIFICADOS")
                widget.bind("<<ComboboxSelected>>", lambda event: self.filter_items_by_tipo())
            elif text == "Item Solicitado:":
                self.item_solicitado_combobox = ttk.Combobox(self.painel_licitacoes, values=self.item_list, width=50)
                self.item_solicitado_combobox.bind(
                    "<<ComboboxSelected>>",
                    lambda event: self.update_uf_label(self.item_solicitado_combobox.get())
                )
                widget = self.item_solicitado_combobox
            elif text == "Saldo da Ata:":
                widget = ctk.CTkEntry(self.painel_licitacoes)
                widget.configure(validate='key')
                widget.configure(validatecommand=(self.root.register(self.validate_int), '%P'))
            elif text == "Vencimento da Ata:":
                widget = DateEntry(self.painel_licitacoes, date_pattern='dd/MM/yyyy')
            elif text == "Status da Licitação:":
                widget = ctk.CTkComboBox(self.painel_licitacoes, values=["CSupAB", "COMRJ", "EMPRESA", "ASSINADO"])
            else:
                widget = ctk.CTkEntry(self.painel_licitacoes)

            widget.grid(row=idx+1, column=1, padx=5, pady=5, sticky='w')
            self.entries_licitacoes[text] = widget

            if text == "Saldo da Ata:":
                ctk.CTkLabel(self.painel_licitacoes, text="UF:").grid(
                    row=idx+2,
                    column=0,
                    padx=5,
                    pady=5,
                    sticky='e'
                )
                self.label_uf = ctk.CTkLabel(self.painel_licitacoes, text="")
                self.label_uf.grid(row=idx+2, column=1, padx=5, pady=5, sticky='w')

        ctk.CTkButton(
            self.painel_licitacoes,
            text="Adicionar Licitação",
            command=self.adicionar_licitacao
        ).grid(row=10, column=0, columnspan=2, padx=5, pady=5)

        columns = ('ID', 'Número do Processo', 'Nome da Empresa', 'Tipo de Produto', 'Item Solicitado', 'Saldo da Ata', 'Status')
        self.tree_licitacoes, scrollbar = self.create_treeview(self.painel_licitacoes, columns)
        self.tree_licitacoes.grid(row=11, column=0, columnspan=2, padx=5, pady=5, sticky='nsew')
        scrollbar.grid(row=11, column=2, sticky='ns')
        self.painel_licitacoes.grid_rowconfigure(11, weight=1)

        btn_frame = ctk.CTkFrame(self.painel_licitacoes)
        btn_frame.grid(row=12, column=0, columnspan=2, pady=5)
        ctk.CTkButton(btn_frame, text="Editar", command=self.editar_licitacao).pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="Excluir", command=self.excluir_licitacao).pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="Exportar", command=self.exportar_licitacoes).pack(side='left', padx=5)

    def open_edit_cmm_window(self):
        cmm_window = ctk.CTkToplevel(self.root)
        cmm_window.title("Editar CMM dos Itens")
        cmm_window.grab_set()

        columns = ('Item', 'Tipo de Produto', 'CMM')
        tree_items, scrollbar = self.create_treeview(cmm_window, columns)
        tree_items.pack(side='left', expand=1, fill='both')
        scrollbar.pack(side='right', fill='y')

        for item_name in self.item_list:
            item_info = self.items_dict[item_name]
            tipo_produto = item_info['tipo_produto']
            cmm = item_info.get('cmm', 0)
            tree_items.insert('', 'end', values=(item_name, tipo_produto, cmm))

        tree_items.bind('<Double-1>', lambda event: self.edit_cmm(tree_items, event))

    def edit_cmm(self, tree, event):
        selected_item = tree.selection()
        if selected_item:
            item = tree.item(selected_item)
            item_name = item['values'][0]
            self.open_edit_cmm_entry_window(item_name)
        else:
            messagebox.showwarning("Aviso", "Selecione um item para editar o CMM.")

    def create_treeview(self, parent, columns):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background="#f0f0f0", foreground="#000000", rowheight=25, fieldbackground="#f0f0f0")
        style.map("Treeview", background=[("selected", "#3399FF")], foreground=[("selected", "#FFFFFF")])
        tree = ttk.Treeview(parent, columns=columns, show='headings', selectmode='browse')
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor='center', width=100)
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        return tree, scrollbar

    def load_item_database(self):
        file_path = filedialog.askopenfilename(
            title="Selecione a base de dados de itens",
            filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)
                if 'Item' in df.columns and 'UF' in df.columns and 'Tipo de Produto' in df.columns:
                    if messagebox.askyesno("Confirmação", "Deseja substituir a lista de itens existente?"):
                        self.cursor.execute('DELETE FROM items')
                        for index, row in df.iterrows():
                            item_name = row['Item']
                            uf = row['UF']
                            tipo_produto = row['Tipo de Produto']
                            self.cursor.execute(
                                'INSERT INTO items (item_name, uf, tipo_produto) VALUES (?, ?, ?)',
                                (item_name, uf, tipo_produto)
                            )
                        self.conn.commit()
                        self.load_items_from_database()
                        self.filter_items_by_tipo()
                        messagebox.showinfo("Sucesso", f"{len(self.item_list)} itens carregados.")
                    else:
                        messagebox.showinfo("Operação cancelada", "Carregamento de itens cancelado.")
                else:
                    messagebox.showerror(
                        "Erro",
                        "As colunas 'Item', 'UF' e 'Tipo de Produto' não foram encontradas na planilha."
                    )
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao carregar a planilha:\n{e}")

    def filter_items_by_tipo(self):
        tipo_produto = self.entries_licitacoes["Tipo de Produto:"].get()
        if tipo_produto:
            filtered_items = [
                item_name for item_name, item_info in self.items_dict.items()
                if item_info.get('tipo_produto') == tipo_produto
            ]
        else:
            filtered_items = self.item_list
        filtered_items.sort()
        self.item_solicitado_combobox['values'] = filtered_items
        self.item_solicitado_combobox.set('')
        self.label_uf.configure(text="")

    def update_uf_label(self, item_name):
        uf = self.items_dict.get(item_name, {}).get('uf', '')
        self.label_uf.configure(text=uf)

    def search_licitacoes(self):
        query = self.entry_search_licitacoes.get()
        self.tree_licitacoes.delete(*self.tree_licitacoes.get_children())
        self.cursor.execute(
            '''
            SELECT id, numero_processo, nome_empresa, tipo_produto, item_solicitado, saldo_ata, status_licitacao
            FROM licitacoes
            WHERE numero_processo LIKE ? OR nome_empresa LIKE ? OR item_solicitado LIKE ?
            ''',
            (f'%{query}%', f'%{query}%', f'%{query}%')
        )
        for row in self.cursor.fetchall():
            self.tree_licitacoes.insert('', 'end', values=row)

    def clear_search_licitacoes(self):
        self.entry_search_licitacoes.delete(0, 'end')
        self.carregar_licitacoes()

    def adicionar_licitacao(self):
        entries = self.entries_licitacoes
        numero_processo = entries["Número do Processo:"].get()
        nome_empresa = entries["Nome da Empresa:"].get()
        dados_empresa = entries["Dados da Empresa:"].get()
        tipo_produto = entries["Tipo de Produto:"].get()
        item_solicitado = entries["Item Solicitado:"].get()
        saldo_ata = entries["Saldo da Ata:"].get()
        vencimento_ata = entries["Vencimento da Ata:"].get_date()
        status_licitacao = entries["Status da Licitação:"].get()

        if numero_processo and nome_empresa and tipo_produto and item_solicitado and saldo_ata and status_licitacao:
            if vencimento_ata < datetime.now().date():
                messagebox.showerror("Erro", "O vencimento da ata não pode ser anterior à data atual.")
                return
            try:
                saldo_ata_int = int(saldo_ata)
                if saldo_ata_int <= 0:
                    messagebox.showerror("Erro", "Saldo da Ata deve ser um número inteiro positivo.")
                    return
                vencimento_ata_str = vencimento_ata.strftime('%d/%m/%Y')
                self.cursor.execute(
                    '''
                    INSERT INTO licitacoes (
                        numero_processo, nome_empresa, dados_empresa, tipo_produto,
                        item_solicitado, saldo_ata, saldo_ata_inicial, vencimento_ata, status_licitacao
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (
                        numero_processo, nome_empresa, dados_empresa, tipo_produto,
                        item_solicitado, saldo_ata_int, saldo_ata_int,
                        vencimento_ata_str, status_licitacao
                    )
                )
                self.conn.commit()
                for entry in entries.values():
                    if isinstance(entry, DateEntry):
                        entry.set_date(datetime.now())
                    elif isinstance(entry, ctk.CTkComboBox) or isinstance(entry, ttk.Combobox):
                        entry.set('')
                    else:
                        entry.delete(0, 'end')
                self.label_uf.configure(text="")
                self.carregar_licitacoes()
                self.carregar_dashboard()
                messagebox.showinfo("Sucesso", "Licitação adicionada com sucesso.", parent=self.root)
            except ValueError:
                messagebox.showerror("Erro", "Saldo da Ata deve ser um número inteiro.")
        else:
            messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios.")

    def carregar_licitacoes(self):
        self.tree_licitacoes.delete(*self.tree_licitacoes.get_children())
        self.cursor.execute(
            '''
            SELECT id, numero_processo, nome_empresa, tipo_produto, item_solicitado, saldo_ata, status_licitacao
            FROM licitacoes
            '''
        )
        for row in self.cursor.fetchall():
            self.tree_licitacoes.insert('', 'end', values=row)

    def editar_licitacao(self):
        selected_item = self.tree_licitacoes.selection()
        if selected_item:
            item = self.tree_licitacoes.item(selected_item)
            licitacao_id = item['values'][0]
            self.open_edit_licitacao_window(licitacao_id)
        else:
            messagebox.showwarning("Aviso", "Selecione uma licitação para editar.")

    def open_edit_licitacao_window(self, licitacao_id):
        edit_window = ctk.CTkToplevel(self.root)
        edit_window.title("Editar Licitação")
        edit_window.grab_set()

        self.cursor.execute(
            '''
            SELECT numero_processo, nome_empresa, dados_empresa, tipo_produto,
                   item_solicitado, saldo_ata, vencimento_ata, status_licitacao
            FROM licitacoes
            WHERE id=?
            ''',
            (licitacao_id,)
        )
        licitacao = self.cursor.fetchone()

        labels = [
            "Número do Processo:", "Nome da Empresa:", "Dados da Empresa:",
            "Tipo de Produto:", "Item Solicitado:", "Saldo da Ata:",
            "Vencimento da Ata:", "Status da Licitação:"
        ]
        entries = {}

        for idx, text in enumerate(labels):
            ctk.CTkLabel(edit_window, text=text).grid(row=idx, column=0, padx=5, pady=5)
            if text == "Tipo de Produto:":
                widget = ctk.CTkComboBox(edit_window, values=["FRIGORIFICADOS", "SECOS"])
                widget.set(licitacao[3])
                widget.bind("<<ComboboxSelected>>", lambda event: self.filter_edit_items_by_tipo(entries))
            elif text == "Item Solicitado:":
                item_combobox = ttk.Combobox(edit_window, width=50)
                widget = item_combobox
                entries['item_combobox'] = item_combobox
            elif text == "Saldo da Ata:":
                widget = ctk.CTkEntry(edit_window)
                widget.configure(validate='key')
                widget.configure(validatecommand=(self.root.register(self.validate_int), '%P'))
            elif text == "Vencimento da Ata:":
                widget = DateEntry(edit_window, date_pattern='dd/MM/yyyy')
            elif text == "Status da Licitação:":
                widget = ctk.CTkComboBox(edit_window, values=["CSupAB", "COMRJ", "EMPRESA", "ASSINADO"])
            else:
                widget = ctk.CTkEntry(edit_window)

            if isinstance(widget, (ctk.CTkEntry, ttk.Entry)):
                widget.insert(0, licitacao[idx])
            elif isinstance(widget, ctk.CTkComboBox):
                widget.set(licitacao[idx])
            elif isinstance(widget, DateEntry):
                date_value = datetime.strptime(licitacao[idx], '%d/%m/%Y')
                widget.set_date(date_value)
            widget.grid(row=idx, column=1, padx=5, pady=5)
            entries[text] = widget

        self.filter_edit_items_by_tipo(entries)
        entries['item_combobox'].set(licitacao[4])

        def salvar_edicao():
            numero_processo = entries["Número do Processo:"].get()
            nome_empresa = entries["Nome da Empresa:"].get()
            dados_empresa = entries["Dados da Empresa:"].get()
            tipo_produto = entries["Tipo de Produto:"].get()
            item_solicitado = entries["Item Solicitado:"].get()
            saldo_ata = entries["Saldo da Ata:"].get()
            vencimento_ata = entries["Vencimento da Ata:"].get_date()
            status_licitacao = entries["Status da Licitação:"].get()

            if numero_processo and nome_empresa and tipo_produto and item_solicitado and saldo_ata and status_licitacao:
                if vencimento_ata < datetime.now().date():
                    messagebox.showerror("Erro", "O vencimento da ata não pode ser anterior à data atual.")
                    return
                try:
                    saldo_ata_int = int(saldo_ata)
                    if saldo_ata_int <= 0:
                        messagebox.showerror("Erro", "Saldo da Ata deve ser um número inteiro positivo.")
                        return
                    vencimento_ata_str = vencimento_ata.strftime('%d/%m/%Y')
                    self.cursor.execute(
                        '''
                        UPDATE licitacoes
                        SET numero_processo=?, nome_empresa=?, dados_empresa=?, tipo_produto=?,
                            item_solicitado=?, saldo_ata=?, vencimento_ata=?, status_licitacao=?
                        WHERE id=?
                        ''',
                        (
                            numero_processo, nome_empresa, dados_empresa, tipo_produto,
                            item_solicitado, saldo_ata_int, vencimento_ata_str,
                            status_licitacao, licitacao_id
                        )
                    )
                    self.conn.commit()
                    self.carregar_licitacoes()
                    self.carregar_dashboard()
                    edit_window.destroy()
                except ValueError:
                    messagebox.showerror("Erro", "Saldo da Ata deve ser um número inteiro.")
            else:
                messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios.")

        ctk.CTkButton(edit_window, text="Salvar Alterações", command=salvar_edicao).grid(row=8, column=0, columnspan=2, pady=10)

    def filter_edit_items_by_tipo(self, entries):
        tipo_produto = entries["Tipo de Produto:"].get()
        filtered_items = [
            item_name for item_name, item_info in self.items_dict.items()
            if item_info.get('tipo_produto') == tipo_produto
        ]
        filtered_items.sort()
        item_combobox = entries['item_combobox']
        item_combobox['values'] = filtered_items

    def excluir_licitacao(self):
        selected_item = self.tree_licitacoes.selection()
        if selected_item:
            item = self.tree_licitacoes.item(selected_item)
            licitacao_id = item['values'][0]
            resposta = messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir esta licitação?")
            if resposta:
                self.cursor.execute('DELETE FROM licitacoes WHERE id=?', (licitacao_id,))
                self.conn.commit()
                self.carregar_licitacoes()
                self.carregar_dashboard()
        else:
            messagebox.showwarning("Aviso", "Selecione uma licitação para excluir.")

    def exportar_licitacoes(self):
        try:
            self.cursor.execute('SELECT * FROM licitacoes')
            data = self.cursor.fetchall()
            columns = [description[0] for description in self.cursor.description]
            df = pd.DataFrame(data, columns=columns)
            df.to_excel('licitacoes.xlsx', index=False)
            messagebox.showinfo("Sucesso", "Licitações exportadas para 'licitacoes.xlsx'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar as licitações:\n{e}")

    # ---------------------------- Inserção de Estoque Atualizado Panel ---------------------------- #
    def create_painel_inserir_estoque(self):
        self.painel_insercao_rm.grid_columnconfigure(0, weight=1)
        self.painel_insercao_rm.grid_rowconfigure(0, weight=1)

        import_frame = ctk.CTkFrame(self.painel_insercao_rm)
        import_frame.pack(pady=20)

        ctk.CTkLabel(import_frame, text="Importar Planilha de Estoque Atualizado", font=("Arial", 20)).pack(pady=10)
        ctk.CTkButton(import_frame, text="Importar Planilha", command=self.importar_planilha_estoque).pack(pady=10)
        ctk.CTkButton(import_frame, text="Ver Planilha", command=self.visualizar_planilha_estoque).pack(pady=10)

    def importar_planilha_estoque(self):
        file_path = filedialog.askopenfilename(
            title="Selecione a planilha de Estoque Atualizado",
            filetypes=(
                ("Arquivos Excel", "*.xlsx;*.xls;*.xlsm"),
                ("Arquivos CSV", "*.csv"),
                ("Todos os arquivos", "*.*")
            )
        )
        if file_path:
            try:
                if file_path.endswith(('.xlsx', '.xls', '.xlsm')):
                    df = pd.read_excel(file_path)
                elif file_path.endswith('.csv'):
                    try:
                        df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8')
                    except UnicodeDecodeError:
                        try:
                            df = pd.read_csv(file_path, sep=None, engine='python', encoding='latin1')
                        except UnicodeDecodeError:
                            df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8', errors='replace')
                else:
                    messagebox.showerror("Erro", "Tipo de arquivo não suportado.")
                    return

                df.columns = [col.strip().upper().replace(" ", "_") for col in df.columns]

                item_columns = ["ITEM", "NOME_ITEM", "ITEM_NAME", "ITEM_SOLICITADO", "PI"]
                qtde_disponivel_columns = ["QTDE_DISPONIVEL", "QUANTIDADE_DISPONIVEL", "SALDO", "ESTOQUE", "QUANTIDADE"]

                item_column = next((col for col in item_columns if col in df.columns), None)
                qtde_disponivel_column = next((col for col in qtde_disponivel_columns if col in df.columns), None)

                if not item_column or not qtde_disponivel_column:
                    messagebox.showerror(
                        "Erro",
                        "As colunas necessárias não foram encontradas na planilha.\n"
                        "Por favor, verifique se a planilha contém as colunas de Item e Quantidade Disponível."
                    )
                    return

                df[item_column] = df[item_column].astype(str).str.strip()
                df_agrupado = df.groupby(item_column, as_index=False)[qtde_disponivel_column].sum()
                df_agrupado.to_excel('converted.xlsx', index=False)

                ultima_atualizacao = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                for index, row in df_agrupado.iterrows():
                    nome_item = str(row[item_column]).strip()
                    qtde_disponivel = row[qtde_disponivel_column]
                    try:
                        qtde_disponivel = float(qtde_disponivel)
                    except (ValueError, TypeError):
                        qtde_disponivel = 0

                    self.cursor.execute(
                        '''
                        UPDATE licitacoes
                        SET estoque_disponivel=?, ultima_atualizacao=?
                        WHERE item_solicitado=?
                        ''',
                        (qtde_disponivel, ultima_atualizacao, nome_item)
                    )

                self.conn.commit()
                self.carregar_dashboard()
                messagebox.showinfo(
                    "Sucesso",
                    "Planilha importada e convertida com sucesso. Estoque atualizado.\n"
                    "Arquivo 'converted.xlsx' gerado (itens somados, se duplicados)."
                )
            except Exception as e:
                messagebox.showerror(
                    "Erro",
                    f"Ocorreu um erro ao importar a planilha:\n{e}"
                )

    def visualizar_planilha_estoque(self):
        try:
            self.cursor.execute('SELECT * FROM licitacoes')
            data = self.cursor.fetchall()
            columns = [description[0] for description in self.cursor.description]
            df = pd.DataFrame(data, columns=columns)

            view_window = ctk.CTkToplevel(self.root)
            view_window.title("Visualizar Planilha de Estoque")
            view_window.geometry("800x600")

            frame = ctk.CTkFrame(view_window)
            frame.pack(expand=True, fill='both')

            tree = ttk.Treeview(frame, columns=columns, show='headings', selectmode='browse')
            tree.pack(side='left', fill='both', expand=True)

            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, anchor='center')

            for index, row in df.iterrows():
                tree.insert('', 'end', values=list(row))

            scrollbar = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side='right', fill='y')

            tree.bind('<Double-1>', lambda event: self.edit_licitacao_from_tree(tree, event))

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao visualizar a planilha:\n{e}")

    def edit_licitacao_from_tree(self, tree, event):
        selected_item = tree.selection()
        if selected_item:
            item = tree.item(selected_item)
            licitacao_id = item['values'][0]
            self.open_edit_licitacao_window(licitacao_id)
        else:
            messagebox.showwarning("Aviso", "Selecione uma licitação para editar.")

    # ---------------------------- OCs Panel ---------------------------- #
    def create_painel_ocs(self):
        self.painel_ocs.grid_columnconfigure(0, weight=1)
        self.painel_ocs.grid_columnconfigure(1, weight=1)

        search_frame = ctk.CTkFrame(self.painel_ocs)
        search_frame.grid(row=0, column=0, columnspan=2, pady=5)
        ctk.CTkLabel(search_frame, text="Buscar:").pack(side='left', padx=5)
        self.entry_search_ocs = ctk.CTkEntry(search_frame)
        self.entry_search_ocs.pack(side='left', padx=5)
        ctk.CTkButton(search_frame, text="Buscar", command=self.search_ocs).pack(side='left', padx=5)
        ctk.CTkButton(search_frame, text="Limpar", command=self.clear_search_ocs).pack(side='left', padx=5)

        labels = ["Número da OC:", "Data do Aceite:", "Empresa/Marca:", "Item (PI):",
                  "Quantidade da OC:", "Status do Pedido:"]
        self.entries_ocs = {}

        for idx, text in enumerate(labels):
            ctk.CTkLabel(self.painel_ocs, text=text).grid(row=idx+1, column=0, padx=5, pady=5, sticky='e')
            if text == "Item (PI):":
                item_frame = ctk.CTkFrame(self.painel_ocs)
                item_frame.grid(row=idx+1, column=1, padx=5, pady=5, sticky='nsew')
                item_frame.grid_columnconfigure(0, weight=1)
                item_frame.grid_rowconfigure(1, weight=1)

                self.oc_item_search_entry = ctk.CTkEntry(item_frame)
                self.oc_item_search_entry.grid(row=0, column=0, sticky='ew')
                self.oc_item_search_entry.bind('<KeyRelease>', self.update_oc_item_treeview)

                self.oc_item_treeview = ttk.Treeview(item_frame, height=5)
                self.oc_item_treeview['columns'] = ('Item',)
                self.oc_item_treeview.column('#0', width=0, stretch=tk.NO)
                self.oc_item_treeview.column('Item', anchor='w', width=200)
                self.oc_item_treeview.heading('Item', text='Item', anchor='w')
                self.oc_item_treeview.grid(row=1, column=0, sticky='nsew')
                self.oc_item_treeview.bind('<<TreeviewSelect>>', self.on_oc_item_select)

                self.update_oc_item_treeview()
            elif text == "Quantidade da OC:":
                widget = ctk.CTkEntry(self.painel_ocs)
                widget.configure(validate='key')
                widget.configure(validatecommand=(self.root.register(self.validate_int), '%P'))
                widget.grid(row=idx+1, column=1, padx=5, pady=5, sticky='w')
                self.entries_ocs[text] = widget
            elif text == "Data do Aceite:":
                widget = DateEntry(self.painel_ocs, date_pattern='dd/MM/yyyy')
                widget.grid(row=idx+1, column=1, padx=5, pady=5, sticky='w')
                self.entries_ocs[text] = widget
            elif text == "Status do Pedido:":
                widget = ctk.CTkComboBox(self.painel_ocs, values=["COMRJ licitação", "A caminho", "Recebido Parcialmente", "Recebido"])
                widget.grid(row=idx+1, column=1, padx=5, pady=5, sticky='w')
                self.entries_ocs[text] = widget
            else:
                widget = ctk.CTkEntry(self.painel_ocs)
                widget.grid(row=idx+1, column=1, padx=5, pady=5, sticky='w')
                self.entries_ocs[text] = widget

        self.oc_selected_item = None

        ctk.CTkButton(
            self.painel_ocs,
            text="Adicionar OC",
            command=self.adicionar_oc
        ).grid(row=7, column=0, columnspan=2, padx=5, pady=5)

        columns = (
            'ID', 'Número da OC', 'Item (PI)', 'Quantidade', 'Quantidade Recebida',
            'Quantidade Pendente', 'Status do Pedido', 'Status da Arrecadacao'
        )
        self.tree_ocs, scrollbar = self.create_treeview(self.painel_ocs, columns)
        self.tree_ocs.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky='nsew')
        scrollbar.grid(row=8, column=2, sticky='ns')
        self.painel_ocs.grid_rowconfigure(8, weight=1)

        btn_frame = ctk.CTkFrame(self.painel_ocs)
        btn_frame.grid(row=9, column=0, columnspan=2, pady=5)
        ctk.CTkButton(btn_frame, text="Editar", command=self.editar_oc).pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="Excluir", command=self.excluir_oc).pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="Exportar", command=self.exportar_ocs).pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="Registrar Recebimento", command=self.registrar_recebimento_oc).pack(side='left', padx=5)

    def update_oc_item_treeview(self, event=None):
        search_term = self.oc_item_search_entry.get().lower()
        matching_items = [item for item in self.item_list if search_term in item.lower()]
        for i in self.oc_item_treeview.get_children():
            self.oc_item_treeview.delete(i)
        for item in matching_items:
            self.oc_item_treeview.insert('', 'end', values=(item,))

    def on_oc_item_select(self, event):
        selected_item = self.oc_item_treeview.selection()
        if selected_item:
            item = self.oc_item_treeview.item(selected_item)['values'][0]
            self.oc_selected_item = item
        else:
            self.oc_selected_item = None

    def search_ocs(self):
        query = self.entry_search_ocs.get()
        self.tree_ocs.delete(*self.tree_ocs.get_children())
        self.cursor.execute(
            '''
            SELECT id, numero_oc, item_pi, quantidade_oc, quantidade_recebida, quantidade_pendente,
                   status_pedido, status_arrecadacao
            FROM ocs
            WHERE numero_oc LIKE ? OR item_pi LIKE ?
            ''',
            (f'%{query}%', f'%{query}%')
        )
        for row in self.cursor.fetchall():
            self.tree_ocs.insert('', 'end', values=row)

    def clear_search_ocs(self):
        self.entry_search_ocs.delete(0, 'end')
        self.carregar_ocs()

    def adicionar_oc(self):
        entries = self.entries_ocs
        numero_oc = entries["Número da OC:"].get()
        data_aceite = entries["Data do Aceite:"].get_date()
        empresa_marca = entries["Empresa/Marca:"].get()
        item_pi = self.oc_selected_item
        quantidade_oc = entries["Quantidade da OC:"].get()
        status_pedido = entries["Status do Pedido:"].get()

        if numero_oc and item_pi and quantidade_oc and status_pedido:
            try:
                quantidade_oc_int = int(quantidade_oc)
                if quantidade_oc_int <= 0:
                    messagebox.showerror("Erro", "Quantidade da OC deve ser um número inteiro positivo.")
                    return

                self.cursor.execute(
                    'SELECT saldo_ata FROM licitacoes WHERE item_solicitado=? AND status_licitacao="ASSINADO"',
                    (item_pi,)
                )
                result = self.cursor.fetchone()
                if result:
                    saldo_ata = result[0]
                else:
                    messagebox.showerror("Erro", "Não foi encontrada uma licitação assinada para este item.")
                    return

                self.cursor.execute('SELECT SUM(quantidade_pendente) FROM ocs WHERE item_pi=?', (item_pi,))
                total_pending_result = self.cursor.fetchone()
                total_pending = total_pending_result[0] if total_pending_result[0] else 0

                new_total_pending = total_pending + quantidade_oc_int

                if new_total_pending > saldo_ata:
                    messagebox.showerror(
                        "Erro",
                        f"A quantidade total pendente ({new_total_pending} kg) excede o saldo da ata disponível ({saldo_ata} kg)."
                    )
                    return

                quantidade_recebida = 0
                quantidade_pendente = quantidade_oc_int
                data_aceite_str = data_aceite.strftime('%d/%m/%Y')
                self.cursor.execute(
                    '''
                    INSERT INTO ocs (
                        numero_oc, data_aceite, empresa_marca, item_pi, quantidade_oc,
                        quantidade_recebida, quantidade_pendente, status_pedido
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (
                        numero_oc, data_aceite_str, empresa_marca, item_pi,
                        quantidade_oc_int, quantidade_recebida, quantidade_pendente,
                        status_pedido
                    )
                )
                self.conn.commit()

                for entry in entries.values():
                    if isinstance(entry, DateEntry):
                        entry.set_date(datetime.now())
                    elif isinstance(entry, ctk.CTkComboBox):
                        entry.set('')
                    else:
                        entry.delete(0, 'end')
                self.oc_item_search_entry.delete(0, 'end')
                self.update_oc_item_treeview()
                self.oc_item_treeview.selection_remove(self.oc_item_treeview.selection())
                self.oc_selected_item = None
                self.carregar_ocs()
                self.carregar_dashboard()
                if status_pedido in ["Recebido", "Recebido Parcialmente"]:
                    self.open_arrecadacao_window(self.cursor.lastrowid, status_pedido)

                messagebox.showinfo("Sucesso", "Ordem de Compra adicionada com sucesso.", parent=self.root)
            except ValueError:
                messagebox.showerror("Erro", "Quantidade da OC deve ser um número inteiro.")
        else:
            messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios.")

    def carregar_ocs(self):
        self.tree_ocs.delete(*self.tree_ocs.get_children())
        self.cursor.execute(
            '''
            SELECT id, numero_oc, item_pi, quantidade_oc, quantidade_recebida,
                   quantidade_pendente, status_pedido, status_arrecadacao
            FROM ocs
            '''
        )
        for row in self.cursor.fetchall():
            self.tree_ocs.insert('', 'end', values=row)

    def editar_oc(self):
        selected_item = self.tree_ocs.selection()
        if selected_item:
            item = self.tree_ocs.item(selected_item)
            oc_id = item['values'][0]
            self.open_edit_oc_window(oc_id)
        else:
            messagebox.showwarning("Aviso", "Selecione uma OC para editar.")

    def open_edit_oc_window(self, oc_id):
        edit_window = ctk.CTkToplevel(self.root)
        edit_window.title("Editar OC")
        edit_window.grab_set()

        self.cursor.execute(
            '''
            SELECT numero_oc, data_aceite, empresa_marca, item_pi, quantidade_oc,
                   quantidade_recebida, quantidade_pendente, status_pedido
            FROM ocs
            WHERE id=?
            ''',
            (oc_id,)
        )
        oc = self.cursor.fetchone()

        labels = [
            "Número da OC:", "Data do Aceite:", "Empresa/Marca:", "Item (PI):",
            "Quantidade da OC:", "Status do Pedido:"
        ]
        entries = {}

        for idx, text in enumerate(labels):
            ctk.CTkLabel(edit_window, text=text).grid(row=idx, column=0, padx=5, pady=5)
            if text == "Item (PI):":
                widget = ttk.Combobox(edit_window, values=self.item_list, width=50)
            elif text == "Quantidade da OC:":
                widget = ctk.CTkEntry(edit_window)
                widget.configure(validate='key')
                widget.configure(validatecommand=(self.root.register(self.validate_int), '%P'))
            elif text == "Data do Aceite:":
                widget = DateEntry(edit_window, date_pattern='dd/MM/yyyy')
            elif text == "Status do Pedido:":
                widget = ctk.CTkComboBox(edit_window, values=["COMRJ licitação", "A caminho", "Recebido Parcialmente", "Recebido"])
            else:
                widget = ctk.CTkEntry(edit_window)

            if isinstance(widget, (ctk.CTkEntry, ttk.Entry)):
                widget.insert(0, oc[idx])
            elif isinstance(widget, ctk.CTkComboBox):
                widget.set(oc[idx])
            elif isinstance(widget, DateEntry):
                date_value = datetime.strptime(oc[idx], '%d/%m/%Y')
                widget.set_date(date_value)
            widget.grid(row=idx, column=1, padx=5, pady=5)
            entries[text] = widget

        def salvar_edicao():
            numero_oc = entries["Número da OC:"].get()
            data_aceite = entries["Data do Aceite:"].get_date()
            empresa_marca = entries["Empresa/Marca:"].get()
            item_pi = entries["Item (PI):"].get()
            quantidade_oc = entries["Quantidade da OC:"].get()
            status_pedido = entries["Status do Pedido:"].get()

            if numero_oc and item_pi and quantidade_oc and status_pedido:
                try:
                    quantidade_oc_int = int(quantidade_oc)
                    if quantidade_oc_int <= 0:
                        messagebox.showerror("Erro", "Quantidade da OC deve ser um número inteiro positivo.")
                        return
                    if quantidade_oc_int < oc[5]:
                        messagebox.showerror("Erro", "Quantidade da OC não pode ser menor que a quantidade já recebida.")
                        return
                    data_aceite_str = data_aceite.strftime('%d/%m/%Y')
                    quantidade_pendente = quantidade_oc_int - oc[5]
                    if quantidade_pendente < 0:
                        quantidade_pendente = 0
                    self.cursor.execute(
                        '''
                        UPDATE ocs
                        SET numero_oc=?, data_aceite=?, empresa_marca=?, item_pi=?,
                            quantidade_oc=?, quantidade_pendente=?, status_pedido=?
                        WHERE id=?
                        ''',
                        (
                            numero_oc, data_aceite_str, empresa_marca, item_pi,
                            quantidade_oc_int, quantidade_pendente, status_pedido, oc_id
                        )
                    )
                    self.conn.commit()
                    self.carregar_ocs()
                    self.carregar_dashboard()
                    edit_window.destroy()
                    if status_pedido in ["Recebido", "Recebido Parcialmente"]:
                        self.open_arrecadacao_window(oc_id, status_pedido)
                except ValueError:
                    messagebox.showerror("Erro", "Quantidade da OC deve ser um número inteiro.")
            else:
                messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios.")

        ctk.CTkButton(edit_window, text="Salvar Alterações", command=salvar_edicao).grid(row=6, column=0, columnspan=2, pady=10)

    def excluir_oc(self):
        selected_item = self.tree_ocs.selection()
        if selected_item:
            item = self.tree_ocs.item(selected_item)
            oc_id = item['values'][0]
            resposta = messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir esta OC?")
            if resposta:
                self.cursor.execute('SELECT item_pi, quantidade_oc FROM ocs WHERE id=?', (oc_id,))
                oc_data = self.cursor.fetchone()
                if oc_data:
                    item_pi, quantidade_oc = oc_data
                    quantidade_oc = int(quantidade_oc)
                    self.cursor.execute('SELECT id, saldo_ata FROM licitacoes WHERE item_solicitado=?', (item_pi,))
                    licitacao_data = self.cursor.fetchone()
                    if licitacao_data:
                        licitacao_id, saldo_ata = licitacao_data
                        novo_saldo_ata = saldo_ata + quantidade_oc
                        self.cursor.execute(
                            'UPDATE licitacoes SET saldo_ata=? WHERE id=?',
                            (novo_saldo_ata, licitacao_id)
                        )
                        self.conn.commit()
                self.cursor.execute('DELETE FROM ocs WHERE id=?', (oc_id,))
                self.conn.commit()
                self.carregar_ocs()
                self.carregar_licitacoes()
                self.carregar_dashboard()
        else:
            messagebox.showwarning("Aviso", "Selecione uma OC para excluir.")

    def exportar_ocs(self):
        try:
            self.cursor.execute('SELECT * FROM ocs')
            data = self.cursor.fetchall()
            columns = [description[0] for description in self.cursor.description]
            df = pd.DataFrame(data, columns=columns)
            df.to_excel('ocs.xlsx', index=False)
            messagebox.showinfo("Sucesso", "OCs exportadas para 'ocs.xlsx'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar as OCs:\n{e}")

    def registrar_recebimento_oc(self):
        selected_item = self.tree_ocs.selection()
        if selected_item:
            item = self.tree_ocs.item(selected_item)
            oc_id = item['values'][0]
            self.cursor.execute('SELECT status_pedido FROM ocs WHERE id=?', (oc_id,))
            status_pedido = self.cursor.fetchone()[0]
            if status_pedido in ["A caminho", "Recebido Parcialmente", "Recebido"]:
                self.open_arrecadacao_window(oc_id, status_pedido)
            else:
                messagebox.showwarning(
                    "Aviso",
                    "Somente OCs com status 'A caminho', 'Recebido Parcialmente' ou 'Recebido' podem registrar recebimento."
                )
        else:
            messagebox.showwarning("Aviso", "Selecione uma OC para registrar recebimento.")

    def open_edit_cmm_entry_window(self, item_name):
        edit_window = ctk.CTkToplevel(self.root)
        edit_window.title(f"Editar CMM - {item_name}")
        edit_window.grab_set()

        ctk.CTkLabel(edit_window, text="Consumo Médio Mensal (CMM):").grid(row=0, column=0, padx=5, pady=5)
        entry_cmm = ctk.CTkEntry(edit_window)
        entry_cmm.grid(row=0, column=1, padx=5, pady=5)

        current_cmm = self.items_dict[item_name].get('cmm', 0)
        entry_cmm.insert(0, str(current_cmm))

        def save_cmm():
            new_cmm = entry_cmm.get()
            try:
                new_cmm_value = float(new_cmm)
                if new_cmm_value < 0:
                    messagebox.showerror("Erro", "O CMM não pode ser negativo.")
                    return
                self.cursor.execute('UPDATE items SET cmm=? WHERE item_name=?', (new_cmm_value, item_name))
                self.conn.commit()
                self.load_items_from_database()
                self.carregar_dashboard()
                edit_window.destroy()
                messagebox.showinfo("Sucesso", f"CMM atualizado para o item '{item_name}'.")
            except ValueError:
                messagebox.showerror("Erro", "Por favor, insira um valor numérico válido para o CMM.")

        salvar_button = ctk.CTkButton(edit_window, text="Salvar", command=save_cmm)
        salvar_button.grid(row=1, column=0, columnspan=2, pady=10)

    def open_arrecadacao_window(self, oc_id, status_pedido):
        arrecadacao_window = ctk.CTkToplevel(self.root)
        arrecadacao_window.title("Status da Arrecadação")
        arrecadacao_window.grab_set()

        ctk.CTkLabel(arrecadacao_window, text="Status da Arrecadação:").grid(row=0, column=0, padx=5, pady=5)
        combo_status_arrecadacao = ctk.CTkComboBox(
            arrecadacao_window,
            values=["1 - Perícia", "2 - Em contagem", "3 - Em estocagem", "4 - Disponível no sistema"]
        )
        combo_status_arrecadacao.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(arrecadacao_window, text="Quantidade Recebida:").grid(row=1, column=0, padx=5, pady=5)
        entry_quantidade_recebida = ctk.CTkEntry(arrecadacao_window)
        entry_quantidade_recebida.configure(validate='key')
        entry_quantidade_recebida.configure(validatecommand=(self.root.register(self.validate_int), '%P'))
        entry_quantidade_recebida.grid(row=1, column=1, padx=5, pady=5)

        def salvar_status_arrecadacao():
            status_arrecadacao = combo_status_arrecadacao.get()
            quantidade_recebida = entry_quantidade_recebida.get()
            if not quantidade_recebida or not status_arrecadacao:
                messagebox.showwarning("Aviso", "Informe todos os campos.")
                return
            quantidade_recebida = int(quantidade_recebida)

            self.cursor.execute(
                'SELECT item_pi, quantidade_oc, quantidade_recebida FROM ocs WHERE id=?',
                (oc_id,)
            )
            oc_data = self.cursor.fetchone()
            if not oc_data:
                messagebox.showerror("Erro", "OC não encontrada.")
                return
            item_pi, quantidade_oc, quantidade_recebida_atual = oc_data
            quantidade_recebida_total = quantidade_recebida_atual + quantidade_recebida

            # For both "1 - Perícia" and "4 - Disponível no sistema", revert to old behavior:
            # set the new status as "Disponível no sistema" (and subtract the received quantity)
            if status_arrecadacao in ["1 - Perícia", "4 - Disponível no sistema"]:
                if quantidade_recebida_total > quantidade_oc:
                    quantidade_recebida_total = quantidade_oc
                quantidade_pendente = quantidade_oc - quantidade_recebida_total
                status_pedido_novo = "Disponível no sistema"
            else:
                if quantidade_recebida_total > quantidade_oc:
                    messagebox.showerror("Erro", "Quantidade recebida total não pode exceder a quantidade da OC.")
                    return
                quantidade_pendente = quantidade_oc - quantidade_recebida_total
                if quantidade_pendente == 0:
                    status_pedido_novo = "Recebido"
                else:
                    status_pedido_novo = "Recebido Parcialmente"

            self.cursor.execute(
                '''
                UPDATE ocs
                SET quantidade_recebida=?, quantidade_pendente=?, status_pedido=?, status_arrecadacao=?
                WHERE id=?
                ''',
                (
                    quantidade_recebida_total, quantidade_pendente,
                    status_pedido_novo, status_arrecadacao, oc_id
                )
            )

            self.cursor.execute('SELECT id, saldo_ata FROM licitacoes WHERE item_solicitado=?', (item_pi,))
            licitacao = self.cursor.fetchone()
            if licitacao:
                licitacao_id, saldo_ata = licitacao
                # Only subtract from licitação if not "1 - Perícia" (but now old behavior subtracts in either case)
                novo_saldo_ata = saldo_ata - quantidade_recebida
                if novo_saldo_ata < 0:
                    novo_saldo_ata = 0
                self.cursor.execute('UPDATE licitacoes SET saldo_ata=? WHERE id=?', (novo_saldo_ata, licitacao_id))
            else:
                messagebox.showerror("Erro", "Licitação correspondente não encontrada.")
                return

            self.conn.commit()
            self.carregar_ocs()
            self.carregar_licitacoes()
            self.carregar_dashboard()
            messagebox.showinfo("Sucesso", "Status da arrecadação salvo com sucesso.")
            arrecadacao_window.destroy()

        ctk.CTkButton(arrecadacao_window, text="Salvar Status", command=salvar_status_arrecadacao).grid(row=2, column=0, columnspan=2, pady=10)

    # ---------------------------- Dashboard Panel ---------------------------- #
    def create_painel_dashboard(self):
        dashboard_frame = ctk.CTkFrame(self.painel_dashboard)
        dashboard_frame.pack(fill='both', expand=True)

        tab_control = ctk.CTkTabview(dashboard_frame)
        tab_control.pack(expand=1, fill='both')

        self.tab_frigorificados = tab_control.add("FRIGORIFICADOS")
        self.tab_secos = tab_control.add("SECOS")
        self.tab_todos = tab_control.add("TODOS")

        # Only "Tabela" and "Cards" tabs are added (Gráficos removed)
        self.create_dashboard_tabs(self.tab_frigorificados, "FRIGORIFICADOS")
        self.create_dashboard_tabs(self.tab_secos, "SECOS")
        self.create_dashboard_tabs(self.tab_todos, "TODOS")

    def create_dashboard_tabs(self, parent_tab, tipo_produto):
        sub_tab_control = ctk.CTkTabview(parent_tab)
        sub_tab_control.pack(expand=1, fill='both')

        tab_tabela = sub_tab_control.add("Tabela")
        tab_cards = sub_tab_control.add("Cards")
        self.create_dashboard_tabela(tab_tabela, tipo_produto)
        self.create_dashboard_cards(tab_cards, tipo_produto)

    def create_dashboard_tabela(self, tab, tipo_produto):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(tab, text=f"Dashboard de Produtos - {tipo_produto}", font=("Calibri", 44)).grid(row=0, column=0, pady=10)

        columns = ('Item', 'Saldo em Ata', 'Vencimento', 'Disp. p/lib.', 'OC Pendente')
        tree_dashboard, scrollbar = self.create_treeview(tab, columns)
        tree_dashboard.grid(row=1, column=0, padx=5, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=1, sticky='ns')
        tab.grid_rowconfigure(1, weight=1)

        ctk.CTkButton(tab, text="Exportar", command=lambda: self.exportar_dashboard(tipo_produto)).grid(row=2, column=0, pady=10)

        if not hasattr(self, 'tree_dashboards'):
            self.tree_dashboards = {}
        self.tree_dashboards[tipo_produto] = tree_dashboard

    def create_dashboard_cards(self, tab, tipo_produto):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        canvas_cards = ctk.CTkCanvas(tab)
        canvas_cards.grid(row=0, column=0, sticky='nsew')
        # Set initial view to top-left
        canvas_cards.xview_moveto(0)
        canvas_cards.yview_moveto(0)

        # Bind mouse wheel for vertical scrolling (Windows/Mac)
        canvas_cards.bind("<MouseWheel>", lambda event: canvas_cards.yview_scroll(-1 * (event.delta // 120), "units"))
        # Also for Linux (scroll up and down)
        canvas_cards.bind("<Button-4>", lambda event: canvas_cards.yview_scroll(-1, "units"))
        canvas_cards.bind("<Button-5>", lambda event: canvas_cards.yview_scroll(1, "units"))

        # Vertical scrollbar
        scrollbar_cards = ctk.CTkScrollbar(tab, orientation='vertical', command=canvas_cards.yview)
        scrollbar_cards.grid(row=0, column=1, sticky='ns')
        canvas_cards.configure(yscrollcommand=scrollbar_cards.set)

        # Horizontal scrollbar
        scrollbar_h = ctk.CTkScrollbar(tab, orientation='horizontal', command=canvas_cards.xview)
        scrollbar_h.grid(row=1, column=0, sticky='ew')
        canvas_cards.configure(xscrollcommand=scrollbar_h.set)
        tab.grid_rowconfigure(1, weight=0)

        frame_cards = ctk.CTkFrame(canvas_cards)
        canvas_cards.create_window((0, 0), window=frame_cards, anchor='nw')
        frame_cards.bind('<Configure>', lambda e: canvas_cards.configure(scrollregion=canvas_cards.bbox('all')))

        if not hasattr(self, 'frame_cards_dict'):
            self.frame_cards_dict = {}
        self.frame_cards_dict[tipo_produto] = (frame_cards, canvas_cards)

    def carregar_dashboard(self):
        self.load_dashboard_data()
        self.update_dashboard_filter()

    def load_dashboard_data(self):
        self.dashboard_data = []
        self.cursor.execute(
            '''
            SELECT id, item_solicitado, tipo_produto, saldo_ata_inicial, vencimento_ata, estoque_disponivel, disp_manual
            FROM licitacoes
            WHERE status_licitacao="ASSINADO"
            '''
        )
        licitacoes = self.cursor.fetchall()
        for lic in licitacoes:
            licitacao_id = lic[0]
            item = lic[1]
            tipo_produto = lic[2]
            saldo_ata_inicial = lic[3]
            vencimento = lic[4]
            estoque = lic[5]
            disp_manual = lic[6]
            self.cursor.execute('SELECT SUM(quantidade_oc) FROM ocs WHERE item_pi=? AND status_pedido="Recebido"', (item,))
            sum_oc = self.cursor.fetchone()[0] or 0
            saldo_restante = saldo_ata_inicial - sum_oc
            self.cursor.execute(
                '''
                SELECT SUM(quantidade_pendente)
                FROM ocs
                WHERE item_pi=? AND status_pedido!="Recebido"
                ''',
                (item,)
            )
            pendente = self.cursor.fetchone()[0] or 0
            consumido_percentual = ((saldo_ata_inicial - saldo_restante) / saldo_ata_inicial * 100) if saldo_ata_inicial else 0
            uf = self.items_dict.get(item, {}).get('uf', '')
            self.dashboard_data.append((licitacao_id, item, tipo_produto, saldo_ata_inicial, saldo_restante, vencimento, estoque, disp_manual, pendente, consumido_percentual, uf))

    def update_dashboard_filter(self):
        tipos_produto = ["FRIGORIFICADOS", "SECOS", "TODOS"]
        for tipo in tipos_produto:
            if tipo == "TODOS":
                filtered_data = self.dashboard_data
            else:
                filtered_data = [d for d in self.dashboard_data if d[2] == tipo]
            # Sort the cards alphabetically by item name
            sorted_data = sorted(filtered_data, key=lambda x: x[1])
            self.update_dashboard_tabela(sorted_data, tipo)
            self.update_dashboard_cards(sorted_data, tipo)

    def update_dashboard_tabela(self, data, tipo_produto):
        tree_dashboard = self.tree_dashboards.get(tipo_produto)
        if tree_dashboard:
            tree_dashboard.delete(*tree_dashboard.get_children())
            for values in data:
                # values: (licitacao_id, item, tipo, saldo_ata_inicial, saldo_restante, vencimento, estoque, disp_manual, pendente, consumido_percentual, uf)
                _, item, _, _, saldo_restante, vencimento, estoque, disp_manual, pendente, _, _ = values
                display_disp = disp_manual if disp_manual is not None else estoque
                tree_dashboard.insert('', 'end', values=(item, saldo_restante, vencimento, display_disp, pendente))

    def update_dashboard_cards(self, data, tipo_produto):
        frame_cards, canvas_cards = self.frame_cards_dict.get(tipo_produto, (None, None))
        if frame_cards:
            for widget in frame_cards.winfo_children():
                widget.destroy()
            num_columns = 4
            for idx, values in enumerate(data):
                # values: (licitacao_id, item, tipo_produto_item, saldo_ata_inicial, saldo_restante, vencimento, estoque, disp_manual, pendente, consumido_percentual, uf)
                licitacao_id, item, tipo_produto_item, saldo_ata_inicial, saldo_restante, vencimento, estoque, disp_manual, pendente, consumido_percentual, uf = values
                row = idx // num_columns
                col = idx % num_columns
                card = ctk.CTkFrame(frame_cards, corner_radius=10, fg_color="white", width=250, height=300)
                card.grid_propagate(False)
                card.grid(row=row, column=col, padx=20, pady=20, sticky='nsew')

                header_frame = ctk.CTkFrame(card, fg_color="white")
                header_frame.pack(fill='x', pady=(10, 5))
                ctk.CTkLabel(header_frame, text=item, font=('Arial', 12, 'bold'), wraplength=230).pack(side='top', padx=(10, 5))

                progress_frame = ctk.CTkFrame(card, fg_color="white")
                progress_frame.pack(fill='x', padx=20, pady=(10, 10))
                progress_bar = ttk.Progressbar(progress_frame, length=200, mode='determinate')
                progress_bar['value'] = consumido_percentual
                progress_bar.pack(pady=10)
                if consumido_percentual <= 30:
                    progress_bar.configure(style="green.Horizontal.TProgressbar")
                elif 30 < consumido_percentual <= 69:
                    progress_bar.configure(style="yellow.Horizontal.TProgressbar")
                else:
                    progress_bar.configure(style="red.Horizontal.TProgressbar")
                style = ttk.Style()
                style.configure("green.Horizontal.TProgressbar", foreground='green', background='green')
                style.configure("yellow.Horizontal.TProgressbar", foreground='yellow', background='yellow')
                style.configure("red.Horizontal.TProgressbar", foreground='red', background='red')
                ctk.CTkLabel(
                    progress_frame,
                    text=f"{consumido_percentual:.1f}% do saldo consumido",
                    font=('Arial', 12)
                ).pack(pady=(0, 5))
                ctk.CTkLabel(
                    progress_frame,
                    text=f"Restam {self.format_number_br(saldo_restante)} KG de {self.format_number_br(saldo_ata_inicial)} KG",
                    font=('Arial', 12, 'bold')
                ).pack()

                try:
                    venc_date = datetime.strptime(vencimento, '%d/%m/%Y')
                    if venc_date - datetime.now() <= timedelta(days=90):
                        venc_color = "red"
                    else:
                        venc_color = "black"
                except Exception:
                    venc_color = "black"
                ctk.CTkLabel(
                    card,
                    text=f"Vencimento da ata: {vencimento}",
                    font=('Arial', 10, 'bold'),
                    text_color=venc_color
                ).pack(pady=(5, 10))

                info_frame = ctk.CTkFrame(card, fg_color="white")
                info_frame.pack(fill='x', padx=20, pady=(10, 10))
                ctk.CTkLabel(info_frame, text="🛒 Disp. p/lib.", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w')
                disp_value = disp_manual if disp_manual is not None else estoque
                ctk.CTkLabel(info_frame, text=f"{self.format_number_br(disp_value)} KG", font=('Arial', 10)).grid(row=0, column=1, sticky='e')
                ctk.CTkLabel(info_frame, text="📅 CMM", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=(10, 0))
                ctk.CTkLabel(
                    info_frame,
                    text=f"{self.items_dict.get(item, {}).get('cmm', 0)}",
                    font=('Arial', 10)
                ).grid(row=0, column=3, sticky='e', padx=(5, 0))
                ctk.CTkLabel(info_frame, text="🚚 OC Pendente", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w')
                ctk.CTkLabel(info_frame, text=f"{pendente} KG", font=('Arial', 10)).grid(row=1, column=1, sticky='e')
                cmm_value = self.items_dict.get(item, {}).get('cmm', 0)
                if cmm_value > 0:
                    autonomia_dias = int((disp_value / cmm_value) * 30)
                    meses = autonomia_dias // 30
                    dias = autonomia_dias % 30
                    if meses > 0 and dias > 0:
                        autonomia_str = f"{meses} Mes{'es' if meses > 1 else ''} e {dias} dias"
                    elif meses > 0 and dias == 0:
                        autonomia_str = f"{meses} Mes{'es' if meses > 1 else ''}"
                    else:
                        autonomia_str = f"{dias} dias"
                else:
                    autonomia_str = "Indefinido"
                ctk.CTkLabel(info_frame, text="⏳ Autonomia", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=(10, 0))
                ctk.CTkLabel(info_frame, text=f"{autonomia_str}", font=('Arial', 10)).grid(row=1, column=3, sticky='e', padx=(5, 0))

                button_frame = ctk.CTkFrame(card, fg_color="white")
                button_frame.pack(fill='x', padx=20, pady=(10, 10))
                ctk.CTkButton(
                    button_frame,
                    text="Verificar OCs",
                    command=lambda lid=licitacao_id: self.open_ocs_for_item_by_id(lid),
                    width=180,
                    fg_color="#FFA500"
                ).pack(pady=(5, 0))
                ctk.CTkButton(
                    button_frame,
                    text="Editar Disp. p/lib",
                    command=lambda lid=licitacao_id: self.open_edit_disp_manual_window(lid),
                    width=180,
                    fg_color="#00BFFF"
                ).pack(pady=(5, 0))
                ctk.CTkButton(
                    button_frame,
                    text="Comentários",
                    command=lambda lid=licitacao_id: self.open_comments_window(lid),
                    width=180,
                    fg_color="#32CD32"
                ).pack(pady=(5, 0))
            canvas_cards.update_idletasks()
            canvas_cards.configure(scrollregion=canvas_cards.bbox('all'))

    def open_ocs_for_item_by_id(self, licitacao_id):
        self.cursor.execute('SELECT item_solicitado FROM licitacoes WHERE id=?', (licitacao_id,))
        result = self.cursor.fetchone()
        if result:
            item = result[0]
            self.open_ocs_for_item(item)
        else:
            messagebox.showerror("Erro", "Licitação não encontrada.")

    def open_ocs_for_item(self, item_name):
        oc_window = ctk.CTkToplevel(self.root)
        oc_window.title(f"OCs for {item_name}")
        oc_window.grab_set()

        columns = (
            'ID', 'Número da OC', 'Item (PI)', 'Quantidade',
            'Quantidade Recebida', 'Quantidade Pendente',
            'Status do Pedido', 'Status da Arrecadacao'
        )
        tree, scrollbar = self.create_treeview(oc_window, columns)
        tree.pack(side='left', expand=1, fill='both')
        scrollbar.pack(side='right', fill='y')

        self.cursor.execute(
            '''
            SELECT id, numero_oc, item_pi, quantidade_oc, quantidade_recebida,
                   quantidade_pendente, status_pedido, status_arrecadacao
            FROM ocs
            WHERE item_pi=?
            ''',
            (item_name,)
        )
        for row in self.cursor.fetchall():
            tree.insert('', 'end', values=row)

        tree.bind('<Double-1>', lambda event: self.edit_oc_from_tree(tree, event))

    def edit_oc_from_tree(self, tree, event):
        selected_item = tree.selection()
        if selected_item:
            item = tree.item(selected_item)
            oc_id = item['values'][0]
            self.open_edit_oc_window(oc_id)
        else:
            messagebox.showwarning("Aviso", "Selecione uma OC para editar.")

    def exportar_dashboard(self, tipo_produto):
        try:
            if tipo_produto == "TODOS":
                data = self.dashboard_data
            else:
                data = [row for row in self.dashboard_data if row[2] == tipo_produto]
            df = pd.DataFrame(
                data,
                columns=[
                    'Licitacao_ID', 'Item', 'Tipo de Produto', 'Saldo Ata Inicial', 'Saldo Restante',
                    'Vencimento', 'Estoque Disponível', 'Disp. Manual', 'OC Pendente',
                    'Consumido Percentual', 'UF'
                ]
            )
            df.to_excel(f'dashboard_{tipo_produto}.xlsx', index=False)
            messagebox.showinfo("Sucesso", f"Dashboard exportado para 'dashboard_{tipo_produto}.xlsx'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar o dashboard:\n{e}")

    # ---------------------------- Comments Window (Multiple Comments) ---------------------------- #
    def open_comments_window(self, licitacao_id):
        comments_win = ctk.CTkToplevel(self.root)
        comments_win.title("Comentários")
        comments_win.geometry("500x400")
        comments_win.grab_set()

        main_frame = ctk.CTkFrame(comments_win)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Frame for comment list with scrollbar
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.pack(fill='both', expand=True)

        canvas = ctk.CTkCanvas(list_frame)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar = ctk.CTkScrollbar(list_frame, orientation='vertical', command=canvas.yview)
        scrollbar.pack(side='right', fill='y')
        canvas.configure(yscrollcommand=scrollbar.set)
        inner_frame = ctk.CTkFrame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor='nw')
        inner_frame.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-1*(event.delta//120), "units"))

        # Functions for comment list management
        def refresh_comments():
            for widget in inner_frame.winfo_children():
                widget.destroy()
            self.cursor.execute("SELECT id, comentario, data FROM comentarios WHERE licitacao_id=? ORDER BY id ASC", (licitacao_id,))
            comments = self.cursor.fetchall()
            for cid, comentario, data in comments:
                comment_frame = ctk.CTkFrame(inner_frame)
                comment_frame.pack(fill='x', pady=5, padx=5)
                lbl = ctk.CTkLabel(comment_frame, text=f"{data} - {comentario}", anchor='w', justify='left', wraplength=400)
                lbl.pack(side='left', fill='x', expand=True)
                btn_edit = ctk.CTkButton(comment_frame, text="Editar", width=60,
                                         command=lambda cid=cid, old=comentario: edit_comment(cid, old))
                btn_edit.pack(side='left', padx=2)
                btn_del = ctk.CTkButton(comment_frame, text="Excluir", width=60,
                                        command=lambda cid=cid: delete_comment(cid))
                btn_del.pack(side='left', padx=2)

        def edit_comment(comment_id, old_text):
            edit_win = ctk.CTkToplevel(comments_win)
            edit_win.title("Editar Comentário")
            edit_win.geometry("400x200")
            ctk.CTkLabel(edit_win, text="Editar Comentário:").pack(pady=5)
            text_entry = ctk.CTkTextbox(edit_win, width=350, height=100)
            text_entry.pack(pady=5)
            text_entry.delete("1.0", "end")
            text_entry.insert("1.0", old_text)
            def save_edit():
                new_text = text_entry.get("1.0", "end-1c")
                self.cursor.execute("UPDATE comentarios SET comentario=?, data=? WHERE id=?", 
                                    (new_text, datetime.now().strftime("%d/%m/%Y %H:%M:%S"), comment_id))
                self.conn.commit()
                edit_win.destroy()
                refresh_comments()
            ctk.CTkButton(edit_win, text="Salvar", command=save_edit).pack(pady=5)

        def delete_comment(comment_id):
            if messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir este comentário?"):
                self.cursor.execute("DELETE FROM comentarios WHERE id=?", (comment_id,))
                self.conn.commit()
                refresh_comments()

        # Area to add a new comment
        add_frame = ctk.CTkFrame(main_frame)
        add_frame.pack(fill='x', pady=10)
        ctk.CTkLabel(add_frame, text="Novo Comentário:").pack(side='left', padx=5)
        new_comment = ctk.CTkEntry(add_frame, width=300)
        new_comment.pack(side='left', padx=5)
        def add_comment():
            comment_text = new_comment.get()
            if comment_text.strip() == "":
                messagebox.showwarning("Aviso", "Comentário não pode ser vazio.")
                return
            self.cursor.execute("INSERT INTO comentarios (licitacao_id, comentario, data) VALUES (?, ?, ?)",
                                (licitacao_id, comment_text, datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
            self.conn.commit()
            new_comment.delete(0, 'end')
            refresh_comments()
        ctk.CTkButton(add_frame, text="Adicionar", command=add_comment).pack(side='left', padx=5)

        refresh_comments()

# ---------------------------- Splash Screen and Main ---------------------------- #
def main():
    # Create main window but keep it hidden while splash is active
    root = ctk.CTk()
    root.withdraw()

    # Create a splash screen window
    splash = tk.Toplevel()
    splash.overrideredirect(True)
    splash.geometry("800x600+500+300")  # Adjust as needed

    # Set the path to your splash image here
    splash_image_path = "logo2.png"
    try:
        splash_image = tk.PhotoImage(file=splash_image_path)
        label = tk.Label(splash, image=splash_image)
        label.pack(expand=True, fill='both')
    except Exception as e:
        label = tk.Label(splash, text="Loading...", font=("Arial", 24))
        label.pack(expand=True, fill='both')

    splash.update()

    # After 3 seconds, destroy the splash and show the main window with the app
    root.after(3000, lambda: (splash.destroy(), root.deiconify(), ControleEstoqueApp(root)))
    root.mainloop()

if __name__ == "__main__":
    main()
    
