# -*- coding: utf-8 -*-

# =============================================================================
# GESTOR DE ESTOQUE DE ATIVOS DE TI (HEXA STOCK)
#
# Autor: Wesley Puckar
# Data: 20/09/2025
#
# Descrição:
# Aplicação de desktop desenvolvida com Tkinter para gerenciar o estoque
# de equipamentos de TI. Utiliza o Google Sheets como banco de dados
# para permitir acesso colaborativo e fácil manutenção dos dados.
#
# Funcionalidades Principais:
# - CRUD (Criar, Ler, Atualizar, Deletar) de equipamentos.
# - Registro de movimentações de estoque (entrada, saída, descarte).
# - Registro de movimentações de ativos entre setores.
# - Dashboard com estatísticas rápidas.
# - Geração de relatórios em HTML.
# =============================================================================


# --- Importação das Bibliotecas ---

# Bibliotecas padrão do Python
import datetime
import os
import webbrowser
import sys

# Bibliotecas de terceiros (necessário instalar: pip install gspread pandas oauth2client)
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

# Biblioteca para a interface gráfica (GUI)
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

# --- Constantes Globais ---
# Define o nome da planilha do Google Sheets que o programa irá acessar.
NOME_PLANILHA = "ControleDeEstoqueTI"

class App(tk.Tk):
    
    """ Classe principal da aplicação. Herda de tk.Tk para criar a janela principal.
    Gerencia a interface gráfica, a lógica de negócios e a comunicação com a
    planilha do Google."""
    

    def resource_path(self, relative_path):
        """ Retorna o caminho absoluto para um recurso (ex: arquivo de template).
        Esta função é crucial para que o PyInstaller encontre os arquivos
        quando o programa é empacotado em um executável (.exe). """
        
        try:
            # PyInstaller cria uma pasta temporária e armazena o caminho na variável _MEIPASS
            base_path = sys._MEIPASS
            
        except Exception:
            
            # Se não estiver rodando via PyInstaller, usa o caminho do script
            base_path = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(base_path, relative_path)

    def __init__(self):
        super().__init__()
        
        # --- Configurações Iniciais da Janela ---
        self.title("Hexa Stock - Gestor de Ativos de TI")
        self.geometry("1250x900")
        self.configure(bg="#F4F6F8")
        
                # --- Conexão e Carregamento de Dados ---
        # Tenta conectar com o Google Sheets. Se falhar, a aplicação é encerrada.
        
        if not self.conectar_google_sheets():
            return  # Impede a continuação se a conexão falhar

        # Carrega as configurações iniciais da aba 'config' da planilha.    
        self._load_config()
        
        # --- Criação da Interface Gráfica (Widgets) ---
        # Define um estilo visual para os componentes da interface.
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        
        # Cria o container principal com abas (Notebook).
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Cria os frames (páginas) para cada aba.
        self.estoque_tab = ttk.Frame(self.notebook, padding="10")
        self.mov_setores_tab = ttk.Frame(self.notebook, padding="10")

        # Adiciona as abas ao Notebook.
        self.notebook.add(self.estoque_tab, text='Estoque da Informática')
        self.notebook.add(self.mov_setores_tab, text='Movimentação Entre Setores')

        # Chama os métodos para popular cada aba com seus respectivos widgets.
        self.criar_aba_estoque()
        self.criar_aba_movimentacao_setores()
        
        self.refresh_all_data()

    def criar_aba_estoque(self):
        
        """Cria e organiza todos os widgets da aba 'Estoque da Informática'."""
        
        self.style.configure("TLabel", background="#F4F6F8", font=("Roboto", 11))
        self.style.configure("TEntry", font=("Roboto", 11))
        self.style.configure("TButton", font=("Roboto", 10, "bold"), padding=5)
        self.style.configure("Treeview", font=("Roboto", 10), rowheight=25)
        self.style.configure("Treeview.Heading", font=("Roboto", 10, "bold"))
        self.style.map("TButton", foreground=[('active', 'white'), ('!disabled', 'white')], background=[('active', '#0056b3'), ('!disabled', '#007BFF')])
        self.style.configure("Card.TFrame", relief="solid", borderwidth=1, background="white")
        self.style.configure("CardIcon.TLabel", font=("Roboto", 28), background="white")
        self.style.configure("CardNumber.TLabel", font=("Roboto", 24, "bold"), background="white")
        self.style.configure("CardText.TLabel", font=("Roboto", 10), background="white")
        self.style.configure("Multiline.Treeview", rowheight=65, font=("Roboto", 10))
        self.style.configure("Multiline.Treeview.Heading", font=("Roboto", 10, "bold"))

        # --- Seção Dashboard (Cards) ---
        dashboard_frame = ttk.LabelFrame(self.estoque_tab, text="Dashboard", padding="10"); dashboard_frame.pack(fill='x', padx=10, pady=10)
        self.total_itens_var = tk.StringVar(value="..."); self.tipos_unicos_var = tk.StringVar(value="..."); self.estoque_baixo_var = tk.StringVar(value="..."); self.mov_mes_var = tk.StringVar(value="...")
        
        # Função auxiliar para criar os cards do dashboard de forma reutilizável.
        def create_dashboard_card(parent, icon, number_var, text, icon_color):
            card_frame = ttk.Frame(parent, style="Card.TFrame", padding=10); card_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            icon_label = ttk.Label(card_frame, text=icon, style="CardIcon.TLabel", foreground=icon_color); icon_label.grid(row=0, column=0, rowspan=2, padx=(0, 10), sticky="ns")
            number_label = ttk.Label(card_frame, textvariable=number_var, style="CardNumber.TLabel"); number_label.grid(row=0, column=1, sticky="sw")
            text_label = ttk.Label(card_frame, text=text, style="CardText.TLabel"); text_label.grid(row=1, column=1, sticky="nw"); card_frame.columnconfigure(1, weight=1)
        
        # Criação dos quatro cards.
        create_dashboard_card(dashboard_frame, "📦", self.total_itens_var, "Itens Totais (Unidades)", "#007bff")
        create_dashboard_card(dashboard_frame, "🔖", self.tipos_unicos_var, "Tipos de Itens Únicos", "#17a2b8")
        create_dashboard_card(dashboard_frame, "⚠️", self.estoque_baixo_var, "Itens em Estoque Baixo", "#ffc107")
        create_dashboard_card(dashboard_frame, "↔️", self.mov_mes_var, "Movimentações este Mês", "#6c757d")

        # --- Seção de Ações Gerais (Pesquisa e Atualização) ---
        general_actions_frame = ttk.Frame(self.estoque_tab)
        general_actions_frame.pack(fill='x', padx=10, pady=(0, 10))
        search_frame = ttk.LabelFrame(general_actions_frame, text="Pesquisa de Estoque", padding="10")
        search_frame.pack(fill='x', expand=True, side="left")
        ttk.Label(search_frame, text="Buscar:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=50)
        search_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # O evento "<KeyRelease>" chama a função de filtro a cada tecla pressionada.
        search_entry.bind("<KeyRelease>", self.filtrar_equipamentos)
        
        refresh_button = ttk.Button(general_actions_frame, text="🔄 Atualizar Dados", command=self.refresh_with_feedback, width=20)
        refresh_button.pack(side="right", padx=10, ipady=8)

        # --- Seção do Formulário de Cadastro ---
        form_frame = ttk.Frame(self.estoque_tab, padding="20 10 20 20"); form_frame.pack(fill='x', padx=10, pady=10)
        ttk.Label(form_frame, text="Nome do Equipamento:").grid(row=0, column=0, padx=5, pady=5, sticky="w"); self.entry_nome = ttk.Entry(form_frame, width=30); self.entry_nome.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(form_frame, text="Nº de Série/SKU (Opcional):").grid(row=1, column=0, padx=5, pady=5, sticky="w"); self.entry_serie = ttk.Entry(form_frame, width=30); self.entry_serie.grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(form_frame, text="Descrição:").grid(row=0, column=2, padx=15, pady=5, sticky="w"); self.entry_descricao = ttk.Entry(form_frame, width=30); self.entry_descricao.grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(form_frame, text="Quantidade Inicial:").grid(row=1, column=2, padx=15, pady=5, sticky="w"); self.entry_quantidade = ttk.Entry(form_frame, width=10); self.entry_quantidade.grid(row=1, column=3, padx=5, pady=5, sticky="w"); self.entry_quantidade.insert(0, "1")
        ttk.Label(form_frame, text="Categoria:").grid(row=2, column=0, padx=5, pady=5, sticky="w"); self.combo_categoria = ttk.Combobox(form_frame, values=self.lista_categorias, width=28, state="readonly"); self.combo_categoria.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(form_frame, text="Estoque Mínimo:").grid(row=2, column=2, padx=15, pady=5, sticky="w"); self.entry_estoque_minimo = ttk.Entry(form_frame, width=10); self.entry_estoque_minimo.grid(row=2, column=3, padx=5, pady=5, sticky="w"); self.entry_estoque_minimo.insert(0, self.default_estoque_minimo)
        
        # --- Seção dos Botões de Ação do Formulário ---
        # Um frame único para agrupar todos os botões e facilitar o layout.
        buttons_container = ttk.Frame(form_frame)
        buttons_container.grid(row=3, column=0, columnspan=5, pady=10, sticky="ew")

        # Botões alinhados à esquerda (ações em itens selecionados).
        btn_movimentar = ttk.Button(buttons_container, text="↔️ Movimentar Estoque", command=self.abrir_janela_movimentacao)
        btn_movimentar.pack(side="left", padx=5)
        btn_editar = ttk.Button(buttons_container, text="✏️ Editar Item", command=self.abrir_janela_edicao)
        btn_editar.pack(side="left", padx=5)
        btn_historico = ttk.Button(buttons_container, text="📜 Ver Histórico", command=self.abrir_janela_historico)
        btn_historico.pack(side="left", padx=5)
        btn_excluir = ttk.Button(buttons_container, text="🗑️ Excluir Item", command=self.excluir_equipamento)
        btn_excluir.pack(side="left", padx=5)

        # Botões alinhados à direita (ações gerais e de adição).
        btn_relatorio = ttk.Button(buttons_container, text="📊 Emitir Relatório", command=self.abrir_janela_relatorio_opcoes)
        btn_relatorio.pack(side="right", padx=5)
        self.btn_adicionar = ttk.Button(buttons_container, text="Adicionar Equipamento", command=self.adicionar_equipamento)
        self.btn_adicionar.pack(side="right", padx=5)
        
        # --- Seção da Tabela de Equipamentos (Treeview) ---
        list_frame = ttk.Frame(self.estoque_tab, padding="20 10 20 20"); list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        self.tree = ttk.Treeview(list_frame, columns=("ID", "Nome", "Nº Série", "Categoria", "Descrição", "Qtd.", "Status"), show='headings', selectmode='extended')
        
        # Configuração dos cabeçalhos e colunas.
        self.tree.heading("ID", text="ID"); self.tree.heading("Nome", text="Nome"); self.tree.heading("Nº Série", text="Nº de Série/SKU"); self.tree.heading("Categoria", text="Categoria"); self.tree.heading("Descrição", text="Descrição"); self.tree.heading("Qtd.", text="Qtd."); self.tree.heading("Status", text="Status")
        self.tree.column("ID", width=40, anchor="center"); self.tree.column("Nome", width=250); self.tree.column("Nº Série", width=150); self.tree.column("Categoria", width=120); self.tree.column("Descrição", width=300); self.tree.column("Qtd.", width=60, anchor="center"); self.tree.column("Status", width=100, anchor="center")
        self.tree.pack(fill='both', expand=True, side='left')
         
        # Adiciona uma barra de rolagem vertical.
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview); scrollbar.pack(side='right', fill='y'); self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Associa o evento de duplo clique na tabela à função de ver histórico.
        self.tree.bind("<Double-1>", self.abrir_janela_historico)

    
    def criar_aba_movimentacao_setores(self):
        """Cria e organiza todos os widgets da aba 'Movimentação Entre Setores'."""
        
        # --- Seção do Formulário de Registro de Movimentação ---
        form_mov_frame = ttk.LabelFrame(self.mov_setores_tab, text="Registrar Nova Movimentação Entre Setores", padding="20")
        form_mov_frame.pack(fill="x", padx=10, pady=10)
        tipos_equip_mov = ["Kit (2x Monitores e 1 desktop)", "WebCam", "Monitor", "Desktop", "Leitor de código de barra"]
        ttk.Label(form_mov_frame, text="Tipo de Equipamento:").grid(row=0, column=0, padx=5, pady=5, sticky="w"); self.mov_tipo_equip_combo = ttk.Combobox(form_mov_frame, values=tipos_equip_mov, width=30, state="readonly"); self.mov_tipo_equip_combo.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        
        # --- Sub-formulário para Itens Normais ---
        self.normal_item_frame = ttk.Frame(form_mov_frame); 
        self.normal_item_frame.grid(row=1, column=0, columnspan=4, sticky="w")
        ttk.Label(self.normal_item_frame, text="Patrimônio:").grid(row=0, column=0, padx=5, pady=5, sticky="w"); 
        self.mov_entry_patrimonio = ttk.Entry(self.normal_item_frame, width=30); 
        self.mov_entry_patrimonio.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.mov_label_servicetag = ttk.Label(self.normal_item_frame, text="ServiceTag:"); 
        self.mov_label_servicetag.grid(row=0, column=2, padx=15, pady=5, sticky="w")
        self.mov_entry_servicetag = ttk.Entry(self.normal_item_frame, width=30); 
        self.mov_entry_servicetag.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # --- Sub-formulário para o item "Kit" ---
        self.kit_frame = ttk.LabelFrame(form_mov_frame, text="Componentes do Kit", padding=10); self.kit_frame.grid(row=1, column=0, columnspan=4, sticky="w", padx=5, pady=5)
        ttk.Label(self.kit_frame, text="Patrimônio").grid(row=0, column=1); ttk.Label(self.kit_frame, text="ServiceTag").grid(row=0, column=2)
        ttk.Label(self.kit_frame, text="Monitor 1:").grid(row=1, column=0, sticky="w"); self.kit_p1 = ttk.Entry(self.kit_frame); self.kit_p1.grid(row=1, column=1, padx=5); self.kit_s1 = ttk.Entry(self.kit_frame); self.kit_s1.grid(row=1, column=2, padx=5)
        ttk.Label(self.kit_frame, text="Monitor 2:").grid(row=2, column=0, sticky="w"); self.kit_p2 = ttk.Entry(self.kit_frame); self.kit_p2.grid(row=2, column=1, padx=5); self.kit_s2 = ttk.Entry(self.kit_frame); self.kit_s2.grid(row=2, column=2, padx=5)
        ttk.Label(self.kit_frame, text="Desktop:").grid(row=3, column=0, sticky="w"); self.kit_p3 = ttk.Entry(self.kit_frame); self.kit_p3.grid(row=3, column=1, padx=5); self.kit_s3 = ttk.Entry(self.kit_frame); self.kit_s3.grid(row=3, column=2, padx=5)
        
        # --- Campos Comuns a Todos os Tipos de Movimentação ---
        common_fields_frame = ttk.Frame(form_mov_frame); common_fields_frame.grid(row=2, column=0, columnspan=4, sticky="w")
        ttk.Label(common_fields_frame, text="Setor de Origem:").grid(row=0, column=0, padx=5, pady=5, sticky="w"); self.mov_combo_origem = ttk.Combobox(common_fields_frame, values=self.lista_destinos, width=30, state="readonly"); self.mov_combo_origem.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(common_fields_frame, text="Setor de Destino:").grid(row=0, column=2, padx=15, pady=5, sticky="w"); self.mov_combo_destino = ttk.Combobox(common_fields_frame, values=self.lista_destinos, width=30, state="readonly"); self.mov_combo_destino.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(common_fields_frame, text="Nº do Chamado:").grid(row=1, column=0, padx=5, pady=5, sticky="w"); self.mov_entry_chamado = ttk.Entry(common_fields_frame, width=30); self.mov_entry_chamado.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(common_fields_frame, text="Solicitante:").grid(row=1, column=2, padx=15, pady=5, sticky="w"); self.mov_entry_solicitante = ttk.Entry(common_fields_frame, width=30); self.mov_entry_solicitante.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(common_fields_frame, text="Responsável (seu nome):").grid(row=2, column=0, padx=5, pady=5, sticky="w"); self.mov_entry_responsavel = ttk.Entry(common_fields_frame, width=30); self.mov_entry_responsavel.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(common_fields_frame, text="Observação:").grid(row=3, column=0, padx=5, pady=5, sticky="nw"); self.mov_text_obs = tk.Text(common_fields_frame, width=80, height=3, font=("Roboto", 10)); self.mov_text_obs.grid(row=3, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        
        self.style.configure("Pendente.Treeview", background="#FFFACD") # Amarelo claro
        self.style.configure("Regularizado.Treeview", background="#64FF88") # Verde claro

        form_mov_frame = ttk.LabelFrame(self.mov_setores_tab, text="Registrar Nova Movimentação Entre Setores", padding="20")
        
        # --- Seção de Botões de Ação da Movimentação ---
        button_action_frame = ttk.Frame(common_fields_frame)
        button_action_frame.grid(row=4, column=0, columnspan=4, sticky="e", pady=10)
        
        btn_regularizar = ttk.Button(button_action_frame, text="✅ Marcar como Regularizado", command=self.marcar_como_regularizado)
        btn_regularizar.pack(side="left", padx=5)

        btn_refresh_setores = ttk.Button(button_action_frame, text="🔄 Atualizar Dados", command=self.refresh_with_feedback)
        btn_refresh_setores.pack(side="left", padx=5)

        btn_relatorio_setores = ttk.Button(button_action_frame, text="📊 Emitir Relatório", command=self.abrir_janela_relatorio_setores_opcoes)
        btn_relatorio_setores.pack(side="left", padx=5)
        
        btn_registrar_mov = ttk.Button(button_action_frame, text="Registrar Movimentação", command=self.registrar_movimentacao_setor)
        btn_registrar_mov.pack(side="left", padx=5)
        
        # --- Lógica para Alternar a Visibilidade dos Formulários (Kit vs. Normal) ---
        def _update_mov_setores_form(event=None):
            
            selected_item = self.mov_tipo_equip_combo.get()
            is_kit = selected_item == "Kit (2x Monitores e 1 desktop)"
            if is_kit: self.normal_item_frame.grid_remove(); self.kit_frame.grid() # Esconde o form normal / # Mostra o form do kit
            else:
                self.kit_frame.grid_remove(); self.normal_item_frame.grid() # Esconde o form do kit / Mostra o form normal
                 # Habilita/desabilita o campo ServiceTag dependendo do tipo de item.
                if selected_item in ["Monitor", "Desktop"]: self.mov_label_servicetag.config(state="normal"); self.mov_entry_servicetag.config(state="normal")
                else: self.mov_entry_servicetag.delete(0, tk.END); self.mov_label_servicetag.config(state="disabled"); self.mov_entry_servicetag.config(state="disabled")
        
        self.mov_tipo_equip_combo.bind("<<ComboboxSelected>>", _update_mov_setores_form); _update_mov_setores_form() # Chama a função uma vez para definir o estado inicial.
        
        # --- Seção do Histórico de Movimentações (Treeview) ---
        hist_mov_frame = ttk.LabelFrame(self.mov_setores_tab, text="Histórico de Movimentações Entre Setores", padding="20"); hist_mov_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.mov_setores_tree = ttk.Treeview(hist_mov_frame, columns=("ID", "Data", "Status", "Equipamento", "Patrimônio", "ServiceTag", "Origem", "Destino", "Responsável", "Chamado", "Solicitante"), show="headings", style="Multiline.Treeview")
        
        # Configuração dos cabeçalhos e colunas.
        self.mov_setores_tree.heading("ID", text="ID"); self.mov_setores_tree.heading("Data", text="Data"); self.mov_setores_tree.heading("Status", text="Status"); self.mov_setores_tree.heading("Equipamento", text="Equipamento"); self.mov_setores_tree.heading("Patrimônio", text="Patrimônio"); self.mov_setores_tree.heading("ServiceTag", text="ServiceTag"); self.mov_setores_tree.heading("Origem", text="Origem"); self.mov_setores_tree.heading("Destino", text="Destino"); self.mov_setores_tree.heading("Responsável", text="Responsável"); self.mov_setores_tree.heading("Chamado", text="Chamado"); self.mov_setores_tree.heading("Solicitante", text="Solicitante")
        self.mov_setores_tree.column("ID", width=40, anchor="center"); self.mov_setores_tree.column("Data", width=130); self.mov_setores_tree.column("Status", width=100, anchor="center"); self.mov_setores_tree.column("Equipamento", width=180); self.mov_setores_tree.column("Patrimônio", width=120); self.mov_setores_tree.column("ServiceTag", width=120); self.mov_setores_tree.column("Origem", width=120); self.mov_setores_tree.column("Destino", width=120); self.mov_setores_tree.column("Responsável", width=110); self.mov_setores_tree.column("Chamado", width=80); self.mov_setores_tree.column("Solicitante", width=110)
        self.mov_setores_tree.pack(fill="both", expand=True)

    # --- INÍCIO: FUNÇÕES DE LÓGICA E DADOS ---
    def conectar_google_sheets(self):
        
        """
        Estabelece a conexão com a API do Google Sheets usando as credenciais.
        Retorna True em caso de sucesso e False se ocorrer algum erro.
        """
        
        try:
            # Define o escopo de permissões da API.
            scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
           
            # Encontra o arquivo 'credentials.json' na mesma pasta do script.
            script_dir = os.path.dirname(os.path.abspath(__file__)); json_path = os.path.join(script_dir, 'credentials.json')
            
            # Autoriza o acesso usando as credenciais.
            creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope); client = gspread.authorize(creds)
            
            # Abre a planilha pelo nome e obtém acesso a cada aba (worksheet).
            self.spreadsheet = client.open(NOME_PLANILHA)
            self.equip_sheet = self.spreadsheet.worksheet("equipamentos"); self.mov_sheet = self.spreadsheet.worksheet("movimentacoes"); self.config_sheet = self.spreadsheet.worksheet("config"); self.mov_setores_sheet = self.spreadsheet.worksheet("movimentacoes_setores")
            return True
        except gspread.exceptions.WorksheetNotFound as e:
            # Erro específico se uma aba não for encontrada.
            messagebox.showerror("Erro de Planilha", f"A aba '{e.worksheet_name}' não foi encontrada. Verifique se ela foi criada corretamente."); self.destroy(); return False
        except Exception as e:
            # Erro genérico para problemas de conexão, autenticação, etc.
            messagebox.showerror("Erro de Conexão", f"Ocorreu um erro: {e}"); self.destroy(); return False
    
    def _load_config(self):
        
        """
        Carrega as configurações da aba 'config' da planilha para variáveis internas.
        Essas configurações (categorias, destinos) são usadas para popular ComboBoxes.
        """
        
        try:
            config_records = self.config_sheet.get_all_records(); config_df = pd.DataFrame(config_records)
            
            # Filtra e cria listas de valores únicos para 'destino' e 'categoria'.
            self.lista_destinos = config_df[config_df['parametro'] == 'destino']['valor'].tolist(); self.lista_destinos.sort()
            self.lista_categorias = config_df[config_df['parametro'] == 'categoria']['valor'].tolist(); self.lista_categorias.sort()
            
            # Pega o valor padrão para o estoque mínimo.
            self.default_estoque_minimo = config_df[config_df['parametro'] == 'default_estoque_minimo']['valor'].iloc[0]
        except (IndexError, KeyError, gspread.exceptions.GSpreadException) as e:
             messagebox.showerror("Erro de Configuração", f"A aba 'config' parece mal formatada ou vazia. Verifique os parâmetros e valores.\n\nErro: {e}"); self.destroy()
    
    def refresh_dataframes(self):
        
        """
        Busca todos os dados das planilhas e os carrega em DataFrames do Pandas.
        Realiza conversões de tipo para garantir que os dados numéricos e de data
        sejam tratados corretamente.
        """
        # Pega todos os registros e cria os DataFrames.
        self.equip_df = pd.DataFrame(self.equip_sheet.get_all_records()); self.mov_df = pd.DataFrame(self.mov_sheet.get_all_records()); self.mov_setores_df = pd.DataFrame(self.mov_setores_sheet.get_all_records())
        
        # Faz a conversão de colunas importantes para o tipo numérico.
        # 'errors=coerce' transforma valores inválidos em NaN, que são preenchidos com 1.
        if not self.equip_df.empty:
            for col in ['id', 'quantidade', 'estoque_minimo']: self.equip_df[col] = pd.to_numeric(self.equip_df[col], errors='coerce').fillna(1)
        if not self.mov_df.empty:
            for col in ['id_equipamento_fk', 'id_movimentacao']: self.mov_df[col] = pd.to_numeric(self.mov_df[col], errors='coerce')
            # Converte a coluna de data para o formato datetime do Pandas.
            self.mov_df['data_movimentacao_dt'] = pd.to_datetime(self.mov_df['data_movimentacao'], format='%d-%m-%Y %H:%M:%S', errors='coerce')
        if not self.mov_setores_df.empty and 'id' in self.mov_setores_df.columns:
            self.mov_setores_df['id'] = pd.to_numeric(self.mov_setores_df['id'])
            
            #Evita que ao colocar patrimonios com 6 dígitos ao emitir o relatório fica como NaN
            self.mov_setores_df['patrimonio'] = self.mov_setores_df['patrimonio'].astype(str)
            
        if not self.mov_setores_df.empty and 'data_movimentacao' in self.mov_setores_df.columns:
            self.mov_setores_df['data_movimentacao_dt'] = pd.to_datetime(self.mov_setores_df['data_movimentacao'], format='%d-%m-%Y %H:%M:%S', errors='coerce')
    
    def refresh_all_data(self):
        
        """
        Função centralizadora que atualiza todos os dados e a interface.
        É chamada no início e após qualquer alteração nos dados.
        """
        
        self.refresh_dataframes(); 
        self.update_dashboard(); 
        self.filtrar_equipamentos(); # Filtra com o termo atual (ou carrega tudo se vazio)
        self.carregar_mov_setores_treeview()

    
    def refresh_with_feedback(self):
        
        """
        Executa a atualização de dados, mas muda o cursor do mouse para 'espera'
        para dar um feedback visual ao usuário de que algo está acontecendo.
        """
        
        self.config(cursor="watch"); # Muda o cursor para "carregando"
        self.update_idletasks()  # Força a atualização da UI
        self.refresh_all_data()
        self.config(cursor="") # Retorna o cursor ao normal

    def carregar_mov_setores_treeview(self):
        
        """Popula a tabela de histórico de movimentação entre setores com os dados do DataFrame."""
        
        # Limpa a tabela antes de inserir novos dados
        for i in self.mov_setores_tree.get_children(): self.mov_setores_tree.delete(i)

        # Configura as 'tags' para colorir as linhas com base no status.
        self.mov_setores_tree.tag_configure('Pendente', background="#FAF6C8") # Amarelo claro
        self.mov_setores_tree.tag_configure('Regularizado', background="#CFF7D1") #Verde Claro

        if hasattr(self, 'mov_setores_df') and not self.mov_setores_df.empty:
            # Ordena pelo ID de forma decrescente para mostrar os mais recentes primeiro.
            df_sorted = self.mov_setores_df.sort_values(by="id", ascending=False)

            # Garante que colunas opcionais existam para evitar erros.
            for col in ['chamado', 'solicitante', 'status_regularizacao']:
                if col not in df_sorted.columns: df_sorted[col] = ''
            df_sorted = df_sorted.fillna('') # Substitui valores nulos (NaN) por strings vazias.

            # Seleciona e reordena as colunas para exibição.
            df_display = df_sorted[["id", "data_movimentacao", "status_regularizacao", "tipo_equipamento", "patrimonio", "servicetag", "setor_origem", "setor_destino", "responsavel", "chamado", "solicitante"]]

            # Itera sobre o DataFrame e insere cada linha na tabela (Treeview).
            for index, row in df_display.iterrows():
                status = row['status_regularizacao'] if row['status_regularizacao'] else 'Pendente'
                tag = status if status in ['Pendente', 'Regularizado'] else 'Pendente'
                self.mov_setores_tree.insert("", "end", values=list(row), tags=(tag,))

    def registrar_movimentacao_setor(self):
        
        """Coleta os dados do formulário e registra uma nova movimentação entre setores na planilha."""
        
        # --- Coleta de Dados do Formulário ---
        tipo_equip = self.mov_tipo_equip_combo.get(); origem = self.mov_combo_origem.get(); destino = self.mov_combo_destino.get(); responsavel = self.mov_entry_responsavel.get(); obs = self.mov_text_obs.get("1.0", tk.END).strip()
        chamado = self.mov_entry_chamado.get(); solicitante = self.mov_entry_solicitante.get()
        patrimonio = ""; servicetag = ""
        
        # --- Lógica Específica para o "Kit" ---
        if tipo_equip == "Kit (2x Monitores e 1 desktop)":
            p1=self.kit_p1.get(); s1=self.kit_s1.get(); p2=self.kit_p2.get(); s2=self.kit_s2.get(); p3=self.kit_p3.get(); s3=self.kit_s3.get()
            if not all([p1, s1, p2, s2, p3, s3]): messagebox.showwarning("Campos Obrigatórios", "Para o Kit, todos os 6 campos (Patrimônio e ServiceTag) devem ser preenchidos."); return
            
            # Concatena os dados do kit em strings com quebra de linha para salvar na planilha.
            patrimonio = f"Monitor 1: {p1}\nMonitor 2: {p2}\nDesktop: {p3}"
            ervicetag = f"Monitor 1: {s1}\nMonitor 2: {s2}\nDesktop: {s3}"
        else:
            patrimonio = self.mov_entry_patrimonio.get(); servicetag = self.mov_entry_servicetag.get()
            if not patrimonio: messagebox.showwarning("Campo Obrigatório", "O campo Patrimônio é obrigatório."); return
            if tipo_equip in ["Monitor", "Desktop"] and not servicetag: messagebox.showwarning("Campo Obrigatório", "Para este tipo, o ServiceTag é obrigatório."); return
       
        # --- Validação dos Campos ---
        if not all([tipo_equip, origem, destino, responsavel, chamado, solicitante]):
            messagebox.showwarning("Campos Obrigatórios", "Preencha todos os campos necessários antes de registrar."); return
        if origem == destino: messagebox.showwarning("Movimentação Inválida", "O setor de origem não pode ser o mesmo que o de destino."); return
        
        # --- Registro na Planilha ---
        novo_id = self._get_next_id(self.mov_setores_sheet); data = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        nova_linha = [novo_id, data, responsavel, tipo_equip, patrimonio, servicetag, origem, destino, obs, chamado, solicitante, 'Pendente']
        self.mov_setores_sheet.append_row(nova_linha); messagebox.showinfo("Sucesso", "Movimentação entre setores registrada com sucesso!")
        
        # --- Limpeza do Formulário ---
        self.mov_tipo_equip_combo.set(''); self.mov_entry_patrimonio.delete(0, tk.END); self.mov_entry_servicetag.delete(0, tk.END)
        self.kit_p1.delete(0, tk.END); self.kit_s1.delete(0, tk.END); self.kit_p2.delete(0, tk.END); self.kit_s2.delete(0, tk.END); self.kit_p3.delete(0, tk.END); self.kit_s3.delete(0, tk.END)
        self.mov_combo_origem.set(''); self.mov_combo_destino.set(''); self.mov_entry_responsavel.delete(0, tk.END); self.mov_text_obs.delete("1.0", tk.END)
        self.mov_entry_chamado.delete(0, tk.END); self.mov_entry_solicitante.delete(0, tk.END)
        
        # Atualiza a interface para mostrar o novo registro.
        self.refresh_all_data()

    def update_dashboard(self):
        
        """
        Calcula as estatísticas do dashboard e atualiza os valores nos cards da interface.
        É chamado sempre que os dados são atualizados.
        """
        # --- Lógica para o Card de Total de Itens e Tipos Únicos ---
        if self.equip_df.empty: self.total_itens_var.set("0"); self.tipos_unicos_var.set("0"); self.estoque_baixo_var.set("0")
        
        # Se não houver dados, zera os contadores para evitar erros.
        else:
            # Filtra o DataFrame para considerar apenas itens que não foram descartados.
            df_ativo = self.equip_df[self.equip_df['status'] != 'Descartado']
            
            # Soma a coluna 'quantidade' de todos os itens ativos.
            self.total_itens_var.set(str(df_ativo['quantidade'].sum())); 
            
            # Conta o número de linhas (tipos de itens únicos) no DataFrame filtrado.
            self.tipos_unicos_var.set(str(len(df_ativo)))
            
            # Conta quantos itens têm a quantidade atual menor ou igual ao estoque mínimo.
            low_stock_count = (df_ativo['quantidade'] <= df_ativo['estoque_minimo']).sum(); self.estoque_baixo_var.set(str(low_stock_count))
        
        # --- Lógica para o Card de Movimentações no Mês --
        if self.mov_df.empty: self.mov_mes_var.set("0")
        else:
            now = datetime.datetime.now()
            
            # Filtra o DataFrame de movimentações para contar apenas os registros do mês e ano atuais.
            mov_mes_count = len(self.mov_df[(self.mov_df['data_movimentacao_dt'].dt.month == now.month) & (self.mov_df['data_movimentacao_dt'].dt.year == now.year)])
            self.mov_mes_var.set(str(mov_mes_count))
    def filtrar_equipamentos(self, event=None):
        
        """
        Filtra os equipamentos na tabela (Treeview) com base no termo digitado na barra de pesquisa.
        Esta função é chamada a cada tecla pressionada no campo de busca.
        """
        
        # Pega o texto da busca e converte para minúsculas para uma busca não sensível a maiúsculas.
        search_term = self.search_var.get().lower()
        
        # Se o campo de busca estiver vazio, carrega todos os equipamentos.
        if not search_term: self.carregar_equipamentos_treeview(); return
        if not self.equip_df.empty:
            
            # Filtra o DataFrame, procurando o termo de busca em várias colunas.
            # `astype(str).str.lower().str.contains` garante que a busca funcione em qualquer tipo de dado.
            df_filtered = self.equip_df[
                self.equip_df['nome'].astype(str).str.lower().str.contains(search_term, na=False) | 
                self.equip_df['numero_serie'].astype(str).str.lower().str.contains(search_term, na=False) | 
                self.equip_df['categoria'].astype(str).str.lower().str.contains(search_term, na=False) | 
                self.equip_df['descricao'].astype(str).str.lower().str.contains(search_term, na=False)
                ]
            
            # Popula a tabela com o DataFrame já filtrado.
            self.populate_treeview(df_filtered)
    def carregar_equipamentos_treeview(self): 
        """Função de atalho para popular a tabela com todos os equipamentos (sem filtro)."""
        
        self.populate_treeview(self.equip_df) 
    def populate_treeview(self, df):
        
        """
        Limpa a tabela de equipamentos e a preenche com os dados de um DataFrame.
        
        Args:
            df (pd.DataFrame): O DataFrame contendo os itens a serem exibidos.
        """
        
        # Primeiro, apaga todas as linhas existentes na tabela para evitar duplicatas.
        for i in self.tree.get_children(): self.tree.delete(i)
        if not df.empty:
            # Ordena pelo ID de forma decrescente para mostrar os itens mais recentes primeiro.
            df_sorted = df.sort_values(by="id", ascending=False)
            
            # Seleciona e reordena as colunas para garantir que elas correspondam à ordem na Treeview.
            df_display = df_sorted[["id", "nome", "numero_serie", "categoria", "descricao", "quantidade", "status"]]
            
            # Itera sobre cada linha do DataFrame e a insere na tabela.
            for index, row in df_display.iterrows(): self.tree.insert("", "end", values=list(row))
            
    def _get_next_id(self, sheet):
        
        """
        Calcula o próximo ID disponível para um novo registro em uma determinada aba da planilha.
        
        Args:
            sheet (gspread.Worksheet): A aba da planilha a ser verificada.
            
        Returns:
            int: O próximo ID a ser usado.
        """
        
        # Pega todos os valores da primeira coluna (coluna de IDs), ignorando o cabeçalho.
        ids = sheet.col_values(1)[1:]; 
        
        # Se não houver IDs, o primeiro será 1.
        if not ids: return 1
        
        # Converte todos os valores para inteiros (ignorando textos) e retorna o maior valor + 1.
        return max([int(i) for i in ids if str(i).isdigit()], default=0) + 1
    
    def _find_sheet_row_index_by_id(self, df, record_id):
        
        """
        Encontra o número da linha na planilha do Google correspondente a um ID específico.
        
        Args:
            df (pd.DataFrame): O DataFrame onde a busca será feita.
            record_id (int): O ID do registro a ser encontrado.
            
        Returns:
            int or None: O número da linha na planilha, ou None se não for encontrado.
        """
        
        try:
            df['id'] = pd.to_numeric(df['id']);
            # Encontra o índice do DataFrame (baseado em 0) para o ID fornecido. 
            df_index = df.index[df['id'] == record_id].tolist()[0]
            # Retorna o índice do DataFrame + 2.
            # O +2 é necessário porque:
            # +1 para converter o índice de 0 para 1 (planilhas começam na linha 1).
            # +1 para pular a linha do cabeçalho da planilha.
            return df_index + 2
        except IndexError: return None  #Retorna None se o ID não for encontrado no DataFrame.
        
    def adicionar_equipamento(self):
        
        """
        Coleta os dados do formulário, valida e adiciona um novo equipamento à planilha.
        """
        
        # --- Coleta de Dados dos Campos de Entrada (Entries) ---
        categoria = self.combo_categoria.get(); 
        nome = self.entry_nome.get(); 
        serie = self.entry_serie.get(); 
        descricao = self.entry_descricao.get(); 
        quantidade_str = self.entry_quantidade.get(); 
        estoque_minimo_str = self.entry_estoque_minimo.get()
        
        # --- Validação dos Dados ---
        if not nome or not categoria: messagebox.showwarning("Campos Vazios", "Os campos 'Nome' e 'Categoria' são obrigatórios."); return
        try:
            quantidade = int(quantidade_str); estoque_minimo = int(estoque_minimo_str)
            if quantidade < 0 or estoque_minimo < 0: messagebox.showwarning("Valor Inválido", "As quantidades não podem ser negativas."); return
        except ValueError: messagebox.showwarning("Valor Inválido", "As quantidades devem ser números inteiros."); return
        
        # --- Preparação e Registro dos Dados ---
        novo_id = self._get_next_id(self.equip_sheet); 
        status = "Em Estoque" if quantidade > 0 else "Fora de Estoque"; 
        data_cadastro = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        
        # A ordem dos itens na lista deve ser EXATAMENTE a mesma ordem das colunas na planilha.
        nova_linha = [novo_id, nome, serie, descricao, quantidade, status, data_cadastro, estoque_minimo, categoria]
        self.equip_sheet.append_row(nova_linha); 
        messagebox.showinfo("Sucesso", "Equipamento adicionado com sucesso!")
        
        # --- Limpeza do Formulário Após Adicionar ---
        self.entry_nome.delete(0, tk.END); 
        self.entry_serie.delete(0, tk.END); 
        self.entry_descricao.delete(0, tk.END); 
        self.entry_quantidade.delete(0, tk.END); 
        self.entry_estoque_minimo.delete(0, tk.END); 
        self.entry_estoque_minimo.insert(0, self.default_estoque_minimo); 
        self.combo_categoria.set('')
        
        # Atualiza a interface para exibir o novo item.
        self.refresh_all_data()
        
    def salvar_edicao(self, item_id, nome, categoria, serie, descricao, quantidade_str, estoque_minimo_str, window):
        
        """
        Salva as alterações feitas na janela de edição de um equipamento.
        
        Args:
            item_id (int): O ID do item que está sendo editado.
            window (tk.Toplevel): A referência da janela de edição para poder fechá-la.
            ...outros args: Os novos valores vindos dos campos de edição.
        """
        # --- Validação dos Dados ---
        if not nome or not categoria: 
            messagebox.showwarning("Campo Vazio", "Os campos 'Nome' e 'Categoria' não podem ficar vazios.", parent=window); 
            return
        try:
            quantidade = int(quantidade_str); estoque_minimo = int(estoque_minimo_str)
            if quantidade < 0 or estoque_minimo < 0: 
                messagebox.showwarning("Valor Inválido", "As quantidades não podem ser negativas.", parent=window); 
                return
        except ValueError: 
            messagebox.showwarning("Valor Inválido", "As quantidades devem ser números inteiros.", parent=window); 
            return
        
        # --- Atualização na Planilha --
        row_index = self._find_sheet_row_index_by_id(self.equip_df, item_id)
        if not row_index: 
            messagebox.showerror("Erro", "Não foi possível encontrar o item para atualizar.", parent=window); 
            return
        novo_status = "Em Estoque" if quantidade > 0 else "Fora de Estoque"
        
        # Preserva a data de cadastro original.
        data_cadastro_original = self.equip_df.loc[self.equip_df['id'] == item_id, 'data_cadastro'].iloc[0]
        
        # Monta a linha com os dados atualizados na ordem correta das colunas.
        linha_atualizada = [item_id, nome, serie, descricao, quantidade, novo_status, data_cadastro_original, estoque_minimo, categoria]
        
        # Atualiza a linha inteira na planilha de uma só vez para maior eficiência.
        self.equip_sheet.update(f'A{row_index}:I{row_index}', [linha_atualizada]); 
        messagebox.showinfo("Sucesso", "Equipamento atualizado com sucesso!")
        
        window.destroy(); # Fecha a janela de edição.
        self.refresh_all_data() # Atualiza a interface.
        
    def excluir_equipamento(self):
        
        """
        Exclui o(s) equipamento(s) selecionado(s) da planilha.
        """
        
        selected_items = self.tree.selection()
        if not selected_items: messagebox.showwarning("Nenhum Item Selecionado", "Selecione um ou mais equipamentos para excluir."); return
        
        # Pede confirmação do usuário, uma boa prática para ações destrutivas.
        confirm = messagebox.askyesno("Confirmar Exclusão", f"Você tem certeza que deseja excluir os {len(selected_items)} equipamentos selecionados? O histórico de movimentações NÃO será afetado.")
        if confirm:
            rows_to_delete = []; 
            
            # Primeiro, coleta os números de todas as linhas a serem deletadas.
            for item in selected_items:
                item_id = int(self.tree.item(item, "values")[0]); 
                row_index = self._find_sheet_row_index_by_id(self.equip_df, item_id)
                
                if row_index: 
                    rows_to_delete.append(row_index)
                    
            # Deleta as linhas em ordem decrescente.
            # Isso é CRUCIAL para evitar que os índices das linhas mudem durante a exclusão.
            for row_index in sorted(rows_to_delete, reverse=True): 
                self.equip_sheet.delete_rows(row_index)
                
            messagebox.showinfo("Sucesso", f"{len(selected_items)} equipamento(s) excluído(s) com sucesso."); 
            
            self.refresh_all_data()
            
    def get_last_movement_info(self, item_id):
        
        """
        Busca a última movimentação de 'Saída' de um item para sugerir a origem em uma devolução.
        
        Args:
            item_id (int): O ID do equipamento.
            
        Returns:
            tuple or None: Uma tupla com (destino, solicitante) ou None se não houver saídas.
        """
        
        if self.mov_df.empty or 'id_equipamento_fk' not in self.mov_df.columns: 
            return None
        
        # Filtra todas as movimentações do tipo 'Saída' para o item específico.
        movs_item = self.mov_df[(self.mov_df['id_equipamento_fk'] == item_id) & (self.mov_df['tipo_movimentacao'] == 'Saída')]
        
        if not movs_item.empty:
            # Ordena pela ID da movimentação (mais recente) e pega a primeira linha.
            last_mov = movs_item.sort_values(by="id_movimentacao", ascending=False).iloc[0]; return (last_mov['destino_origem'], last_mov['solicitante'])
        return None
    
    def abrir_janela_movimentacao(self):
        
        """
        Cria e configura a janela de movimentação de estoque com base nos itens selecionados.
        Esta função lida com uma interface complexa e dinâmica.
        """
        
        selected_items_ids = self.tree.selection()
        if not selected_items_ids: 
            messagebox.showwarning("Nenhum Item Selecionado", "Por favor, selecione um ou mais equipamentos."); 
            return
        
        # Pega os dados completos dos itens selecionados.
        items_para_movimentar = self.get_data_from_tree_selection(selected_items_ids)
        
        # --- Cálculo Dinâmico da Altura da Janela ---
        base_height = 340; 
        height_per_item = 40; 
        window_height = base_height + (len(items_para_movimentar) * height_per_item)
        
        # --- Criação e Configuração da Janela (Toplevel) ---
        mov_window = tk.Toplevel(self); 
        mov_window.title(f"Movimentar {len(items_para_movimentar)} Iten(s)"); 
        mov_window.geometry(f"550x{window_height}"); 
        mov_window.resizable(False, False); mov_window.transient(self); # Mantém a janela sobre a principal.
        mov_window.grab_set() # Bloqueia interação com a janela principal.
        
        # --- Criação dos Frames Principais ---
        main_content_frame = ttk.Frame(mov_window); 
        main_content_frame.pack(fill="both", expand=True)
        button_frame = ttk.Frame(mov_window, padding="10"); 
        button_frame.pack(side="bottom", fill="x")
        
        # --- Widgets Comuns a Todas as Movimentações ---
        shared_frame = ttk.Frame(main_content_frame, padding="20 10"); 
        shared_frame.pack(fill="x")
        ttk.Label(shared_frame, text="Tipo de Movimentação:", font=("Roboto", 10, "bold")).grid(row=0, column=0, padx=5, pady=8, sticky="w")
        mov_type_var = tk.StringVar(value="Saída")
        saida_rb = ttk.Radiobutton(shared_frame, text="Saída", variable=mov_type_var, value="Saída"); 
        entrada_rb = ttk.Radiobutton(shared_frame, text="Entrada/Dev.", variable=mov_type_var, value="Entrada"); 
        descarte_rb = ttk.Radiobutton(shared_frame, text="Descarte", variable=mov_type_var, value="Descarte")
        saida_rb.grid(row=0, column=1, sticky="w", padx=5); 
        entrada_rb.grid(row=0, column=1); descarte_rb.grid(row=0, column=1, sticky="e", padx=5)
        saida_entrada_frame = ttk.Frame(shared_frame)
        descarte_frame = ttk.Frame(shared_frame)
        ttk.Label(shared_frame, text="Seu Nome (Responsável):").grid(row=1, column=0, padx=5, pady=8, sticky="w"); 
        
        # --- Frames e Widgets Específicos para cada Tipo de Movimentação ---
        entry_responsavel = ttk.Entry(shared_frame, width=40); # Para campos de Saída/Entrada
        entry_responsavel.grid(row=1, column=1, padx=5, pady=8) # Para campos de Descarte
        
        # Widgets de Saída/Entrada
        ttk.Label(saida_entrada_frame, text="Nome do Solicitante:").grid(row=0, column=0, padx=5, pady=8, sticky="w"); 
        entry_solicitante = ttk.Entry(saida_entrada_frame, width=40); 
        entry_solicitante.grid(row=0, column=1, padx=5, pady=8)
        ttk.Label(saida_entrada_frame, text="Nº do Chamado (Opcional):").grid(row=1, column=0, padx=5, pady=8, sticky="w"); 
        entry_chamado = ttk.Entry(saida_entrada_frame, width=40); 
        entry_chamado.grid(row=1, column=1, padx=5, pady=8)
        label_destino_origem = ttk.Label(saida_entrada_frame, text="Destino:"); 
        label_destino_origem.grid(row=2, column=0, padx=5, pady=8, sticky="w")
        combo_destino_saida = ttk.Combobox(saida_entrada_frame, values=self.lista_destinos, width=38, state="readonly")
        combo_origem_entrada = ttk.Combobox(saida_entrada_frame, width=38)
        outra_origem_frame = ttk.Frame(saida_entrada_frame)
        ttk.Label(outra_origem_frame, text="Especifique:").pack(side="left", padx=5); 
        
        # Widget de Descarte
        entry_origem_outros = ttk.Entry(outra_origem_frame, width=40); 
        entry_origem_outros.pack(side="left")
        ttk.Label(descarte_frame, text="Motivo/Laudo (Obrigatório):").grid(row=0, column=0, padx=5, pady=8, sticky="nw"); 
        text_laudo = scrolledtext.ScrolledText(descarte_frame, width=38, height=4, wrap=tk.WORD); text_laudo.grid(row=0, column=1, padx=5, pady=8)
        
        self.extra_field_visible = False # Flag para controlar a visibilidade do campo "Outra Origem"

        # --- Lógica Dinâmica da Interface ---
        def toggle_mov_type(*args):
            
            """
            Função interna que mostra/esconde os campos do formulário
            com base no tipo de movimentação selecionado (Saída, Entrada, Descarte).
            """
            
            tipo = mov_type_var.get(); 
             # Esconde todos os frames específicos primeiro.
             
            saida_entrada_frame.grid_remove(); 
            descarte_frame.grid_remove(); outra_origem_frame.grid_remove(); 
            self.extra_field_visible = False
            
            if tipo == "Descarte":
                
                 # Mostra o frame de descarte e preenche a quantidade com o total em estoque.
                descarte_frame.grid(row=2, column=0, columnspan=2, sticky="w")
                for item_data, qtd_entry in movimentacao_details:
                    qtd_atual = item_data['quantidade']; 
                    qtd_entry.config(state="normal"); 
                    qtd_entry.delete(0, tk.END); qtd_entry.insert(0, str(qtd_atual)); 
                    qtd_entry.config(state="disabled")  #Desabilita a edição da quantidade.
                    
            else: # Lógica para Saída e Entrada
                
                saida_entrada_frame.grid(row=2, column=0, columnspan=2, sticky="w")
                # Reseta a quantidade para 1 e habilita a edição.
                for item_data, qtd_entry in movimentacao_details: 
                    qtd_entry.config(state="normal"); 
                    qtd_entry.delete(0, tk.END); 
                    qtd_entry.insert(0, "1")
                    
                # Esconde os widgets de destino/origem para configurá-los corretamente.
                combo_destino_saida.grid_remove(); 
                combo_origem_entrada.grid_remove(); 
                entry_origem_outros.grid_remove()
                
                if tipo == "Saída":
                    label_destino_origem.config(text="Destino:"); combo_destino_saida.grid(row=2, column=1, padx=5, pady=8)
                else: #Entrada
                    label_destino_origem.config(text="Origem / Devolvido por:")
                    # Se for apenas um item, tenta sugerir a última origem.
                    
                    if len(items_para_movimentar) == 1:
                        item_id = int(items_para_movimentar[0]['id']); 
                        last_mov = self.get_last_movement_info(item_id)
                        opcoes = ["Item Novo (Entrada inicial)", "Outra Origem (Especificar)"]
                        
                        if last_mov:
                            destino, solicitante = last_mov; 
                            sugestao = f"{destino} (devolvido por {solicitante})"; 
                            opcoes.insert(0, sugestao)
                        combo_origem_entrada.config(values=opcoes); 
                        combo_origem_entrada.grid(row=2, column=1, padx=5, pady=8)
                        
                    else: # Se forem múltiplos itens, força o usuário a digitar a origem.
                        entry_origem_outros.pack_forget(); 
                        ttk.Label(outra_origem_frame, text="Origem (múltiplos itens):").pack(side="left", padx=5); 
                        entry_origem_outros.pack(side="left")
                        outra_origem_frame.grid(row=2, column=1, sticky='w', padx=5, pady=8); 
                        self.extra_field_visible = True
                        
        def on_origem_selected(event):
            
            """Redimensiona a janela se o campo 'Outra Origem' for exibido/ocultado."""
            
            height_extra_row = 50; 
            current_width = mov_window.winfo_width(); 
            current_height = mov_window.winfo_height()
            
            if combo_origem_entrada.get() == "Outra Origem (Especificar)":
                if not self.extra_field_visible:
                    outra_origem_frame.grid(row=3, column=0, columnspan=2, sticky="w", pady=8); 
                    mov_window.geometry(f"{current_width}x{current_height + height_extra_row}")
                    self.extra_field_visible = True
            else:
                if self.extra_field_visible:
                    outra_origem_frame.grid_forget(); 
                    mov_window.geometry(f"{current_width}x{current_height - height_extra_row}"); 
                    self.extra_field_visible = False
        # Vincula os eventos às funções de controle da interface.           
        combo_origem_entrada.bind("<<ComboboxSelected>>", on_origem_selected); 
        mov_type_var.trace("w", toggle_mov_type)
        
        # --- Cria a Lista de Itens a Movimentar ---
        items_frame = ttk.LabelFrame(main_content_frame, text="Itens a Movimentar", padding="20 10"); 
        items_frame.pack(fill="both", expand=True, padx=20, pady=10)
        ttk.Label(items_frame, text="Item (Estoque Atual)", font=("Roboto", 10, "bold")).grid(row=0, column=0, sticky="w"); 
        ttk.Label(items_frame, text="Qtd. a Mover", font=("Roboto", 10, "bold")).grid(row=0, column=1, sticky="w", padx=10)
        
        movimentacao_details = [] # Lista para guardar os dados e widgets de cada item.
        
        for i, item_data in enumerate(items_para_movimentar, start=1):
            nome, qtd_atual = item_data['nome'], item_data['quantidade']
            label_text = f"{nome} (Estoque: {qtd_atual})"; 
            item_label = ttk.Label(items_frame, text=label_text, wraplength=350); 
            item_label.grid(row=i, column=0, sticky="w", pady=5); 
            qtd_entry = ttk.Entry(items_frame, width=10); 
            qtd_entry.grid(row=i, column=1, padx=10); qtd_entry.insert(0, "1"); 
            movimentacao_details.append((item_data, qtd_entry))
            
        toggle_mov_type() # Chama a função uma vez para configurar o estado inicial da janela.
        
        def get_destino_origem_value():
            
            """Retorna o valor do campo de destino/origem correto com base no contexto."""
            
            if mov_type_var.get() == "Descarte": return ""
            if mov_type_var.get() == 'Saída': return combo_destino_saida.get()
            else: #Entrada
                if len(items_para_movimentar) > 1: return entry_origem_outros.get()
                else:
                    valor_selecionado = combo_origem_entrada.get()
                    if valor_selecionado == "Outra Origem (Especificar)": return entry_origem_outros.get()
                    return valor_selecionado
                
        # --- Botão de Confirmação ---
        # O lambda é usado para passar todos os valores atuais dos widgets para a função de confirmação.
        btn_confirmar = ttk.Button(button_frame, text="Confirmar Movimentação", command=lambda: self.confirmar_movimentacao(mov_type_var.get(), movimentacao_details, entry_responsavel.get(), entry_solicitante.get(), entry_chamado.get(), get_destino_origem_value(), text_laudo.get("1.0", tk.END), mov_window))
        btn_confirmar.pack()
    
    def confirmar_movimentacao(self, tipo_mov, movimentacao_details, responsavel, solicitante, chamado, destino_origem, motivo_laudo, window):
        
        """
        Valida e executa a movimentação de estoque para os itens selecionados.
        Esta função é a lógica "backend" da janela de movimentação, chamada após o clique em "Confirmar".

        Args:
            tipo_mov (str): O tipo de movimentação ('Saída', 'Entrada', 'Descarte').
            movimentacao_details (list): Lista de tuplas, onde cada tupla contém (dados_do_item, widget_de_quantidade).
            responsavel (str): Nome do responsável pela movimentação.
            solicitante (str): Nome do solicitante (para Entradas/Saídas).
            chamado (str): Número do chamado associado.
            destino_origem (str): Destino (para Saída) ou Origem (para Entrada).
            motivo_laudo (str): Justificativa para o Descarte.
            window (tk.Toplevel): A janela de movimentação, para que possa ser fechada no final.
        """
        
        # --- Validação dos Campos Obrigatórios ---
        # Verifica se os campos essenciais para cada tipo de movimentação foram preenchidos.
        if not responsavel:
            messagebox.showwarning("Campo Vazio", "O campo 'Seu Nome (Responsável)' é obrigatório.", parent=window)
            return
    
        if tipo_mov in ['Saída', 'Entrada'] and not destino_origem: 
            messagebox.showwarning("Campos Vazios", "Preencha o campo 'Destino/Origem'.", parent=window); 
            return
        
        if tipo_mov == 'Descarte' and (not motivo_laudo or motivo_laudo.strip() == ""): 
            messagebox.showwarning("Campo Obrigatório", "Para o descarte, o campo 'Motivo / Laudo' deve ser preenchido.", parent=window); 
            return
        
        # --- Validação das Quantidades para Cada Item ---
        items_validados = []
        
        for item_data, qtd_entry in movimentacao_details:
            nome_item, qtd_atual = item_data['nome'], item_data['quantidade']
            try:
                qtd_mov = int(qtd_entry.get())
                if qtd_mov <= 0: 
                    messagebox.showerror("Valor Inválido", f"A quantidade para '{nome_item}' deve ser maior que zero.", parent=window); 
                    return
                
                # Validação crucial para saídas: não permitir que saia mais do que há em estoque.
                if tipo_mov == 'Saída' and qtd_mov > qtd_atual: 
                    messagebox.showerror("Estoque Insuficiente", f"Item '{nome_item}': A quantidade a mover ({qtd_mov}) excede o estoque ({qtd_atual}).", parent=window); 
                    return
                
                 # Adiciona o item e sua quantidade a uma lista de itens prontos para serem processados.
                items_validados.append({'data': item_data, 'qtd_a_mover': qtd_mov})
            except ValueError: 
                messagebox.showerror("Valor Inválido", f"A quantidade para '{nome_item}' deve ser um número inteiro.", parent=window)
                return
            
        # --- Processamento e Atualização dos Dados na Planilha ---
        data_mov = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        novas_movimentacoes = [] # Lista para armazenar todos os registros de movimentação a serem adicionados.
        
        for item in items_validados:
            item_id, qtd_atual = item['data']['id'], item['data']['quantidade']
            qtd_a_mover = item['qtd_a_mover']
            
            # Calcula a nova quantidade em estoque com base no tipo de movimentação.
            if tipo_mov == 'Saída': 
                nova_qtd = qtd_atual - qtd_a_mover; 
                novo_status = "Em Estoque" if nova_qtd > 0 else "Fora de Estoque"
            elif tipo_mov == 'Entrada': 
                nova_qtd = qtd_atual + qtd_a_mover; 
                novo_status = "Em Estoque"
            else: 
                nova_qtd = 0; 
                novo_status = "Descartado"
                
            # Atualiza a quantidade e o status na aba 'equipamentos'.    
            row_index = self._find_sheet_row_index_by_id(self.equip_df, item_id)
            
            if row_index:
                self.equip_sheet.update_cell(row_index, 5, nova_qtd) # Coluna 5 é 'quantidade'
                self.equip_sheet.update_cell(row_index, 6, novo_status) # Coluna 6 é 'status'
                
            # Prepara a nova linha para ser adicionada na aba 'movimentacoes'.   
            mov_id = self._get_next_id(self.mov_sheet) + len(novas_movimentacoes)
            novas_movimentacoes.append([
                mov_id, item_id, tipo_mov, qtd_a_mover, destino_origem, 
                solicitante if tipo_mov != 'Descarte' else '', chamado if tipo_mov != 'Descarte' else '',
                responsavel, data_mov, 
                motivo_laudo.strip() if tipo_mov == 'Descarte' else ''])
        
        # --- Registro e Atualização Final ---
        # Adiciona todas as novas movimentações à planilha de uma só vez.    
        if novas_movimentacoes: 
            self.mov_sheet.append_rows(novas_movimentacoes)
            
        messagebox.showinfo("Sucesso", "Movimentação registrada com sucesso!")
        window.destroy() # Fecha a janela de movimentação. 
        self.refresh_all_data() # Atualiza toda a interface.

    def abrir_janela_historico(self, event=None):
        
        """
        Abre uma nova janela para exibir o histórico de movimentações do(s) item(ns) selecionado(s).
        Pode ser chamada por um botão ou por um duplo clique na tabela ('event').
        """
        
        # --- Coleta e Validação da Seleção ---
        selected_items_ids = self.tree.selection()
        if not selected_items_ids:
            # Lógica para tratar o duplo clique, que não "seleciona" o item, mas o "foca".
            if event:
                item_selecionado = self.tree.focus()
                if not item_selecionado: return
                selected_items_ids = [item_selecionado]
            else: 
                messagebox.showwarning("Nenhum Item Selecionado", "Selecione um ou mais equipamentos."); 
                return
        # Obtém os IDs numéricos (da coluna 0) dos itens selecionados na tabela.   
        ids_para_buscar = [int(self.tree.item(item_id, "values")[0]) for item_id in selected_items_ids]
        
        # --- Criação da Janela de Histórico ---
        hist_window = tk.Toplevel(self); hist_window.title(f"Histórico Consolidado"); hist_window.geometry("1150x500"); hist_window.transient(self); hist_window.grab_set()
        hist_frame = ttk.Frame(hist_window, padding="10"); hist_frame.pack(fill="both", expand=True)
        hist_tree = ttk.Treeview(hist_frame, columns=("Data", "Equipamento", "Tipo", "Qtd", "Responsável", "Solicitante", "Chamado", "Destino/Origem", "Motivo/Laudo"), show="headings")
        
        # Configuração dos cabeçalhos e colunas da tabela de histórico.
        hist_tree.heading("Data", text="Data e Hora"); 
        hist_tree.heading("Equipamento", text="Equipamento"); 
        hist_tree.heading("Tipo", text="Tipo"); hist_tree.heading("Qtd", text="Qtd"); 
        hist_tree.heading("Responsável", text="Responsável"); 
        hist_tree.heading("Solicitante", text="Solicitante"); 
        hist_tree.heading("Chamado", text="Chamado"); 
        hist_tree.heading("Destino/Origem", text="Destino/Origem"); 
        hist_tree.heading("Motivo/Laudo", text="Motivo/Laudo")
        hist_tree.column("Data", width=130); hist_tree.column("Equipamento", width=150); 
        hist_tree.column("Tipo", width=60, anchor="center"); 
        hist_tree.column("Qtd", width=50, anchor="center"); 
        hist_tree.column("Responsável", width=120); 
        hist_tree.column("Solicitante", width=120); 
        hist_tree.column("Chamado", width=80); 
        hist_tree.column("Destino/Origem", width=150); 
        hist_tree.column("Motivo/Laudo", width=200)
        hist_tree.pack(fill="both", expand=True)
        
        # --- Busca e Exibição dos Dados ---
        if not self.mov_df.empty and 'id_equipamento_fk' in self.mov_df.columns:
            
            # Filtra o DataFrame de movimentações para pegar apenas os registros dos IDs selecionados.
            historico_df = self.mov_df[self.mov_df['id_equipamento_fk'].isin(ids_para_buscar)]
            
            # Cria um mapeamento de ID para nome. Isso permite mostrar o nome de um item
            # mesmo que ele já tenha sido excluído da planilha principal de equipamentos.
            id_to_name = self.equip_df.set_index('id')['nome'].to_dict()
            historico_df['nome_equipamento'] = historico_df['id_equipamento_fk'].map(id_to_name).fillna("Item Excluído")
            
            # Ordena para mostrar as movimentações mais recentes primeiro.
            historico_df = historico_df.sort_values(by="id_movimentacao", ascending=False)
            
            # Garante que a coluna 'motivo_laudo' exista e preenche valores nulos para evitar erros.
            if 'motivo_laudo' not in historico_df.columns: historico_df['motivo_laudo'] = ''
            historico_df['motivo_laudo'] = historico_df['motivo_laudo'].fillna('')
            
            df_display = historico_df[['data_movimentacao', 'nome_equipamento', 'tipo_movimentacao', 'quantidade_movida', 'responsavel_movimentacao', 'solicitante', 'chamado', 'destino_origem', 'motivo_laudo']]
            
            # Itera sobre o DataFrame filtrado e insere cada linha na tabela da janela.
            for index, row in df_display.iterrows(): hist_tree.insert("", "end", values=list(row))

    def abrir_janela_edicao(self):
        
        """Abre uma nova janela para editar os detalhes de UM item selecionado."""
        
        # --- Validação da Seleção ---
        selected_items = self.tree.selection()
        if len(selected_items) != 1: 
            messagebox.showwarning("Seleção Inválida", "Por favor, selecione apenas UM equipamento para editar."); 
            return
        
        # --- Coleta dos Dados Atuais ---
        item_id = int(self.tree.item(selected_items[0], "values")[0])
        item_data = self.equip_df[self.equip_df['id'] == item_id].iloc[0]
        
        # --- Criação da Janela de Edição ---
        edit_window = tk.Toplevel(self); edit_window.title("Editar Equipamento"); edit_window.geometry("400x380"); edit_window.transient(self); edit_window.grab_set()
        edit_frame = ttk.Frame(edit_window, padding="20"); edit_frame.pack(fill="both", expand=True)
        
        # --- Preenchimento do Formulário ---
        # Cria os widgets e os pré-popula com os dados atuais do item.
        ttk.Label(edit_frame, text="Nome:").grid(row=0, column=0, padx=5, pady=8, sticky="w"); entry_edit_nome = ttk.Entry(edit_frame, width=35); entry_edit_nome.grid(row=0, column=1, padx=5, pady=8); entry_edit_nome.insert(0, item_data['nome'])
        ttk.Label(edit_frame, text="Categoria:").grid(row=1, column=0, padx=5, pady=8, sticky="w")
        combo_edit_categoria = ttk.Combobox(edit_frame, values=self.lista_categorias, width=33, state="readonly"); combo_edit_categoria.grid(row=1, column=1, padx=5, pady=8)
        if 'categoria' in item_data and pd.notna(item_data['categoria']): combo_edit_categoria.set(item_data['categoria'])
        ttk.Label(edit_frame, text="Nº de Série/SKU:").grid(row=2, column=0, padx=5, pady=8, sticky="w"); entry_edit_serie = ttk.Entry(edit_frame, width=35); entry_edit_serie.grid(row=2, column=1, padx=5, pady=8); entry_edit_serie.insert(0, item_data['numero_serie'])
        ttk.Label(edit_frame, text="Descrição:").grid(row=3, column=0, padx=5, pady=8, sticky="w"); entry_edit_descricao = ttk.Entry(edit_frame, width=35); entry_edit_descricao.grid(row=3, column=1, padx=5, pady=8); entry_edit_descricao.insert(0, item_data['descricao'])
        ttk.Label(edit_frame, text="Quantidade Atual:").grid(row=4, column=0, padx=5, pady=8, sticky="w"); entry_edit_quantidade = ttk.Entry(edit_frame, width=35); entry_edit_quantidade.grid(row=4, column=1, padx=5, pady=8); entry_edit_quantidade.insert(0, item_data['quantidade'])
        ttk.Label(edit_frame, text="Estoque Mínimo:").grid(row=5, column=0, padx=5, pady=8, sticky="w"); entry_edit_estoque_minimo = ttk.Entry(edit_frame, width=35); entry_edit_estoque_minimo.grid(row=5, column=1, padx=5, pady=8); entry_edit_estoque_minimo.insert(0, item_data['estoque_minimo'])
        
        # --- Botões de Ação (Salvar e Cancelar) ---
        button_frame = ttk.Frame(edit_frame); button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        # O 'lambda' captura os valores dos campos no momento do clique e os passa para a função 'salvar_edicao'.
        btn_salvar = ttk.Button(button_frame, text="Salvar Alterações", command=lambda: self.salvar_edicao(
            item_id, entry_edit_nome.get(), combo_edit_categoria.get(), entry_edit_serie.get(), 
            entry_edit_descricao.get(), entry_edit_quantidade.get(), 
            entry_edit_estoque_minimo.get(), edit_window))
        btn_salvar.pack(side="left", padx=10)
        btn_cancelar = ttk.Button(button_frame, text="Cancelar", command=edit_window.destroy); btn_cancelar.pack(side="left", padx=10)

    def get_data_from_tree_selection(self, selected_item_ids):
        
        """
        Função auxiliar que converte a seleção da Treeview em uma lista de dicionários com os dados completos dos itens.
        """
        
        if not selected_item_ids: return []
        
        # Extrai os IDs numéricos (da coluna 0) dos itens selecionados.
        tree_ids = [int(self.tree.item(item_id, "values")[0]) for item_id in selected_item_ids]
        
        # Filtra o DataFrame principal e retorna as linhas correspondentes como uma lista de dicionários.
        selected_data = self.equip_df[self.equip_df['id'].isin(tree_ids)].to_dict('records')
        return selected_data
    
    def abrir_janela_relatorio_opcoes(self):
        
        """Cria uma janela modal para o usuário configurar as opções do relatório de estoque."""
        
        opts_window = tk.Toplevel(self)
        opts_window.title("Opções do Relatório")
        opts_window.geometry("400x280")
        opts_window.resizable(False, False)
        opts_window.transient(self) # Faz a janela aparecer sobre a principal.
        opts_window.grab_set() # Bloqueia a interação com a janela principal.

        opts_frame = ttk.Frame(opts_window, padding="20")
        opts_frame.pack(expand=True, fill="both")

        # --- Lógica para Simular Placeholder nos Campos de Data ---
        # Melhora a experiência do usuário, mostrando o formato esperado (dd/mm/aaaa).
        placeholder_text = "dd/mm/aaaa"
        placeholder_color = "grey"
        default_fg_color = self.entry_nome.cget("foreground")

        def on_focus_in(event):
            
            """Quando o campo de data ganha foco, limpa o texto do placeholder."""
            
            widget = event.widget
            if widget.get() == placeholder_text:
                widget.delete(0, tk.END)
                widget.config(foreground=default_fg_color)

        def on_focus_out(event):
            
            """Quando o campo de data perde o foco, se estiver vazio, reinsere o placeholder."""
            
            widget = event.widget
            if not widget.get():
                widget.insert(0, placeholder_text)
                widget.config(foreground=placeholder_color)

        opts_frame.grid_columnconfigure(0, weight=1) # Faz a coluna expandir e centralizar

        # Checkbox principal
        incluir_historico_var = tk.BooleanVar()
        check_historico = ttk.Checkbutton(opts_frame, text="Incluir histórico de movimentações?", variable=incluir_historico_var)
        check_historico.grid(row=0, column=0, pady=(0,10))

        # Frame para as opções de data (que aparece e desaparece)
        date_filter_frame = ttk.Frame(opts_frame)
        date_filter_frame.grid(row=1, column=0, pady=(0, 10))
        
        filtro_data_var = tk.StringVar(value="todos")
        rb_todos = ttk.Radiobutton(date_filter_frame, text="Todo o período", variable=filtro_data_var, value="todos")
        rb_todos.pack(anchor="w")
        rb_intervalo = ttk.Radiobutton(date_filter_frame, text="Intervalo de datas específico:", variable=filtro_data_var, value="intervalo")
        rb_intervalo.pack(anchor="w", pady=(5,0))

        # Campos de data
        datas_frame = ttk.Frame(date_filter_frame)
        datas_frame.pack(pady=5, padx=20)
        
        ttk.Label(datas_frame, text="De:").pack(side="left")
        entry_inicio = ttk.Entry(datas_frame, width=15, justify="center")
        entry_inicio.pack(side="left", padx=5)
        
        ttk.Label(datas_frame, text="Até:").pack(side="left", padx=(10,0))
        entry_fim = ttk.Entry(datas_frame, width=15, justify="center")
        entry_fim.pack(side="left", padx=5)

        # Configurando o placeholder para os campos de data
        for entry in [entry_inicio, entry_fim]:
            entry.insert(0, placeholder_text)
            entry.config(foreground=placeholder_color)
            entry.bind("<FocusIn>", on_focus_in)
            entry.bind("<FocusOut>", on_focus_out)

        def toggle_date_filter_visibility(*args):
            if incluir_historico_var.get():
                date_filter_frame.grid()
            else:
                date_filter_frame.grid_remove()

        def toggle_date_entries_visibility(*args):
            if filtro_data_var.get() == "intervalo":
                datas_frame.pack(pady=5, padx=20)
            else:
                datas_frame.pack_forget()

        incluir_historico_var.trace("w", toggle_date_filter_visibility)
        filtro_data_var.trace("w", toggle_date_entries_visibility)
        
        toggle_date_filter_visibility()
        toggle_date_entries_visibility()

        def on_gerar_click():
            # Pega todos os valores primeiro
            incluir_hist = incluir_historico_var.get()
            filtro_data = filtro_data_var.get()
            
            data_inicio = entry_inicio.get()
            if data_inicio == placeholder_text:
                data_inicio = ""

            data_fim = entry_fim.get()
            if data_fim == placeholder_text:
                data_fim = ""
            
            # Depois, fecha a janela
            opts_window.destroy()
            
            # Finalmente, chama a função de gerar o relatório com os valores salvos
            self.gerar_relatorio(incluir_hist, filtro_data, data_inicio, data_fim)

        btn_gerar = ttk.Button(opts_frame, text="Gerar Relatório", command=on_gerar_click)
        btn_gerar.grid(row=2, column=0, pady=20)

    def gerar_relatorio(self, incluir_historico, filtro_data, data_inicio_str, data_fim_str):
        
        """
        Gera um relatório em HTML com base nos itens visíveis na tabela de estoque.
        
        Args:
            incluir_historico (bool): Se True, inclui o histórico de movimentações para cada item.
            filtro_data (str): 'todos' ou 'intervalo', define como filtrar o histórico por data.
            data_inicio_str (str): Data de início do filtro (formato dd/mm/aaaa).
            data_fim_str (str): Data de fim do filtro (formato dd/mm/aaaa).
        """
        # Pega os IDs dos itens que estão atualmente visíveis na tabela (Treeview).
        visible_items_tree_ids = self.tree.get_children()
        if not visible_items_tree_ids:
            messagebox.showinfo("Relatório Vazio", "Não há itens na lista para gerar um relatório."); return

        # Filtra o DataFrame principal para conter apenas os itens visíveis.
        visible_ids = [int(self.tree.item(item_id, "values")[0]) for item_id in visible_items_tree_ids]
        df_relatorio = self.equip_df[self.equip_df['id'].isin(visible_ids)]

        # Abre uma janela para o usuário escolher onde salvar o arquivo.
        filepath = filedialog.asksaveasfilename(
            defaultextension=".html", filetypes=[("Arquivos HTML", "*.html"), ("Todos os arquivos", "*.*")],
            title="Salvar Relatório Como...", initialfile=f"Relatorio_Estoque_{datetime.datetime.now().strftime('%Y-%m-%d')}.html"
        )
        if not filepath: return # Se o usuário cancelar, a função termina.

        try:
            # Tenta carregar um template HTML customizado.
            try:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                template_path = os.path.join(script_dir, 'relatorio_template.html')
                with open(template_path, 'r', encoding='utf-8') as f: template_html = f.read()
            except FileNotFoundError:
                
                # Se o template não for encontrado, usa um HTML básico como fallback.
                messagebox.showinfo("Template não encontrado", "Arquivo 'relatorio_template.html' não encontrado. Usando layout básico.")
                template_html = """
                <!DOCTYPE html><html lang="pt-br"><head><meta charset="UTF-8"><title>Relatório de Estoque</title>
                <style>body{font-family:sans-serif;} h1{color:#007bff;} .header{border-bottom:1px solid #ccc;padding-bottom:10px;} table{border-collapse:collapse;width:100%;margin-top:15px;} th,td{border:1px solid #ddd;padding:8px;} th{background-color:#f2f2f2;}</style>
                </head><body><div class="header"><h1>Relatório de Estoque</h1><p><strong>Gerado em:</strong> {{data_geracao}}</p><p><strong>Total de Tipos de Itens no Relatório:</strong> {{total_itens}}</p></div><hr>{{conteudo_relatorio}}</body></html>
                """

            # --- Processamento dos Filtros de Data ---
            data_inicio_dt, data_fim_dt = None, None
            if incluir_historico and filtro_data == 'intervalo':
                try:
                    data_inicio_dt = pd.to_datetime(data_inicio_str, format='%d/%m/%Y')
                    
                    # Adiciona 1 dia e subtrai 1 segundo para incluir o dia final inteiro no intervalo.
                    data_fim_dt = pd.to_datetime(data_fim_str, format='%d/%m/%Y') + pd.Timedelta(days=1, seconds=-1)
                except ValueError:
                    messagebox.showerror("Data Inválida", "Formato de data inválido. Use dd/mm/aaaa."); return

            # --- Construção do Conteúdo HTML Dinâmico ---
            conteudo_dinamico = ""
            for index, item in df_relatorio.iterrows():
                # Adiciona as informações básicas do item.
                conteudo_dinamico += "<div class='item-section'>"
                conteudo_dinamico += f"<h2>{item['nome']} (ID: {item['id']})</h2>"
                conteudo_dinamico += "<div class='item-details-grid'>"
                conteudo_dinamico += f"<p><strong>Categoria:</strong> {item.get('categoria', 'N/A')}</p>"
                conteudo_dinamico += f"<p><strong>Status:</strong> {item['status']}</p>"
                conteudo_dinamico += f"<p><strong>Quantidade Atual:</strong> {item['quantidade']}</p>"
                conteudo_dinamico += f"<p><strong>Estoque Mínimo:</strong> {item['estoque_minimo']}</p>"
                conteudo_dinamico += f"<p><strong>Nº de Série/SKU:</strong> {item.get('numero_serie', '')}</p>"
                conteudo_dinamico += f"<p><strong>Descrição:</strong> {item.get('descricao', '')}</p>"
                conteudo_dinamico += "</div>"

                # Se a opção foi marcada, busca e adiciona o histórico do item.
                if incluir_historico:
                    hist_df = self.mov_df[self.mov_df['id_equipamento_fk'] == item['id']]
                    
                    # Aplica o filtro de data, se necessário.
                    if filtro_data == 'intervalo' and not hist_df.empty:
                        hist_df = hist_df[(hist_df['data_movimentacao_dt'] >= data_inicio_dt) & (hist_df['data_movimentacao_dt'] <= data_fim_dt)]

                     # Se houver histórico, cria uma tabela HTML para ele.
                    if not hist_df.empty:
                        hist_df_sorted = hist_df.sort_values(by='id_movimentacao', ascending=False)
                        conteudo_dinamico += "<h3>Histórico de Movimentações:</h3>"
                        
                                        
                        conteudo_dinamico += "<table><thead><tr><th>Data</th><th>Tipo</th><th>Qtd</th><th>Responsável</th><th>Solicitante</th><th>Destino/Origem</th><th>Chamado</th><th>Motivo/Laudo</th></tr></thead><tbody>"
                        for h_index, mov in hist_df_sorted.iterrows():
                            conteudo_dinamico += "<tr>"
                            conteudo_dinamico += f"<td>{mov.get('data_movimentacao', '')}</td>"
                            conteudo_dinamico += f"<td>{mov.get('tipo_movimentacao', '')}</td>"
                            conteudo_dinamico += f"<td>{mov.get('quantidade_movida', '')}</td>"
                            conteudo_dinamico += f"<td>{mov.get('responsavel_movimentacao', '')}</td>"
                            conteudo_dinamico += f"<td>{mov.get('solicitante', '')}</td>"
                            conteudo_dinamico += f"<td>{mov.get('destino_origem', '')}</td>"
                            conteudo_dinamico += f"<td>{mov.get('chamado', '')}</td>"
                            conteudo_dinamico += f"<td>{mov.get('motivo_laudo', '')}</td>"
                            conteudo_dinamico += "</tr>"
                        conteudo_dinamico += "</tbody></table>"
                
                conteudo_dinamico += "</div>" # Fim do .item-section
            
            # --- Montagem do HTML Final ---
            # Substitui os placeholders (ex: {{data_geracao}}) no template pelos valores reais.
            final_html = template_html.replace('{{data_geracao}}', datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
            final_html = final_html.replace('{{total_itens}}', str(len(df_relatorio)))
            final_html = final_html.replace('{{conteudo_relatorio}}', conteudo_dinamico)
            
            # Escreve o HTML final no arquivo escolhido pelo usuário.
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(final_html)
            
            messagebox.showinfo("Sucesso", f"Relatório salvo com sucesso em:\n{filepath}")
            
            # Pergunta se o usuário deseja abrir o arquivo gerado no navegador.
            if messagebox.askyesno("Abrir Relatório", "Deseja abrir o relatório agora no seu navegador?"):
                webbrowser.open(f"file://{os.path.realpath(filepath)}")

        except Exception as e:
            messagebox.showerror("Erro ao Gerar Relatório", f"Não foi possível gerar o arquivo.\n\nErro: {e}")
            
    def gerar_relatorio_setores(self, filtro_status, filtro_data, data_inicio_str, data_fim_str):
        
        """
        Gera um relatório HTML das movimentações entre setores, aplicando filtros de status e data.

        Args:
            filtro_status (str): O critério de filtro para o status ('todos', 'pendentes', 'regularizados').
            filtro_data (str): O critério de filtro para a data ('todos', 'intervalo').
            data_inicio_str (str): A data de início para o filtro de intervalo ('dd/mm/aaaa').
            data_fim_str (str): A data de fim para o filtro de intervalo ('dd/mm/aaaa').
        """

        # --- Validação Inicial ---
        if not hasattr(self, 'mov_setores_df') or self.mov_setores_df.empty:
            messagebox.showinfo("Relatório Vazio", "Não há movimentações para gerar um relatório.")
            return

        df_relatorio = self.mov_setores_df.copy()

        # --- Filtro por Status ---
        # Garante que a coluna de status exista e preenche valores vazios com 'Pendente'.
        if 'status_regularizacao' not in df_relatorio.columns:
            df_relatorio['status_regularizacao'] = 'Pendente'
        df_relatorio['status_regularizacao'] = df_relatorio['status_regularizacao'].fillna('Pendente') 
        
        # Aplica o filtro com base na escolha do usuário.
        if filtro_status == 'pendentes':
            df_relatorio = df_relatorio[df_relatorio['status_regularizacao'] == 'Pendente']
        elif filtro_status == 'regularizados':
            df_relatorio = df_relatorio[df_relatorio['status_regularizacao'] == 'Regularizado']
        
        # --- Filtro por Data ---
        # Aplica o filtro de data no DataFrame que já pode ter sido filtrado por status.
        if filtro_data == 'intervalo':
            try:
                if 'data_movimentacao_dt' not in df_relatorio.columns:
                     df_relatorio['data_movimentacao_dt'] = pd.to_datetime(df_relatorio['data_movimentacao'], format='%d-%m-%Y %H:%M:%S', errors='coerce')
                
                # Converte as strings de data para o formato datetime para permitir a comparação.
                data_inicio_dt = pd.to_datetime(data_inicio_str, format='%d/%m/%Y')
                data_fim_dt = pd.to_datetime(data_fim_str, format='%d/%m/%Y') + pd.Timedelta(days=1, seconds=-1)
                
                df_relatorio = df_relatorio[
                    (df_relatorio['data_movimentacao_dt'] >= data_inicio_dt) & 
                    (df_relatorio['data_movimentacao_dt'] <= data_fim_dt)
                ]
            except (ValueError, TypeError):
                messagebox.showerror("Data Inválida", "Formato de data inválido. Use dd/mm/aaaa.")
                return
            except Exception as e:
                messagebox.showerror("Erro no Filtro", f"Ocorreu um erro ao filtrar as datas: {e}")
                return

        if df_relatorio.empty:
            messagebox.showinfo("Relatório Vazio", "Nenhuma movimentação encontrada para os filtros selecionados.")
            return

        # --- Preparação para Geração do HTML ---
        filepath = filedialog.asksaveasfilename(
            defaultextension=".html", filetypes=[("Arquivos HTML", "*.html"), ("Todos os arquivos", "*.*")],
            title="Salvar Relatório de Movimentação Entre Setores",
            initialfile=f"Relatorio_Mov_Setores_{datetime.datetime.now().strftime('%Y-%m-%d')}.html"
        )
        if not filepath: return

        try:
            # --- Carregamento do Template e Preparação dos Dados ---
            template_path = self.resource_path('relatorio_setores_template.html')
            with open(template_path, 'r', encoding='utf-8') as f: template_html = f.read()
            
            # Garante que todas as colunas necessárias existam no DataFrame para evitar erros.
            for col in ['chamado', 'solicitante', 'status_regularizacao']:
                if col not in df_relatorio.columns: df_relatorio[col] = ''
            df_relatorio = df_relatorio.fillna('')

            # Renomeia as colunas para nomes mais amigáveis no relatório.
            colunas_para_exibir = {'data_movimentacao': 'Data', 'tipo_equipamento': 'Equipamento', 'patrimonio': 'Patrimônio', 'servicetag': 'ServiceTag', 'setor_origem': 'Origem', 'setor_destino': 'Destino', 'responsavel': 'Responsável', 'chamado': 'Chamado', 'solicitante': 'Solicitante', 'status_regularizacao': 'Status', 'observacao': 'Observação'}
            df_relatorio_display = df_relatorio.rename(columns=colunas_para_exibir)
            df_relatorio_display = df_relatorio_display[list(colunas_para_exibir.values())] # Garante a ordem
            
            #Converte quebras de linha (\n) em tags HTML (<br>) para exibição correta dos dados do "Kit".
            df_relatorio_display['Patrimônio'] = df_relatorio_display['Patrimônio'].str.replace('\n', '<br>')
            df_relatorio_display['ServiceTag'] = df_relatorio_display['ServiceTag'].str.replace('\n', '<br>')

            # --- Geração da Tabela HTML ---
            # Usa a função to_html do Pandas para converter o DataFrame em uma tabela HTML.
            tabela_html = df_relatorio_display.to_html(index=False, justify='left', border=0, classes="styled-table", escape=False)
            
            # --- Montagem e Salvamento do Arquivo Final ---
            final_html = template_html.replace('{{data_geracao}}', datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
            final_html = final_html.replace('{{tabela_movimentacoes}}', tabela_html)
            
            with open(filepath, 'w', encoding='utf-8') as f: f.write(final_html)
            
            messagebox.showinfo("Sucesso", f"Relatório salvo com sucesso em:\n{filepath}")
            if messagebox.askyesno("Abrir Relatório", "Deseja abrir o relatório agora no seu navegador?"):
                webbrowser.open(f"file://{os.path.realpath(filepath)}")
        except Exception as e:
            messagebox.showerror("Erro ao Gerar Relatório", f"Não foi possível gerar o arquivo.\n\nErro: {e}")

            # --- FUNÇÃO NOVA PARA AS OPÇÕES DO RELATÓRIO DE SETORES ---

    def abrir_janela_relatorio_setores_opcoes(self):
        
        """Abre uma janela de opções para o relatório de movimentação entre setores."""
        
        opts_window = tk.Toplevel(self)
        opts_window.title("Opções do Relatório de Setores")
        opts_window.geometry("400x320") # Altura ajustada para os novos filtros.
        opts_window.resizable(False, False)
        opts_window.transient(self)
        opts_window.grab_set()

        opts_frame = ttk.Frame(opts_window, padding="20")
        opts_frame.pack(expand=True, fill="both")
        opts_frame.grid_columnconfigure(0, weight=1)

        # --- Filtro de Status ---
        status_filter_frame = ttk.LabelFrame(opts_frame, text="Filtrar por Status", padding=10)
        status_filter_frame.grid(row=0, column=0, pady=(0, 15), sticky="ew")
        filtro_status_var = tk.StringVar(value="todos")
        ttk.Radiobutton(status_filter_frame, text="Todos", variable=filtro_status_var, value="todos").pack(anchor="w")
        ttk.Radiobutton(status_filter_frame, text="Somente Pendentes (Exceto Regularizados)", variable=filtro_status_var, value="pendentes").pack(anchor="w")
        ttk.Radiobutton(status_filter_frame, text="Somente Regularizados", variable=filtro_status_var, value="regularizados").pack(anchor="w")
        
        # --- Filtro de Data ---
        date_filter_frame = ttk.LabelFrame(opts_frame, text="Filtrar por Data", padding=10)
        date_filter_frame.grid(row=1, column=0, sticky="ew")

        # Lógica de placeholder para os campos de data.
        placeholder_text = "dd/mm/aaaa"; placeholder_color = "grey"; default_fg_color = self.entry_nome.cget("foreground")
        def on_focus_in(event):
            widget = event.widget
            if widget.get() == placeholder_text: widget.delete(0, tk.END); widget.config(foreground=default_fg_color)
        def on_focus_out(event):
            widget = event.widget
            if not widget.get(): widget.insert(0, placeholder_text); widget.config(foreground=placeholder_color)

        filtro_data_var = tk.StringVar(value="todos")
        rb_todos = ttk.Radiobutton(date_filter_frame, text="Todo o período", variable=filtro_data_var, value="todos")
        rb_todos.pack(anchor="w")
        rb_intervalo = ttk.Radiobutton(date_filter_frame, text="Intervalo de datas específico:", variable=filtro_data_var, value="intervalo")
        rb_intervalo.pack(anchor="w", pady=(5,0))
        datas_frame = ttk.Frame(date_filter_frame)
        datas_frame.pack(pady=5, padx=20)
        
        ttk.Label(datas_frame, text="De:").pack(side="left"); entry_inicio = ttk.Entry(datas_frame, width=15, justify="center"); entry_inicio.pack(side="left", padx=5)
        ttk.Label(datas_frame, text="Até:").pack(side="left", padx=(10,0)); entry_fim = ttk.Entry(datas_frame, width=15, justify="center"); entry_fim.pack(side="left", padx=5)

        for entry in [entry_inicio, entry_fim]:
            entry.insert(0, placeholder_text); entry.config(foreground=placeholder_color); entry.bind("<FocusIn>", on_focus_in); entry.bind("<FocusOut>", on_focus_out)

        # --- Lógica de Visibilidade Dinâmica ---
        def toggle_date_entries_visibility(*args):
            
            """Mostra ou esconde os campos de entrada de data."""
            
            if filtro_data_var.get() == "intervalo": datas_frame.pack(pady=5, padx=20)
            else: datas_frame.pack_forget()
        
        # Associa a função de toggle à variável do radio button.
        filtro_data_var.trace("w", toggle_date_entries_visibility); toggle_date_entries_visibility()

        def on_gerar_click():
            
            """Coleta todas as opções selecionadas e chama a função de gerar relatório."""
            
            filtro_status = filtro_status_var.get()
            filtro_data = filtro_data_var.get()
            data_inicio = entry_inicio.get()
            if data_inicio == placeholder_text: data_inicio = ""
            data_fim = entry_fim.get()
            if data_fim == placeholder_text: data_fim = ""
            opts_window.destroy()
            self.gerar_relatorio_setores(filtro_status, filtro_data, data_inicio, data_fim)

        btn_gerar = ttk.Button(opts_frame, text="Gerar Relatório", command=on_gerar_click)
        btn_gerar.grid(row=2, column=0, pady=20)
        
    def marcar_como_regularizado(self):
        
        """Marca uma ou mais movimentações selecionadas como 'Regularizado' na planilha."""
        
        selected_items_ids = self.mov_setores_tree.selection()
        if not selected_items_ids:
            messagebox.showwarning("Nenhum Item Selecionado", "Selecione uma ou mais movimentações pendentes para marcar como regularizadas.")
            return

        updates = [] # Lista para armazenar as atualizações a serem enviadas em lote.
        
        for item_id_widget in selected_items_ids:
            item_values = self.mov_setores_tree.item(item_id_widget, "values")
            mov_id = int(item_values[0])

            row_index = self._find_sheet_row_index_by_id(self.mov_setores_df, mov_id)
            if row_index:
                # Monta um dicionário de atualização para a célula correta.
                # 'L' é a 12ª letra, correspondente à coluna 'status_regularizacao'.
                updates.append({'range': f'L{row_index}', 'values': [['Regularizado']]})

        if not updates:
            messagebox.showerror("Erro", "Não foi possível encontrar os registros selecionados para atualizar.")
            return

        # Envia todas as atualizações de uma só vez para a API do Google Sheets.
        # Isso é muito mais rápido e eficiente do que atualizar uma célula de cada vez.
        self.mov_setores_sheet.batch_update(updates)
        messagebox.showinfo("Sucesso", f"{len(updates)} movimentação(ões) marcada(s) como 'Regularizado'.")
        self.refresh_all_data()

if __name__ == "__main__":
    app = App()
    app.mainloop()