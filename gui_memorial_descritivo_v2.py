#!/usr/bin/env python3
"""
Interface Gráfica para Processador de Memorial Descritivo - Versão 2.0

Funcionalidades:
- Modo Normal: Drag & drop de PDF
- Modo INCRA: Busca automática por prenotação
- Escolha de saída: Excel, Word ou Ambos
- Multi-thread: Interface não trava

Requisitos:
- pip install tkinterdnd2 google-generativeai openpyxl python-docx pillow pdf2image --break-system-packages
"""

import os
import sys
import json
import math
import shutil
import threading
import subprocess
import platform
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# Importa funções do script principal
try:
    from process_memorial_descritivo_v2 import (
        formatar_prenotacao,
        calcular_pasta_milhar,
        buscar_arquivo_incra,
        copiar_para_downloads,
        converter_tiff_para_pdf,
        extrair_memorial_incra,
        extract_table_from_pdf,
        create_excel_file,
        create_word_file,
        testar_acesso_rede,
        INCRA_CONFIG
    )
except ImportError:
    print("❌ Erro: process_memorial_descritivo_v2.py não encontrado!")
    print("Certifique-se de que o arquivo está no mesmo diretório.")
    sys.exit(1)


class MemorialGUI_V2:
    """Interface gráfica v2.0 para processamento de Memoriais Descritivos"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Memorial Descritivo INCRA")
        self.root.geometry("1100x800")
        self.root.resizable(True, True)
        self.root.configure(bg='#FFFFFF')
        
        # Variáveis
        self.pdf_path = StringVar()
        self.prenotacao = StringVar()
        self.api_key = StringVar()
        self.modo_operacao = StringVar(value="normal")  # "normal" ou "incra"
        self.submodo_incra = StringVar(value="paginas")  # "ia" ou "paginas"
        self.paginas_incra = StringVar()  # Páginas do documento INCRA
        self.status_text = StringVar(value="Aguardando...")
        self.progress_value = IntVar(value=0)
        self.processing = False
        self.table_data = None

        # Checkboxes para saída
        self.gerar_excel = BooleanVar(value=False)
        self.gerar_word = BooleanVar(value=False)

        # API Key fixa
        self.api_key.set('AIzaSyAdA_GO7cQ0m1ouie4wGwXf4a4SnHKjBh8')

        # Caminhos dos arquivos gerados
        self.excel_gerado = None
        self.word_gerado = None

        # Caminhos dos arquivos temporários para deletar
        self.arquivos_temporarios = []
        
        # Configurar estilo
        self.setup_style()
        
        # Criar interface
        self.create_widgets()
        
        # Configurar drag & drop (apenas para modo normal)
        self.setup_drag_drop()

        # Inicializar visual dos botões de submodo
        self.update_submodo_buttons()
    
    def setup_style(self):
        """Configura estilo moderno e minimalista"""
        style = ttk.Style()
        style.theme_use('clam')

        # Paleta moderna inspirada em VSCode e Slack
        self.colors = {
            'primary': '#5865F2',      # Azul vibrante (Discord-style)
            'primary_hover': '#4752C4',
            'success': '#3BA55D',      # Verde moderno
            'bg': '#FFFFFF',           # Branco puro
            'bg_secondary': '#F7F8FA', # Cinza muito claro
            'text': '#2E3338',         # Cinza escuro (quase preto)
            'text_secondary': '#5E6C84', # Cinza médio
            'border': '#E3E5E8',       # Borda sutil
            'card_bg': '#FFFFFF',
            'sidebar': '#F7F8FA'
        }

        # Fonte moderna
        font_family = 'Inter' if 'Inter' in self.root.tk.call('font', 'families') else 'Segoe UI'

        style.configure('Title.TLabel',
                       font=(font_family, 22, 'bold'),
                       foreground=self.colors['text'],
                       background=self.colors['bg'])

        style.configure('Subtitle.TLabel',
                       font=(font_family, 11),
                       foreground=self.colors['text_secondary'],
                       background=self.colors['bg'])

        style.configure('CardTitle.TLabel',
                       font=(font_family, 14, 'bold'),
                       foreground=self.colors['text'],
                       background=self.colors['card_bg'])

        style.configure('Body.TLabel',
                       font=(font_family, 11),
                       foreground=self.colors['text'],
                       background=self.colors['card_bg'])

        style.configure('Primary.TButton',
                       font=(font_family, 12, 'bold'),
                       background=self.colors['primary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(30, 15))

        style.map('Primary.TButton',
                 background=[('active', self.colors['primary_hover'])])

        style.configure('Success.TButton',
                       font=(font_family, 11),
                       background=self.colors['success'],
                       foreground='white',
                       borderwidth=0,
                       padding=(20, 10))

        style.map('Success.TButton',
                 background=[('active', self.colors['success'])],
                 foreground=[('active', 'white')])

        style.configure('Secondary.TButton',
                       font=(font_family, 11),
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text'],
                       borderwidth=1,
                       padding=(20, 10))

        style.configure('TRadiobutton',
                       font=(font_family, 12),
                       background=self.colors['card_bg'])

        style.configure('TCheckbutton',
                       font=(font_family, 13, 'bold'),
                       background=self.colors['card_bg'],
                       foreground=self.colors['text'])
        
    def create_widgets(self):
        """Cria todos os widgets da interface - Design Profissional"""

        # Container principal com Canvas e Scrollbar
        container = Frame(self.root, bg=self.colors['bg'])
        container.pack(fill=BOTH, expand=True)

        # Canvas para permitir scroll
        canvas = Canvas(container, bg=self.colors['bg'], highlightthickness=0)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        # Scrollbar profissional
        scrollbar = ttk.Scrollbar(container, orient=VERTICAL, command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)

        # Configurar canvas
        canvas.configure(yscrollcommand=scrollbar.set)

        # Frame principal dentro do canvas
        main_frame = Frame(canvas, bg=self.colors['bg'], padx=40, pady=30)
        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor=NW)

        # Função para atualizar o scroll quando o conteúdo mudar
        def configure_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Ajustar largura do frame ao canvas
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())

        main_frame.bind('<Configure>', configure_scroll)
        canvas.bind('<Configure>', configure_scroll)

        # Scroll com a roda do mouse
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Armazenar referências
        self.canvas = canvas
        self.main_frame = main_frame

        # ===== CABEÇALHO PROFISSIONAL =====
        header_frame = Frame(main_frame, bg=self.colors['bg'])
        header_frame.pack(fill=X, pady=(0, 35))

        # Título com ícone
        title_label = Label(header_frame,
                           text="📋  Memorial Descritivo INCRA",
                           font=('Segoe UI', 28, 'bold'),
                           fg=self.colors['primary'],
                           bg=self.colors['bg'])
        title_label.pack(anchor=CENTER)

        # Subtítulo
        subtitle_label = Label(header_frame,
                              text="Sistema Profissional de Processamento Automatizado",
                              font=('Segoe UI', 13),
                              fg=self.colors['text_secondary'],
                              bg=self.colors['bg'])
        subtitle_label.pack(anchor=CENTER, pady=(8, 0))

        # Linha separadora
        separator1 = Frame(main_frame, height=2, bg=self.colors['border'])
        separator1.pack(fill=X, pady=25)
        
        # ===== SEÇÃO 1: MODO DE OPERAÇÃO =====
        modo_card = Frame(main_frame, bg=self.colors['card_bg'], relief=FLAT, bd=0)
        modo_card.pack(fill=X, pady=(0, 20))

        # Padding interno do card
        modo_inner = Frame(modo_card, bg=self.colors['card_bg'], padx=30, pady=25)
        modo_inner.pack(fill=BOTH, expand=True)

        # Título da seção
        modo_title = Label(modo_inner,
                          text="🎯  Modo de Operação",
                          font=('Segoe UI', 18, 'bold'),
                          fg=self.colors['primary'],
                          bg=self.colors['card_bg'])
        modo_title.pack(anchor=W, pady=(0, 8))

        # Subtítulo da seção
        modo_subtitle = Label(modo_inner,
                             text="Selecione como deseja processar o memorial descritivo:",
                             font=('Segoe UI', 12),
                             fg=self.colors['text_secondary'],
                             bg=self.colors['card_bg'])
        modo_subtitle.pack(anchor=W, pady=(0, 20))

        # Opções
        modo_normal_radio = ttk.Radiobutton(
            modo_inner,
            text="📄  Processar arquivo PDF já existente",
            variable=self.modo_operacao,
            value="normal",
            command=self.atualizar_modo
        )
        modo_normal_radio.pack(anchor=W, pady=10)

        modo_incra_radio = ttk.Radiobutton(
            modo_inner,
            text="🏛️  Buscar automaticamente por número de Prenotação INCRA",
            variable=self.modo_operacao,
            value="incra",
            command=self.atualizar_modo
        )
        modo_incra_radio.pack(anchor=W, pady=10)
        
        # ===== FRAMES DE ENTRADA (Normal e INCRA) =====

        # Container para trocar entre modos
        self.input_container = ttk.Frame(main_frame)
        self.input_container.pack(fill=BOTH, expand=True, pady=(0, 25))

        # Frame Modo Normal (PDF)
        self.normal_frame = ttk.LabelFrame(self.input_container,
                                          text="  📄 Selecione o arquivo PDF  ",
                                          padding="25")

        self.drop_frame = Frame(self.normal_frame, bg='#E8F4FF', relief=GROOVE, bd=3,
                               height=200)
        self.drop_frame.pack(fill=BOTH, expand=True, pady=(0, 15))

        drop_label = Label(self.drop_frame,
                          text="📂\n\nClique aqui para selecionar o arquivo PDF\n\nou arraste o arquivo para esta área",
                          bg='#E8F4FF', fg='#000000', font=('Arial', 16, 'bold'),
                          cursor='hand2')
        drop_label.pack(expand=True, pady=40)
        drop_label.bind('<Button-1>', lambda e: self.select_pdf())

        path_frame = ttk.Frame(self.normal_frame)
        path_frame.pack(fill=X)

        path_label = ttk.Label(path_frame, text="Arquivo selecionado:",
                              font=('Arial', 14, 'bold'))
        path_label.pack(anchor=W, pady=(0, 5))

        path_entry = ttk.Entry(path_frame, textvariable=self.pdf_path,
                              state='readonly', font=('Arial', 12))
        path_entry.pack(fill=X, pady=(2, 0), ipady=8)
        
        # Frame Modo INCRA (Prenotação)
        self.incra_frame = ttk.LabelFrame(self.input_container,
                                         text="  🏛️ Busca por Prenotação INCRA  ",
                                         padding="25")

        # ===== TOGGLE: IA vs Páginas =====
        toggle_frame = ttk.Frame(self.incra_frame)
        toggle_frame.pack(fill=X, pady=(0, 20))

        toggle_label = ttk.Label(toggle_frame,
                                text="Método de Detecção:",
                                font=('Arial', 14, 'bold'))
        toggle_label.pack(anchor=W, pady=(0, 10))

        # Container dos botões de toggle
        toggle_buttons = ttk.Frame(toggle_frame)
        toggle_buttons.pack(fill=X)

        # Botão "Por Páginas"
        self.btn_paginas = Frame(toggle_buttons, bg='#E3E5E8', relief=SOLID, bd=2, cursor='hand2')
        self.btn_paginas.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))

        paginas_inner = Frame(self.btn_paginas, bg='#E3E5E8', padx=15, pady=10)
        paginas_inner.pack(fill=BOTH, expand=True)

        paginas_icon = Label(paginas_inner, text="📄", font=('Segoe UI', 20), bg='#E3E5E8')
        paginas_icon.pack(side=LEFT, padx=(0, 10))

        paginas_label = Label(paginas_inner, text="Por Páginas",
                             font=('Segoe UI', 12, 'bold'),
                             fg=self.colors['text'], bg='#E3E5E8')
        paginas_label.pack(side=LEFT)

        # Botão "Por IA"
        self.btn_ia = Frame(toggle_buttons, bg='#E3E5E8', relief=SOLID, bd=2, cursor='hand2')
        self.btn_ia.pack(side=LEFT, fill=BOTH, expand=True, padx=(5, 0))

        ia_inner = Frame(self.btn_ia, bg='#E3E5E8', padx=15, pady=10)
        ia_inner.pack(fill=BOTH, expand=True)

        ia_icon = Label(ia_inner, text="🤖", font=('Segoe UI', 20), bg='#E3E5E8')
        ia_icon.pack(side=LEFT, padx=(0, 10))

        ia_label = Label(ia_inner, text="Por IA (Automático)",
                        font=('Segoe UI', 12, 'bold'),
                        fg=self.colors['text'], bg='#E3E5E8')
        ia_label.pack(side=LEFT)

        # Bind de cliques para toggle
        def toggle_paginas(event=None):
            self.submodo_incra.set("paginas")
            self.update_submodo_buttons()

        def toggle_ia(event=None):
            self.submodo_incra.set("ia")
            self.update_submodo_buttons()

        self.btn_paginas.bind('<Button-1>', toggle_paginas)
        paginas_inner.bind('<Button-1>', toggle_paginas)
        paginas_icon.bind('<Button-1>', toggle_paginas)
        paginas_label.bind('<Button-1>', toggle_paginas)

        self.btn_ia.bind('<Button-1>', toggle_ia)
        ia_inner.bind('<Button-1>', toggle_ia)
        ia_icon.bind('<Button-1>', toggle_ia)
        ia_label.bind('<Button-1>', toggle_ia)

        # Armazena referências
        self.paginas_widgets = [self.btn_paginas, paginas_inner, paginas_icon, paginas_label]
        self.ia_widgets = [self.btn_ia, ia_inner, ia_icon, ia_label]

        # Separador
        ttk.Separator(self.incra_frame, orient=HORIZONTAL).pack(fill=X, pady=15)

        # ===== PRENOTAÇÃO =====
        prenotacao_frame = ttk.Frame(self.incra_frame)
        prenotacao_frame.pack(fill=X, pady=(0, 15))

        prenotacao_label = ttk.Label(prenotacao_frame, text="Número da Prenotação:",
                                    font=('Arial', 16, 'bold'))
        prenotacao_label.pack(anchor=W, pady=(0, 10))

        prenotacao_entry = ttk.Entry(prenotacao_frame, textvariable=self.prenotacao,
                                     font=('Arial', 18), justify='center')
        prenotacao_entry.pack(fill=X, pady=(2, 0), ipady=15)

        # ===== CAMPOS DE PÁGINAS (visíveis apenas no modo "paginas") =====
        self.paginas_fields_frame = ttk.Frame(self.incra_frame)

        # Linha informativa
        info_label = ttk.Label(self.paginas_fields_frame,
                              text="📌 Digite os números das páginas separados por vírgula",
                              font=('Arial', 11),
                              foreground=self.colors['text_secondary'])
        info_label.pack(anchor=W, pady=(0, 10))

        # Campo de páginas
        paginas_input_frame = ttk.Frame(self.paginas_fields_frame)
        paginas_input_frame.pack(fill=X)

        paginas_input_label = ttk.Label(paginas_input_frame,
                                        text="Páginas do Documento INCRA:",
                                        font=('Arial', 14, 'bold'))
        paginas_input_label.pack(anchor=W, pady=(0, 8))

        paginas_entry = ttk.Entry(paginas_input_frame,
                                 textvariable=self.paginas_incra,
                                 font=('Arial', 14),
                                 justify='center')
        paginas_entry.pack(fill=X, ipady=8)

        # Exemplo
        exemplo_label = ttk.Label(paginas_input_frame,
                                 text="Exemplo: 1,2,4,7",
                                 font=('Arial', 10, 'italic'),
                                 foreground=self.colors['text_secondary'])
        exemplo_label.pack(anchor=W, pady=(5, 0))
        
        # Mostra frame inicial (Normal)
        self.normal_frame.pack(fill=BOTH, expand=True)
        
        # ===== ESCOLHA DE SAÍDA =====
        output_card = Frame(main_frame, bg=self.colors['card_bg'], relief=FLAT, bd=0)
        output_card.pack(fill=X, pady=(0, 25))

        output_inner = Frame(output_card, bg=self.colors['card_bg'], padx=30, pady=25)
        output_inner.pack(fill=BOTH, expand=True)

        # Título da seção
        output_title = Label(output_inner,
                            text="💾  Escolha os Arquivos a Gerar",
                            font=('Segoe UI', 18, 'bold'),
                            fg=self.colors['primary'],
                            bg=self.colors['card_bg'])
        output_title.pack(anchor=W, pady=(0, 8))

        # Subtítulo
        output_subtitle = Label(output_inner,
                               text="Selecione os formatos de saída desejados (você pode escolher ambos):",
                               font=('Segoe UI', 12),
                               fg=self.colors['text_secondary'],
                               bg=self.colors['card_bg'])
        output_subtitle.pack(anchor=W, pady=(0, 20))

        # Container dos botões
        buttons_container = Frame(output_inner, bg=self.colors['card_bg'])
        buttons_container.pack(fill=X)

        # Botão Excel
        self.excel_button = Frame(buttons_container, bg='#E3E5E8', relief=SOLID, bd=2, cursor='hand2')
        self.excel_button.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))

        excel_inner = Frame(self.excel_button, bg='#E3E5E8', padx=20, pady=15)
        excel_inner.pack(fill=BOTH, expand=True)

        excel_icon = Label(excel_inner, text="📊", font=('Segoe UI', 32), bg='#E3E5E8')
        excel_icon.pack(pady=(5, 10))

        excel_label = Label(excel_inner, text="Planilha Excel",
                           font=('Segoe UI', 14, 'bold'),
                           fg=self.colors['text'], bg='#E3E5E8')
        excel_label.pack()

        excel_desc = Label(excel_inner, text="Formato .xlsx",
                          font=('Segoe UI', 10),
                          fg=self.colors['text_secondary'], bg='#E3E5E8')
        excel_desc.pack(pady=(2, 5))

        # Botão Word
        self.word_button = Frame(buttons_container, bg='#E3E5E8', relief=SOLID, bd=2, cursor='hand2')
        self.word_button.pack(side=LEFT, fill=BOTH, expand=True, padx=(10, 0))

        word_inner = Frame(self.word_button, bg='#E3E5E8', padx=20, pady=15)
        word_inner.pack(fill=BOTH, expand=True)

        word_icon = Label(word_inner, text="📝", font=('Segoe UI', 32), bg='#E3E5E8')
        word_icon.pack(pady=(5, 10))

        word_label = Label(word_inner, text="Documento Word",
                          font=('Segoe UI', 14, 'bold'),
                          fg=self.colors['text'], bg='#E3E5E8')
        word_label.pack()

        word_desc = Label(word_inner, text="Formato .docx",
                         font=('Segoe UI', 10),
                         fg=self.colors['text_secondary'], bg='#E3E5E8')
        word_desc.pack(pady=(2, 5))

        # Bind de cliques
        def toggle_excel(event=None):
            self.gerar_excel.set(not self.gerar_excel.get())
            self.update_output_buttons()

        def toggle_word(event=None):
            self.gerar_word.set(not self.gerar_word.get())
            self.update_output_buttons()

        self.excel_button.bind('<Button-1>', toggle_excel)
        excel_inner.bind('<Button-1>', toggle_excel)
        excel_icon.bind('<Button-1>', toggle_excel)
        excel_label.bind('<Button-1>', toggle_excel)
        excel_desc.bind('<Button-1>', toggle_excel)

        self.word_button.bind('<Button-1>', toggle_word)
        word_inner.bind('<Button-1>', toggle_word)
        word_icon.bind('<Button-1>', toggle_word)
        word_label.bind('<Button-1>', toggle_word)
        word_desc.bind('<Button-1>', toggle_word)

        # Armazena referências para atualizar cores
        self.excel_widgets = [self.excel_button, excel_inner, excel_icon, excel_label, excel_desc]
        self.word_widgets = [self.word_button, word_inner, word_icon, word_label, word_desc]
        
        # ===== BOTÕES DE AÇÃO =====
        ttk.Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=20)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(10, 15))

        self.process_btn = ttk.Button(button_frame, text="✅ PROCESSAR AGORA",
                                      command=self.process_memorial,
                                      style='Primary.TButton')
        self.process_btn.pack(fill=X, pady=(0, 10), ipady=10)

        clear_btn = ttk.Button(button_frame, text="🗑️ Limpar Tudo",
                              command=self.clear_all,
                              style='Big.TButton')
        clear_btn.pack(fill=X, ipady=8)
        
        # ===== BARRA DE PROGRESSO =====
        progress_frame = ttk.LabelFrame(main_frame, text="  📊 Andamento  ",
                                       padding="20")
        progress_frame.pack(fill=X, pady=(0, 20))

        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate',
                                           variable=self.progress_value,
                                           length=400)
        self.progress_bar.pack(fill=X, pady=(0, 10), ipady=8)

        self.status_label = ttk.Label(progress_frame, textvariable=self.status_text,
                                      style='Status.TLabel')
        self.status_label.pack(anchor=CENTER)
        
        # ===== LOG =====
        log_frame = ttk.LabelFrame(main_frame, text="  📋 Mensagens do Sistema  ",
                                  padding="15")
        log_frame.pack(fill=BOTH, expand=True)

        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=RIGHT, fill=Y)

        self.log_text = Text(log_frame, height=6, font=('Courier New', 12, 'bold'),
                            yscrollcommand=log_scroll.set, wrap=WORD,
                            bg='#000000', fg='#00FF00', insertbackground='white')
        self.log_text.pack(fill=BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)

        self.log_text.tag_config('info', foreground='#00BFFF')
        self.log_text.tag_config('success', foreground='#00FF00')
        self.log_text.tag_config('error', foreground='#FF3333')
        self.log_text.tag_config('warning', foreground='#FFD700')
        self.log_text.tag_config('incra', foreground='#FF00FF')

        self.log("✅ Sistema pronto para uso!", 'success')
        self.log("👉 Escolha como deseja trabalhar acima", 'info')

        # ===== BOTÕES PARA ABRIR ARQUIVOS (inicialmente ocultos) =====
        self.results_frame = Frame(main_frame, bg=self.colors['bg'])
        # Não adiciona ao pack ainda - aparece após processamento

        results_title = Label(self.results_frame,
                            text="✨  Arquivos Gerados",
                            font=('Segoe UI', 18, 'bold'),
                            fg=self.colors['success'],
                            bg=self.colors['bg'])
        results_title.pack(anchor=W, pady=(0, 15))

        results_buttons = Frame(self.results_frame, bg=self.colors['bg'])
        results_buttons.pack(fill=X)

        self.btn_abrir_excel = ttk.Button(results_buttons,
                                         text="📊  Abrir Tabela Excel",
                                         command=self.abrir_excel,
                                         style='Success.TButton',
                                         cursor='hand2')
        self.btn_abrir_excel.pack(side=LEFT, padx=(0, 15), ipady=10, ipadx=20)

        self.btn_abrir_word = ttk.Button(results_buttons,
                                        text="📝  Abrir Documento Word",
                                        command=self.abrir_word,
                                        style='Success.TButton',
                                        cursor='hand2')
        self.btn_abrir_word.pack(side=LEFT, ipady=10, ipadx=20)

    def update_output_buttons(self):
        """Atualiza visual dos botões de seleção de saída"""
        # Cor para não selecionado (cinza claro)
        color_inactive = '#E3E5E8'
        # Cor para selecionado (verde)
        color_active = '#D4EDDA'
        border_active = '#28A745'
        border_inactive = '#E3E5E8'

        # Atualiza botão Excel
        if self.gerar_excel.get():
            for widget in self.excel_widgets:
                widget.config(bg=color_active)
            self.excel_button.config(relief=SOLID, bd=3, highlightbackground=border_active, highlightthickness=2)
        else:
            for widget in self.excel_widgets:
                widget.config(bg=color_inactive)
            self.excel_button.config(relief=SOLID, bd=2, highlightthickness=0)

        # Atualiza botão Word
        if self.gerar_word.get():
            for widget in self.word_widgets:
                widget.config(bg=color_active)
            self.word_button.config(relief=SOLID, bd=3, highlightbackground=border_active, highlightthickness=2)
        else:
            for widget in self.word_widgets:
                widget.config(bg=color_inactive)
            self.word_button.config(relief=SOLID, bd=2, highlightthickness=0)

    def update_submodo_buttons(self):
        """Atualiza visual dos botões de submodo INCRA e mostra/esconde campos"""
        # Cor para não selecionado (cinza claro)
        color_inactive = '#E3E5E8'
        # Cor para selecionado (azul claro)
        color_active = '#D4E8FF'
        border_active = '#5865F2'

        # Atualiza botão "Por Páginas"
        if self.submodo_incra.get() == "paginas":
            for widget in self.paginas_widgets:
                widget.config(bg=color_active)
            self.btn_paginas.config(relief=SOLID, bd=3, highlightbackground=border_active, highlightthickness=2)
            # Mostra campos de páginas
            self.paginas_fields_frame.pack(fill=X, pady=(15, 0))
        else:
            for widget in self.paginas_widgets:
                widget.config(bg=color_inactive)
            self.btn_paginas.config(relief=SOLID, bd=2, highlightthickness=0)

        # Atualiza botão "Por IA"
        if self.submodo_incra.get() == "ia":
            for widget in self.ia_widgets:
                widget.config(bg=color_active)
            self.btn_ia.config(relief=SOLID, bd=3, highlightbackground=border_active, highlightthickness=2)
            # Esconde campos de páginas
            self.paginas_fields_frame.pack_forget()
        else:
            for widget in self.ia_widgets:
                widget.config(bg=color_inactive)
            self.btn_ia.config(relief=SOLID, bd=2, highlightthickness=0)

    def setup_drag_drop(self):
        """Configura funcionalidade de drag & drop"""
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.handle_drop)
    
    def atualizar_modo(self):
        """Atualiza interface baseado no modo selecionado"""
        # Esconde ambos os frames
        self.normal_frame.pack_forget()
        self.incra_frame.pack_forget()

        # Mostra o frame correto
        if self.modo_operacao.get() == "normal":
            self.normal_frame.pack(fill=BOTH, expand=True)
            self.log("📄 MODO: Processar arquivo PDF", 'info')
            self.log("👉 Selecione um arquivo PDF acima", 'success')
        else:
            self.incra_frame.pack(fill=BOTH, expand=True)
            self.log("🏛️ MODO: Busca por Prenotação INCRA", 'incra')
            self.log("👉 Digite o número da prenotação acima", 'success')
    
    def handle_drop(self, event):
        """Manipula evento de drop de arquivo"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0].strip('{}')
            if file_path.lower().endswith('.pdf'):
                self.pdf_path.set(file_path)
                self.log(f"Arquivo selecionado: {Path(file_path).name}", 'success')
                self.update_drop_frame(True)
            else:
                messagebox.showwarning("Formato Inválido", 
                                     "Por favor, selecione um arquivo PDF.")
    
    def select_pdf(self):
        """Abre diálogo para selecionar arquivo PDF"""
        file_path = filedialog.askopenfilename(
            title="Selecionar Memorial Descritivo (PDF)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            self.log(f"Arquivo selecionado: {Path(file_path).name}", 'success')
            self.update_drop_frame(True)
    
    def update_drop_frame(self, has_file):
        """Atualiza visual da área de drop"""
        if has_file:
            self.drop_frame.config(bg='#90EE90')  # Verde claro
        else:
            self.drop_frame.config(bg='#E8F4FF')  # Azul claro

    def get_output_directory(self, prefixo=None):
        """
        Retorna o diretório de saída em Documents/Tabelas_Incra

        Args:
            prefixo: Nome da subpasta (ex: 'Prenotacao_00229885')

        Returns:
            Path do diretório de saída
        """
        home = Path.home()
        tabelas_incra = home / 'Documents' / 'Tabelas_Incra'

        if prefixo:
            output_dir = tabelas_incra / prefixo
        else:
            # Se não tem prefixo, usa pasta raiz
            output_dir = tabelas_incra

        # Cria pasta se não existir
        output_dir.mkdir(parents=True, exist_ok=True)

        return output_dir

    def abrir_arquivo(self, caminho):
        """
        Abre um arquivo com o aplicativo padrão do sistema

        Args:
            caminho: Path ou string do arquivo a abrir
        """
        try:
            if not caminho or not os.path.exists(caminho):
                messagebox.showerror("Erro", "Arquivo não encontrado!")
                return

            sistema = platform.system()

            if sistema == 'Windows':
                os.startfile(caminho)
            elif sistema == 'Darwin':  # macOS
                subprocess.run(['open', caminho])
            else:  # Linux
                subprocess.run(['xdg-open', caminho])

            self.log(f"✅ Arquivo aberto: {Path(caminho).name}", 'success')

        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}")
            self.log(f"❌ Erro ao abrir arquivo: {e}", 'error')

    def abrir_excel(self):
        """Abre o arquivo Excel gerado"""
        if self.excel_gerado:
            self.abrir_arquivo(self.excel_gerado)
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo Excel foi gerado ainda!")

    def abrir_word(self):
        """Abre o arquivo Word gerado"""
        if self.word_gerado:
            self.abrir_arquivo(self.word_gerado)
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo Word foi gerado ainda!")

    def mostrar_botoes_resultados(self):
        """Mostra os botões para abrir arquivos após processamento"""
        self.results_frame.pack(fill=X, pady=(25, 0))
        # Scroll para mostrar os botões
        self.canvas.yview_moveto(1.0)

    def limpar_arquivos_temporarios(self):
        """
        Remove arquivos temporários (.tif e .pdf convertido)
        Mantém apenas os arquivos finais (.xlsx e .docx)
        """
        if not self.arquivos_temporarios:
            return

        self.log("🗑️ Limpando arquivos temporários...", 'info')

        for arquivo in self.arquivos_temporarios:
            try:
                if os.path.exists(arquivo):
                    os.remove(arquivo)
                    nome_arquivo = Path(arquivo).name
                    self.log(f"  ✓ Removido: {nome_arquivo}", 'info')
            except Exception as e:
                self.log(f"  ⚠️ Erro ao remover {Path(arquivo).name}: {e}", 'warning')

        # Limpa a lista
        self.arquivos_temporarios = []
        self.log("✅ Arquivos temporários removidos!", 'success')

    def log(self, message, tag='info'):
        """Adiciona mensagem ao log"""
        self.log_text.insert(END, f"{message}\n", tag)
        self.log_text.see(END)
        self.root.update_idletasks()
    
    def clear_all(self):
        """Limpa todos os campos"""
        self.pdf_path.set("")
        self.prenotacao.set("")
        self.progress_value.set(0)
        self.status_text.set("Aguardando...")
        self.update_drop_frame(False)
        self.table_data = None
        self.log("Interface limpa. Pronto para novo processamento.", 'info')
    
    def validate_inputs(self):
        """Valida entradas antes de processar"""
        if self.modo_operacao.get() == "normal":
            if not self.pdf_path.get():
                messagebox.showerror("⚠️ Atenção",
                                   "Por favor, selecione um arquivo PDF!",
                                   icon='warning')
                return False
            if not os.path.exists(self.pdf_path.get()):
                messagebox.showerror("⚠️ Atenção",
                                   "O arquivo selecionado não foi encontrado!",
                                   icon='warning')
                return False
        else:  # modo incra
            if not self.prenotacao.get():
                messagebox.showerror("⚠️ Atenção",
                                   "Por favor, digite o número da prenotação!",
                                   icon='warning')
                return False

            # Validação para modo "Por Páginas"
            if self.submodo_incra.get() == "paginas":
                if not self.paginas_incra.get():
                    messagebox.showerror("⚠️ Atenção",
                                       "Por favor, especifique as páginas a serem extraídas!\n\n"
                                       "Exemplo: 1,2,4,7",
                                       icon='warning')
                    return False

                # Valida formato das páginas
                try:
                    paginas_str = self.paginas_incra.get()
                    paginas_list = [int(p.strip()) for p in paginas_str.split(',') if p.strip()]
                    if not paginas_list:
                        raise ValueError("Nenhuma página válida")
                    if any(p < 1 for p in paginas_list):
                        raise ValueError("Número de página inválido")
                except ValueError as e:
                    messagebox.showerror("⚠️ Atenção",
                                       "Formato de páginas inválido!\n\n"
                                       "Use números separados por vírgula.\n"
                                       "Exemplo: 1,2,4,7",
                                       icon='warning')
                    return False

        if not self.gerar_excel.get() and not self.gerar_word.get():
            messagebox.showwarning("⚠️ Atenção",
                                 "Selecione pelo menos um tipo de arquivo!\n\n"
                                 "Marque Excel ou Word (ou ambos).")
            return False

        return True
    
    def process_memorial(self):
        """Inicia processamento em thread separada"""
        if self.processing:
            messagebox.showwarning("⚠️ Atenção",
                                 "Aguarde! Já existe um processamento em andamento.")
            return

        if not self.validate_inputs():
            return

        self.process_btn.config(state='disabled', text='⏳ PROCESSANDO... AGUARDE')
        self.processing = True

        thread = threading.Thread(target=self.process_thread, daemon=True)
        thread.start()
    
    def process_thread(self):
        """Thread de processamento"""
        try:
            api_key = self.api_key.get()
            
            if self.modo_operacao.get() == "normal":
                # Modo Normal
                pdf_path = self.pdf_path.get()
                pdf_nome = Path(pdf_path).stem
                output_dir = self.get_output_directory(f"Processamento_{pdf_nome}")
                prefixo = pdf_nome

                self.update_progress(10, "Conectando com API...")
                self.log("📡 Processando PDF...", 'info')

                self.table_data = extract_table_from_pdf(pdf_path, api_key)
                
            else:
                # Modo INCRA
                prenotacao = self.prenotacao.get()
                
                # Testa acesso à rede primeiro
                self.update_progress(3, "Testando acesso à rede...")
                self.log("🔌 Testando acesso à rede INCRA...", 'incra')
                
                if not testar_acesso_rede():
                    raise Exception(
                        "Não foi possível acessar a rede do INCRA!\n\n"
                        "Verifique:\n"
                        "1. Conexão com a rede\n"
                        "2. Permissões de acesso\n"
                        f"3. Caminho: {INCRA_CONFIG['base_path']}"
                    )
                
                self.log("✅ Rede acessível!", 'success')
                
                self.update_progress(5, "Formatando prenotação...")
                prenotacao_formatada = formatar_prenotacao(prenotacao)
                self.log(f"✅ Prenotação: {prenotacao_formatada}", 'incra')
                
                self.update_progress(10, "Buscando arquivo na rede...")
                self.log("🔍 Buscando na rede INCRA...", 'incra')
                
                arquivo_tiff = buscar_arquivo_incra(prenotacao_formatada)
                if not arquivo_tiff:
                    raise Exception("Arquivo não encontrado na rede do INCRA!")
                
                self.log(f"✅ Arquivo encontrado!", 'success')
                
                self.update_progress(20, "Copiando para Documentos...")
                arquivo_local = copiar_para_downloads(arquivo_tiff, prenotacao_formatada)
                self.log(f"📁 Copiado para: {arquivo_local.parent.name}", 'success')

                # Registra arquivo .tif para deletar depois
                self.arquivos_temporarios.append(str(arquivo_local))

                self.update_progress(30, "Convertendo TIFF → PDF...")
                self.log("🔄 Convertendo TIFF para PDF...", 'info')
                pdf_path = converter_tiff_para_pdf(arquivo_local)
                self.log(f"✅ PDF criado", 'success')

                # Registra arquivo .pdf para deletar depois
                self.arquivos_temporarios.append(str(pdf_path))

                # Verifica submodo: IA ou Páginas
                if self.submodo_incra.get() == "paginas":
                    # Modo Manual - Extração por Páginas
                    paginas_str = self.paginas_incra.get()

                    if not paginas_str:
                        raise Exception("Por favor, especifique as páginas a serem extraídas!")

                    self.update_progress(35, "Extraindo páginas selecionadas...")
                    self.log(f"📄 Extraindo páginas: {paginas_str}", 'info')

                    # Importa a função de extração de páginas
                    from process_memorial_descritivo_v2 import extrair_paginas_pdf

                    # Extrai as páginas específicas
                    pdf_paginas = extrair_paginas_pdf(pdf_path, paginas_str)
                    self.log(f"✅ Páginas extraídas com sucesso!", 'success')

                    # Registra o PDF de páginas para deletar depois
                    self.arquivos_temporarios.append(str(pdf_paginas))

                    # Usa o PDF com páginas selecionadas para extração
                    self.update_progress(40, "Extraindo dados das páginas...")
                    self.log("📊 Extraindo Memorial das páginas selecionadas...", 'incra')
                    self.table_data = extract_table_from_pdf(str(pdf_paginas), api_key)

                else:
                    # Modo IA - Detecção Automática (fluxo original)
                    self.update_progress(40, "Extraindo dados (IA)...")
                    self.log("🤖 Extraindo Memorial do INCRA com IA...", 'incra')
                    self.table_data = extrair_memorial_incra(pdf_path, api_key)

                # Usa o diretório que já foi criado pela função copiar_para_downloads
                output_dir = pdf_path.parent
                prefixo = f"Prenotacao_{prenotacao_formatada}"
            
            # Dados extraídos com sucesso
            num_linhas = len(self.table_data.get('data', []))
            self.update_progress(60, f"Dados extraídos: {num_linhas} linhas")
            self.log(f"✅ Tabela extraída: {num_linhas} linhas", 'success')
            
            # Gera arquivos conforme escolha do usuário
            arquivos_gerados = []

            # Reset dos caminhos de arquivos gerados
            self.excel_gerado = None
            self.word_gerado = None

            if self.gerar_excel.get():
                self.update_progress(70, "Gerando Excel...")
                excel_path = output_dir / f"{prefixo}.xlsx"
                create_excel_file(self.table_data, str(excel_path))
                self.excel_gerado = str(excel_path)  # Salva caminho
                arquivos_gerados.append(f"📊 {excel_path.name}")
                self.log(f"✅ Excel: {excel_path.name}", 'success')

            if self.gerar_word.get():
                self.update_progress(85, "Gerando Word...")
                word_path = output_dir / f"{prefixo}.docx"
                create_word_file(self.table_data, str(word_path))
                self.word_gerado = str(word_path)  # Salva caminho
                arquivos_gerados.append(f"📝 {word_path.name}")
                self.log(f"✅ Word: {word_path.name}", 'success')

            # Limpa arquivos temporários (.tif e .pdf)
            if self.arquivos_temporarios:
                self.limpar_arquivos_temporarios()

            self.update_progress(100, "Concluído!")
            self.log("="*50, 'success')
            self.log("✨ PROCESSAMENTO CONCLUÍDO!", 'success')
            self.log(f"📂 Pasta: {output_dir.name}", 'info')
            self.log(f"📍 Local completo:", 'info')
            self.log(f"   {output_dir}", 'info')
            self.log("="*50, 'success')

            # Mostra botões para abrir arquivos
            if self.excel_gerado or self.word_gerado:
                self.root.after(100, self.mostrar_botoes_resultados)

            # Mensagem de sucesso
            msg = "✅ Processamento concluído com sucesso!\n\n"
            msg += "Arquivos gerados:\n"
            msg += "\n".join(arquivos_gerados)
            msg += f"\n\n📂 Salvos em:\n{output_dir}"
            msg += "\n\nUse os botões abaixo para abrir os arquivos!"

            self.root.after(100, lambda: messagebox.showinfo("✨ Sucesso!", msg))
            
        except Exception as ex:
            self.log(f"❌ ERRO: {str(ex)}", 'error')
            self.update_progress(0, "Erro no processamento!")
            self.root.after(100, lambda: messagebox.showerror(
                "Erro",
                f"Erro durante o processamento:\n\n{str(ex)}\n\n"
                f"Verifique o log para mais detalhes."
            ))
        
        finally:
            self.root.after(100, lambda: self.process_btn.config(
                state='normal',
                text='✅ PROCESSAR AGORA'
            ))
            self.processing = False
    
    def update_progress(self, value, status):
        """Atualiza barra de progresso e status"""
        self.progress_value.set(value)
        self.status_text.set(status)
        self.root.update_idletasks()


def main():
    """Função principal"""
    try:
        root = TkinterDnD.Tk()
    except:
        print("❌ Erro: tkinterdnd2 não está instalado!")
        print("Instale com: pip install tkinterdnd2 --break-system-packages")
        sys.exit(1)
    
    app = MemorialGUI_V2(root)
    
    # Centraliza janela
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()