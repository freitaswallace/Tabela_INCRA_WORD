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
        self.root.title("Processador de Memorial Descritivo v2.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Variáveis
        self.pdf_path = StringVar()
        self.prenotacao = StringVar()
        self.api_key = StringVar()
        self.modo_operacao = StringVar(value="normal")  # "normal" ou "incra"
        self.status_text = StringVar(value="Aguardando...")
        self.progress_value = IntVar(value=0)
        self.processing = False
        self.table_data = None
        
        # Checkboxes para saída
        self.gerar_excel = BooleanVar(value=True)
        self.gerar_word = BooleanVar(value=True)
        
        # Tenta carregar API Key do ambiente
        env_key = os.environ.get('GEMINI_API_KEY', '')
        if env_key:
            self.api_key.set(env_key)
        
        # Configurar estilo
        self.setup_style()
        
        # Criar interface
        self.create_widgets()
        
        # Configurar drag & drop (apenas para modo normal)
        self.setup_drag_drop()
    
    def setup_style(self):
        """Configura o estilo visual da interface"""
        style = ttk.Style()
        style.theme_use('clam')
        
        self.colors = {
            'primary': '#2196F3',
            'success': '#4CAF50',
            'danger': '#F44336',
            'warning': '#FF9800',
            'incra': '#8E24AA',
            'bg': '#F5F5F5',
            'text': '#212121',
            'border': '#E0E0E0'
        }
        
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground=self.colors['primary'])
        style.configure('Subtitle.TLabel', font=('Segoe UI', 10), foreground='gray')
        style.configure('Status.TLabel', font=('Segoe UI', 10), foreground=self.colors['text'])
        style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'))
        
    def create_widgets(self):
        """Cria todos os widgets da interface"""
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=BOTH, expand=True)
        
        # ===== CABEÇALHO =====
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=X, pady=(0, 20))
        
        title = ttk.Label(header_frame, text="🚀 Processador de Memorial Descritivo v2.0", 
                         style='Title.TLabel')
        title.pack(anchor=W)
        
        subtitle = ttk.Label(header_frame, 
                            text="Extração automatizada com Modo INCRA • Gemini 2.5 Flash Lite",
                            style='Subtitle.TLabel')
        subtitle.pack(anchor=W)
        
        ttk.Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=10)
        
        # ===== SEÇÃO API KEY =====
        api_frame = ttk.LabelFrame(main_frame, text="🔑 API Key do Google Gemini", padding="15")
        api_frame.pack(fill=X, pady=(0, 15))
        
        api_input_frame = ttk.Frame(api_frame)
        api_input_frame.pack(fill=X)
        
        api_entry = ttk.Entry(api_input_frame, textvariable=self.api_key, 
                             font=('Consolas', 10), show='•')
        api_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        
        def toggle_api_visibility():
            if api_entry['show'] == '•':
                api_entry['show'] = ''
                show_btn.config(text='👁️')
            else:
                api_entry['show'] = '•'
                show_btn.config(text='🔒')
        
        show_btn = ttk.Button(api_input_frame, text='🔒', width=3, 
                             command=toggle_api_visibility)
        show_btn.pack(side=LEFT, padx=(0, 5))
        
        help_btn = ttk.Button(api_input_frame, text='❓', width=3, 
                             command=self.show_api_help)
        help_btn.pack(side=LEFT)
        
        # ===== MODO DE OPERAÇÃO =====
        modo_frame = ttk.LabelFrame(main_frame, text="🎯 Modo de Operação", padding="15")
        modo_frame.pack(fill=X, pady=(0, 15))
        
        radio_frame = ttk.Frame(modo_frame)
        radio_frame.pack(fill=X)
        
        modo_normal_radio = ttk.Radiobutton(
            radio_frame, 
            text="📄 Modo Normal (Fornecer PDF)",
            variable=self.modo_operacao,
            value="normal",
            command=self.atualizar_modo
        )
        modo_normal_radio.pack(side=LEFT, padx=(0, 20))
        
        modo_incra_radio = ttk.Radiobutton(
            radio_frame,
            text="🏛️ Modo Prenotação INCRA (Busca Automática)",
            variable=self.modo_operacao,
            value="incra",
            command=self.atualizar_modo
        )
        modo_incra_radio.pack(side=LEFT)
        
        # ===== FRAMES DE ENTRADA (Normal e INCRA) =====
        
        # Container para trocar entre modos
        self.input_container = ttk.Frame(main_frame)
        self.input_container.pack(fill=BOTH, expand=True, pady=(0, 15))
        
        # Frame Modo Normal (PDF)
        self.normal_frame = ttk.LabelFrame(self.input_container, text="📄 Arquivo PDF", padding="15")
        
        self.drop_frame = Frame(self.normal_frame, bg='white', relief=GROOVE, bd=2)
        self.drop_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        drop_label = Label(self.drop_frame, 
                          text="📂\n\nArraste o PDF aqui\nou\nClique para selecionar",
                          bg='white', fg='gray', font=('Segoe UI', 12),
                          cursor='hand2')
        drop_label.pack(expand=True)
        drop_label.bind('<Button-1>', lambda e: self.select_pdf())
        
        path_frame = ttk.Frame(self.normal_frame)
        path_frame.pack(fill=X)
        
        path_label = ttk.Label(path_frame, text="Arquivo selecionado:")
        path_label.pack(anchor=W)
        
        path_entry = ttk.Entry(path_frame, textvariable=self.pdf_path, 
                              state='readonly', font=('Segoe UI', 9))
        path_entry.pack(fill=X, pady=(2, 0))
        
        # Frame Modo INCRA (Prenotação)
        self.incra_frame = ttk.LabelFrame(self.input_container, text="🏛️ Prenotação INCRA", padding="15")
        
        incra_info = ttk.Label(self.incra_frame, 
                              text="Digite o número da prenotação (ex: 229885 ou 00229885)\n"
                                   "O sistema buscará automaticamente na rede do INCRA",
                              foreground='gray', font=('Segoe UI', 9))
        incra_info.pack(anchor=W, pady=(0, 10))
        
        prenotacao_frame = ttk.Frame(self.incra_frame)
        prenotacao_frame.pack(fill=X)
        
        prenotacao_label = ttk.Label(prenotacao_frame, text="Número da Prenotação:")
        prenotacao_label.pack(anchor=W)
        
        prenotacao_entry = ttk.Entry(prenotacao_frame, textvariable=self.prenotacao,
                                     font=('Consolas', 12))
        prenotacao_entry.pack(fill=X, pady=(2, 0))
        
        # Mostra frame inicial (Normal)
        self.normal_frame.pack(fill=BOTH, expand=True)
        
        # ===== ESCOLHA DE SAÍDA =====
        output_frame = ttk.LabelFrame(main_frame, text="💾 Arquivos de Saída", padding="15")
        output_frame.pack(fill=X, pady=(0, 15))
        
        output_info = ttk.Label(output_frame, 
                               text="Escolha quais arquivos deseja gerar após o processamento:",
                               foreground='gray', font=('Segoe UI', 9))
        output_info.pack(anchor=W, pady=(0, 10))
        
        check_frame = ttk.Frame(output_frame)
        check_frame.pack(fill=X)
        
        excel_check = ttk.Checkbutton(check_frame, text="📊 Excel (.xlsx)",
                                     variable=self.gerar_excel)
        excel_check.pack(side=LEFT, padx=(0, 20))
        
        word_check = ttk.Checkbutton(check_frame, text="📝 Word (.docx)",
                                    variable=self.gerar_word)
        word_check.pack(side=LEFT)
        
        # ===== BOTÕES DE AÇÃO =====
        process_frame = ttk.Frame(main_frame)
        process_frame.pack(fill=X, pady=(0, 15))
        
        self.process_btn = ttk.Button(process_frame, text="🚀 Processar", 
                                      command=self.process_memorial,
                                      style='Primary.TButton')
        self.process_btn.pack(side=LEFT, padx=(0, 10))
        
        clear_btn = ttk.Button(process_frame, text="🗑️ Limpar", 
                              command=self.clear_all)
        clear_btn.pack(side=LEFT)
        
        # ===== BARRA DE PROGRESSO =====
        progress_frame = ttk.LabelFrame(main_frame, text="📊 Progresso", padding="15")
        progress_frame.pack(fill=X, pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate',
                                           variable=self.progress_value)
        self.progress_bar.pack(fill=X, pady=(0, 5))
        
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_text,
                                      style='Status.TLabel')
        self.status_label.pack(anchor=W)
        
        # ===== LOG =====
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log", padding="10")
        log_frame.pack(fill=BOTH, expand=True)
        
        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=RIGHT, fill=Y)
        
        self.log_text = Text(log_frame, height=8, font=('Consolas', 9),
                            yscrollcommand=log_scroll.set, wrap=WORD,
                            bg='#1E1E1E', fg='#D4D4D4', insertbackground='white')
        self.log_text.pack(fill=BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)
        
        self.log_text.tag_config('info', foreground='#4FC3F7')
        self.log_text.tag_config('success', foreground='#81C784')
        self.log_text.tag_config('error', foreground='#E57373')
        self.log_text.tag_config('warning', foreground='#FFB74D')
        self.log_text.tag_config('incra', foreground='#CE93D8')
        
        self.log("Bem-vindo ao Processador v2.0!", 'info')
        self.log("🎯 Escolha o modo de operação", 'info')
        self.log("💡 Modo INCRA: Busca automática em rede", 'incra')
        
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
            self.log("📄 Modo Normal ativado", 'info')
        else:
            self.incra_frame.pack(fill=BOTH, expand=True)
            self.log("🏛️ Modo INCRA ativado", 'incra')
            self.log(f"📂 Rede: {INCRA_CONFIG['base_path']}", 'info')
    
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
            self.drop_frame.config(bg='#E8F5E9')
        else:
            self.drop_frame.config(bg='white')
    
    def show_api_help(self):
        """Mostra ajuda sobre API Key"""
        help_text = """🔑 Como obter a API Key do Google Gemini:

1. Acesse: https://aistudio.google.com/app/apikey
2. Faça login com sua conta Google
3. Clique em "Create API Key"
4. Copie a chave gerada (formato: AIza...)
5. Cole no campo acima ou configure como variável de ambiente

📊 Limites do plano gratuito:
• 15 requisições/minuto
• 1.500 requisições/dia"""
        
        messagebox.showinfo("Ajuda - API Key", help_text)
    
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
        if not self.api_key.get():
            messagebox.showerror("Erro", "API Key não configurada!")
            return False
        
        if self.modo_operacao.get() == "normal":
            if not self.pdf_path.get():
                messagebox.showerror("Erro", "Nenhum arquivo PDF selecionado!")
                return False
            if not os.path.exists(self.pdf_path.get()):
                messagebox.showerror("Erro", "Arquivo não encontrado!")
                return False
        else:  # modo incra
            if not self.prenotacao.get():
                messagebox.showerror("Erro", "Número da prenotação não informado!")
                return False
        
        if not self.gerar_excel.get() and not self.gerar_word.get():
            messagebox.showwarning("Aviso", "Selecione pelo menos um tipo de arquivo para gerar!")
            return False
        
        return True
    
    def process_memorial(self):
        """Inicia processamento em thread separada"""
        if self.processing:
            messagebox.showwarning("Aviso", "Já existe um processamento em andamento.")
            return
        
        if not self.validate_inputs():
            return
        
        self.process_btn.config(state='disabled', text='⏳ Processando...')
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
                output_dir = Path(pdf_path).parent
                prefixo = "output"
                
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
                
                self.update_progress(20, "Copiando para Downloads...")
                arquivo_local = copiar_para_downloads(arquivo_tiff, prenotacao_formatada)
                self.log(f"📁 Copiado para: {arquivo_local.parent.name}", 'success')
                
                self.update_progress(30, "Convertendo TIFF → PDF...")
                self.log("🔄 Convertendo TIFF para PDF...", 'info')
                pdf_path = converter_tiff_para_pdf(arquivo_local)
                self.log(f"✅ PDF criado", 'success')
                
                self.update_progress(40, "Extraindo dados...")
                self.log("📊 Extraindo Memorial do INCRA...", 'incra')
                self.table_data = extrair_memorial_incra(pdf_path, api_key)
                
                output_dir = pdf_path.parent
                prefixo = f"prenotacao_{prenotacao_formatada}"
            
            # Dados extraídos com sucesso
            num_linhas = len(self.table_data.get('data', []))
            self.update_progress(60, f"Dados extraídos: {num_linhas} linhas")
            self.log(f"✅ Tabela extraída: {num_linhas} linhas", 'success')
            
            # Gera arquivos conforme escolha do usuário
            arquivos_gerados = []
            
            if self.gerar_excel.get():
                self.update_progress(70, "Gerando Excel...")
                excel_path = output_dir / f"{prefixo}.xlsx"
                create_excel_file(self.table_data, str(excel_path))
                arquivos_gerados.append(f"📊 {excel_path.name}")
                self.log(f"✅ Excel: {excel_path.name}", 'success')
            
            if self.gerar_word.get():
                self.update_progress(85, "Gerando Word...")
                word_path = output_dir / f"{prefixo}.docx"
                create_word_file(self.table_data, str(word_path))
                arquivos_gerados.append(f"📝 {word_path.name}")
                self.log(f"✅ Word: {word_path.name}", 'success')
            
            self.update_progress(100, "Concluído!")
            self.log("="*50, 'success')
            self.log("✨ PROCESSAMENTO CONCLUÍDO!", 'success')
            self.log(f"📂 Local: {output_dir}", 'info')
            self.log("="*50, 'success')
            
            # Mensagem de sucesso
            msg = "Processamento concluído!\n\n"
            msg += "\n".join(arquivos_gerados)
            msg += f"\n\n📂 Local: {output_dir}"
            
            self.root.after(100, lambda: messagebox.showinfo("Sucesso!", msg))
            
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
                text='🚀 Processar'
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