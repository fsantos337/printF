import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import threading
import json
import csv
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple


class LiquidGlassStyle:
    """Estilo Liquid Glass para a aplicação"""
    
    # Cores do tema Liquid Glass
    BG_PRIMARY = "#0a0e27"
    BG_SECONDARY = "#1a1f3a"
    BG_CARD = "#252b48"
    BG_HOVER = "#2d3454"
    
    ACCENT_PRIMARY = "#00d4ff"
    ACCENT_SECONDARY = "#7b2ff7"
    ACCENT_SUCCESS = "#00ff88"
    ACCENT_WARNING = "#ffd93d"
    ACCENT_ERROR = "#ff6b6b"
    
    TEXT_PRIMARY = "#ffffff"
    TEXT_SECONDARY = "#a0aec0"
    TEXT_MUTED = "#718096"
    
    GLASS_ALPHA = 0.15
    BORDER_RADIUS = 16
    
    @staticmethod
    def configure_style():
        """Configura o estilo ttk com tema Liquid Glass"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Frame com efeito glass
        style.configure("Glass.TFrame",
                       background=LiquidGlassStyle.BG_CARD,
                       relief="flat",
                       borderwidth=0)
        
        # Labels
        style.configure("Glass.TLabel",
                       background=LiquidGlassStyle.BG_CARD,
                       foreground=LiquidGlassStyle.TEXT_PRIMARY,
                       font=("Segoe UI", 10))
        
        style.configure("Title.TLabel",
                       background=LiquidGlassStyle.BG_PRIMARY,
                       foreground=LiquidGlassStyle.ACCENT_PRIMARY,
                       font=("Segoe UI", 24, "bold"))
        
        style.configure("Subtitle.TLabel",
                       background=LiquidGlassStyle.BG_PRIMARY,
                       foreground=LiquidGlassStyle.TEXT_SECONDARY,
                       font=("Segoe UI", 10))
        
        style.configure("Header.TLabel",
                       background=LiquidGlassStyle.BG_CARD,
                       foreground=LiquidGlassStyle.TEXT_PRIMARY,
                       font=("Segoe UI", 12, "bold"))
        
        # Entry
        style.configure("Glass.TEntry",
                       fieldbackground=LiquidGlassStyle.BG_SECONDARY,
                       foreground=LiquidGlassStyle.TEXT_PRIMARY,
                       bordercolor=LiquidGlassStyle.ACCENT_PRIMARY,
                       lightcolor=LiquidGlassStyle.ACCENT_PRIMARY,
                       darkcolor=LiquidGlassStyle.BG_SECONDARY,
                       borderwidth=1,
                       relief="flat")
        
        # Buttons
        style.configure("Accent.TButton",
                       background=LiquidGlassStyle.ACCENT_PRIMARY,
                       foreground=LiquidGlassStyle.BG_PRIMARY,
                       borderwidth=0,
                       focuscolor=LiquidGlassStyle.ACCENT_PRIMARY,
                       font=("Segoe UI", 10, "bold"),
                       relief="flat")
        
        style.map("Accent.TButton",
                 background=[("active", LiquidGlassStyle.ACCENT_SECONDARY)],
                 relief=[("pressed", "flat")])
        
        style.configure("Glass.TButton",
                       background=LiquidGlassStyle.BG_HOVER,
                       foreground=LiquidGlassStyle.TEXT_PRIMARY,
                       borderwidth=0,
                       font=("Segoe UI", 9),
                       relief="flat")
        
        style.map("Glass.TButton",
                 background=[("active", LiquidGlassStyle.BG_SECONDARY)])
        
        # Progressbar
        style.configure("Glass.Horizontal.TProgressbar",
                       background=LiquidGlassStyle.ACCENT_PRIMARY,
                       troughcolor=LiquidGlassStyle.BG_SECONDARY,
                       borderwidth=0,
                       thickness=8)
        
        # Notebook
        style.configure("Glass.TNotebook",
                       background=LiquidGlassStyle.BG_PRIMARY,
                       borderwidth=0)
        
        style.configure("Glass.TNotebook.Tab",
                       background=LiquidGlassStyle.BG_SECONDARY,
                       foreground=LiquidGlassStyle.TEXT_SECONDARY,
                       padding=[20, 10],
                       borderwidth=0,
                       font=("Segoe UI", 10))
        
        style.map("Glass.TNotebook.Tab",
                 background=[("selected", LiquidGlassStyle.BG_CARD)],
                 foreground=[("selected", LiquidGlassStyle.ACCENT_PRIMARY)])
        
        # Separator
        style.configure("Glass.TSeparator",
                       background=LiquidGlassStyle.BG_HOVER)


class ConfigManager:
    """Gerenciador de configuração de campos"""
    
    DEFAULT_CONFIG = [
        {"label": "Projeto:", "key": "projeto"},
        {"label": "Módulo:", "key": "modulo"},
        {"label": "Versão:", "key": "versao"},
        {"label": "Responsável:", "key": "responsavel"},
        {"label": "Data:", "key": "data"},
        {"label": "Ambiente:", "key": "ambiente"}
    ]

    def __init__(self, config_file='config_campos.json'):
        self.config_file = Path(config_file)

    def load_config(self):
        """Carrega configuração do arquivo JSON"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return self._create_default_config()
        except:
            return self.DEFAULT_CONFIG

    def _create_default_config(self):
        """Cria arquivo de configuração padrão"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)
            return self.DEFAULT_CONFIG
        except:
            return self.DEFAULT_CONFIG


class CSVReader:
    """Responsável pela leitura de arquivos CSV"""
    
    ENCODINGS = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'windows-1252']

    @staticmethod
    def read_csv(file_path: str) -> Optional[List[str]]:
        """Lê um arquivo CSV e retorna a lista de nomes"""
        try:
            return CSVReader._read_with_pandas(file_path) or CSVReader._read_manual(file_path)
        except Exception as e:
            print(f"Erro ao ler CSV: {e}")
            return None

    @staticmethod
    def _read_with_pandas(file_path: str) -> Optional[List[str]]:
        """Tenta ler o CSV usando pandas"""
        for encoding in CSVReader.ENCODINGS:
            try:
                df = pd.read_csv(file_path, encoding=encoding, engine='python', 
                               on_bad_lines='skip')
                if 'Nome' in df.columns:
                    nomes = df['Nome'].dropna().str.strip()
                    return nomes[nomes != ''].tolist()
            except Exception:
                continue
        return None

    @staticmethod
    def _read_manual(file_path: str) -> Optional[List[str]]:
        """Leitura manual do CSV como fallback"""
        for encoding in CSVReader.ENCODINGS:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    return CSVReader._parse_csv_lines(file.readlines())
            except Exception:
                continue
        return None

    @staticmethod
    def _parse_csv_lines(lines: List[str]) -> Optional[List[str]]:
        """Parseia as linhas do CSV manualmente"""
        if not lines:
            return None

        header_line = next((line for line in lines if 'Nome' in line), None)
        if not header_line:
            return None

        headers = header_line.strip().split(',')
        if 'Nome' not in headers:
            return None

        nome_index = headers.index('Nome')
        nomes = []

        for line in lines[1:]:
            try:
                if '"' in line:
                    reader = csv.reader([line])
                    parts = next(reader)
                else:
                    parts = line.strip().split(',')

                if len(parts) > nome_index:
                    nome = parts[nome_index].strip().strip('"')
                    if nome:
                        nomes.append(nome)
            except Exception:
                continue

        return nomes if nomes else None


class DocumentProcessor:
    """Processa e gera documentos Word baseados em templates"""
    
    @staticmethod
    def clean_filename(filename: str, max_length: int = 100) -> str:
        """Limpa o nome do arquivo removendo caracteres inválidos"""
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', filename)
        cleaned = cleaned.strip()[:max_length]
        return cleaned or "caso_teste"

    @staticmethod
    def adjust_template_fields(doc: Document, field_mapping: Dict[str, str]) -> None:
        """Ajusta os campos do template conforme o mapeamento fornecido"""
        for paragraph in doc.paragraphs:
            DocumentProcessor._adjust_paragraph_fields(paragraph, field_mapping)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        DocumentProcessor._adjust_paragraph_fields(paragraph, field_mapping)

    @staticmethod
    def _adjust_paragraph_fields(paragraph, field_mapping: Dict[str, str]) -> None:
        """Ajusta os campos em um parágrafo específico"""
        texto_original = paragraph.text.strip()
        if ':' in texto_original:
            field_key = texto_original.split(':', 1)[0].strip()
            if field_key in field_mapping:
                paragraph.text = f"{field_mapping[field_key]}: "

    @staticmethod
    def fill_template(doc: Document, data: Dict[str, str], field_mapping: Dict[str, str]) -> None:
        """Preenche o template com os dados fornecidos"""
        DocumentProcessor.adjust_template_fields(doc, field_mapping)
        
        label_to_value = {}
        for original_key, label in field_mapping.items():
            label_to_value[label] = data.get(original_key, '')
        
        label_to_value['Caso de Teste'] = data.get('Caso de Teste', '')
        
        for paragraph in doc.paragraphs:
            DocumentProcessor._fill_paragraph(paragraph, label_to_value)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        DocumentProcessor._fill_paragraph(paragraph, label_to_value)

    @staticmethod
    def _fill_paragraph(paragraph, label_to_value: Dict[str, str]) -> None:
        """Preenche um parágrafo específico com os dados"""
        texto = paragraph.text.strip()
        if ':' in texto:
            field_name = texto.split(':', 1)[0].strip()
            if field_name in label_to_value:
                paragraph.text = f"{field_name}: {label_to_value[field_name]}"


class DefaultDocumentGenerator:
    """Gera documentos padrão quando nenhum template é fornecido"""
    
    @staticmethod
    def create_default_document(data: Dict[str, str], field_config: List[Dict]) -> Document:
        """Cria um documento padrão com estrutura organizada"""
        doc = Document()
        
        title = doc.add_heading('Evidências de Teste - Documentação', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        date_para = doc.add_paragraph(f"Gerado em: {current_time}")
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        
        doc.add_heading('Informações do Teste', level=1)
        
        table = doc.add_table(rows=len(field_config) + 1, cols=2)
        table.style = 'Light Grid Accent 1'
        
        header_cells = table.rows[0].cells
        header_cells[0].text = "Campo"
        header_cells[1].text = "Valor"
        
        for cell in header_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
        
        for i, campo_info in enumerate(field_config, 1):
            key = campo_info['key']
            label = campo_info['label'].rstrip(':')
            
            row_cells = table.rows[i].cells
            row_cells[0].text = label
            row_cells[1].text = data.get(key, 'Não informado')
        
        doc.add_paragraph()
        
        doc.add_heading('Caso de Teste', level=1)
        caso_teste_para = doc.add_paragraph()
        caso_teste_para.add_run('Nome do Caso de Teste: ').bold = True
        caso_teste_para.add_run(data.get('Caso de Teste', 'Não informado'))
        
        doc.add_heading('Descrição do Teste', level=2)
        doc.add_paragraph(
            "Esta seção deve conter a descrição detalhada do caso de teste executado, "
            "incluindo pré-condições, passos de execução e resultados esperados."
        )
        
        doc.add_heading('Evidências Coletadas', level=2)
        doc.add_paragraph("Registro das evidências coletadas durante a execução do teste:")
        
        evidencias_table = doc.add_table(rows=5, cols=3)
        evidencias_table.style = 'Light Grid Accent 1'
        
        evidencias_header = evidencias_table.rows[0].cells
        headers = ['Etapa', 'Evidência', 'Resultado']
        for col, header in enumerate(headers):
            evidencias_header[col].text = header
            for paragraph in evidencias_header[col].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
        
        etapas = [
            'Pré-condições',
            'Configuração Inicial', 
            'Execução do Teste',
            'Pós-condições',
            'Resultado Final'
        ]
        
        for row, etapa in enumerate(etapas, 1):
            row_cells = evidencias_table.rows[row].cells
            row_cells[0].text = etapa
            row_cells[1].text = "[Descreva a evidência coletada]"
            row_cells[2].text = "[Resultado obtido - OK/Erro]"
        
        doc.add_paragraph()
        
        doc.add_heading('Observações e Comentários', level=2)
        doc.add_paragraph("Adicione observações relevantes sobre a execução do teste:")
        
        obs_para = doc.add_paragraph()
        obs_para.add_run("Observações Gerais:\n").bold = True
        obs_para.add_run("• [Insira observações sobre problemas encontrados]\n")
        obs_para.add_run("• [Comentários sobre o comportamento do sistema]\n")
        obs_para.add_run("• [Sugestões de melhorias]\n")
        obs_para.add_run("• [Outras informações relevantes]")
        
        doc.add_paragraph()
        footer = doc.add_paragraph()
        footer.add_run("Documento gerado automaticamente pelo PrintF - Gerador de Templates").italic = True
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return doc


class TemplateGenerator:
    """Gerador de templates de exemplo"""
    
    @staticmethod
    def create_example_template(field_config: List[Dict]) -> bool:
        """Cria um template de exemplo com base na configuração"""
        try:
            doc = Document()
            doc.add_heading('Template de Evidências de Teste', level=1)
            
            info_para = doc.add_paragraph()
            info_para.add_run("Instruções: ").bold = True
            info_para.add_run("Este é um template de exemplo. Os campos abaixo serão preenchidos automaticamente.")
            
            doc.add_paragraph()
            
            for campo_info in field_config:
                doc.add_paragraph(f"{campo_info['label']} [VALOR]")
            
            doc.add_paragraph()
            
            doc.add_heading('Detalhes do Caso de Teste', level=2)
            doc.add_paragraph("Caso de Teste: [NOME_DO_CASO]")
            
            table = doc.add_table(rows=4, cols=2)
            table.style = 'Table Grid'
            table.cell(0, 0).text = 'Caminho da Funcionalidade:'
            table.cell(0, 1).text = '[CAMINHO]'
            table.cell(1, 0).text = 'Resultado Esperado:'
            table.cell(1, 1).text = '[RESULTADO_ESPERADO]'
            table.cell(2, 0).text = 'Resultado Obtido:'
            table.cell(2, 1).text = '[RESULTADO_OBTIDO]'
            table.cell(3, 0).text = 'Observações:'
            table.cell(3, 1).text = '[OBSERVACOES]'
            
            doc.save('template_evidencias.docx')
            return True
        except Exception as e:
            print(f"Erro ao criar template: {e}")
            return False


class TemplateGeneratorModule:
    """Módulo completo de geração de templates com design Liquid Glass"""
    
    def __init__(self, parent, settings):
        self.parent = parent
        self.settings = settings
        self.window = None
        
        self.config_manager = ConfigManager()
        self.csv_reader = CSVReader()
        self.doc_processor = DocumentProcessor()
        self.default_doc_generator = DefaultDocumentGenerator()
        
        self.campos_config = self.config_manager.load_config()
        self.campos_entries = {}
        
        self.progress = None
        self.log_text = None
        self.gerar_btn = None
        
        LiquidGlassStyle.configure_style()

    def show(self):
        """Mostra interface completa"""
        if self.window and self.window.winfo_exists():
            self.window.lift()
            return
        
        self.window = tk.Toplevel(self.parent)
        self.window.title("PrintF - Gerador de Templates")
        self.window.geometry("1000x800")
        self.window.minsize(900, 700)
        
        # Configurar cor de fundo
        self.window.configure(bg=LiquidGlassStyle.BG_PRIMARY)
        
        self._create_complete_ui()

    def hide(self):
        """Esconde o módulo"""
        if self.window and self.window.winfo_exists():
            self.window.destroy()
        self.window = None

    def _create_complete_ui(self):
        """Cria interface completa com design Liquid Glass"""
        main_frame = tk.Frame(self.window, bg=LiquidGlassStyle.BG_PRIMARY)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header com gradiente visual
        self._create_header(main_frame)
        
        # Notebook com abas estilizadas
        notebook = ttk.Notebook(main_frame, style="Glass.TNotebook")
        notebook.pack(fill=tk.BOTH, expand=True, pady=20)
        
        # Abas
        config_frame = tk.Frame(notebook, bg=LiquidGlassStyle.BG_CARD)
        notebook.add(config_frame, text="⚙️  Configuração")
        
        fields_frame = tk.Frame(notebook, bg=LiquidGlassStyle.BG_CARD)
        notebook.add(fields_frame, text="📝  Campos")
        
        log_frame = tk.Frame(notebook, bg=LiquidGlassStyle.BG_CARD)
        notebook.add(log_frame, text="📋  Log")
        
        self._create_config_tab(config_frame)
        self._create_fields_tab(fields_frame)
        self._create_log_tab(log_frame)

    def _create_header(self, parent):
        """Cria header com design moderno"""
        header_frame = tk.Frame(parent, bg=LiquidGlassStyle.BG_PRIMARY)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Título com efeito de brilho
        title_label = ttk.Label(
            header_frame, 
            text="📄 GERADOR DE TEMPLATES", 
            style="Title.TLabel"
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            header_frame,
            text="Crie documentos em lote com design profissional e automação inteligente",
            style="Subtitle.TLabel"
        )
        subtitle_label.pack(pady=(5, 0))

    def _create_config_tab(self, parent):
        """Cria aba de configuração com cards"""
        container = tk.Frame(parent, bg=LiquidGlassStyle.BG_CARD)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Card para CSV
        self._create_file_card(
            container, 
            "📊 Arquivo CSV com Casos de Teste", 
            "csv_var",
            self._select_csv,
            0,
            required=True
        )
        
        # Card para Template
        self._create_file_card(
            container,
            "📄 Template DOCX (Opcional)",
            "template_var",
            self._select_template,
            1
        )
        
        # Card para Output
        self._create_file_card(
            container,
            "📁 Diretório de Saída",
            "output_var",
            self._select_output,
            2,
            default_value=self.settings.get('output_dir', 'evidencias_geradas')
        )
        
        # Botões de ação
        self._create_action_buttons(container)
        
        # Card de informações
        self._create_info_card(container)

    def _create_file_card(self, parent, title, var_name, command, row, required=False, default_value=""):
        """Cria um card para seleção de arquivo"""
        card = tk.Frame(parent, bg=LiquidGlassStyle.BG_SECONDARY, relief="flat")
        card.pack(fill=tk.X, pady=10)
        
        # Padding interno
        card_content = tk.Frame(card, bg=LiquidGlassStyle.BG_SECONDARY)
        card_content.pack(fill=tk.X, padx=15, pady=15)
        
        # Título
        title_text = f"{title} {'*' if required else ''}"
        title_label = tk.Label(
            card_content,
            text=title_text,
            bg=LiquidGlassStyle.BG_SECONDARY,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            font=("Segoe UI", 11, "bold")
        )
        title_label.pack(anchor="w", pady=(0, 10))
        
        # Frame para entry e botão
        input_frame = tk.Frame(card_content, bg=LiquidGlassStyle.BG_SECONDARY)
        input_frame.pack(fill=tk.X)
        
        # Entry
        var = tk.StringVar(value=default_value)
        setattr(self, var_name, var)
        
        entry = tk.Entry(
            input_frame,
            textvariable=var,
            bg=LiquidGlassStyle.BG_PRIMARY,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            insertbackground=LiquidGlassStyle.ACCENT_PRIMARY,
            relief="flat",
            font=("Segoe UI", 10),
            bd=0
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, ipadx=10)
        
        # Botão
        btn = tk.Button(
            input_frame,
            text="📂 Procurar",
            command=command,
            bg=LiquidGlassStyle.BG_HOVER,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            activebackground=LiquidGlassStyle.ACCENT_PRIMARY,
            activeforeground=LiquidGlassStyle.BG_PRIMARY,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            bd=0,
            padx=20,
            pady=8
        )
        btn.pack(side=tk.RIGHT, padx=(10, 0))

    def _create_action_buttons(self, parent):
        """Cria botões de ação com estilo moderno"""
        button_frame = tk.Frame(parent, bg=LiquidGlassStyle.BG_CARD)
        button_frame.pack(pady=20)
        
        # Botão principal
        self.gerar_btn = tk.Button(
            button_frame,
            text="🎬 GERAR TEMPLATES",
            command=self._iniciar_processamento,
            bg=LiquidGlassStyle.ACCENT_PRIMARY,
            fg=LiquidGlassStyle.BG_PRIMARY,
            activebackground=LiquidGlassStyle.ACCENT_SECONDARY,
            activeforeground=LiquidGlassStyle.TEXT_PRIMARY,
            relief="flat",
            font=("Segoe UI", 12, "bold"),
            cursor="hand2",
            bd=0,
            padx=30,
            pady=12
        )
        self.gerar_btn.pack(side=tk.LEFT, padx=5)
        
        # Botão secundário
        clear_btn = tk.Button(
            button_frame,
            text="🔄 LIMPAR",
            command=self._limpar_campos,
            bg=LiquidGlassStyle.BG_HOVER,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            activebackground=LiquidGlassStyle.BG_SECONDARY,
            relief="flat",
            font=("Segoe UI", 10),
            cursor="hand2",
            bd=0,
            padx=20,
            pady=10
        )
        clear_btn.pack(side=tk.LEFT, padx=5)

    def _create_info_card(self, parent):
        """Cria card de informações"""
        info_card = tk.Frame(parent, bg=LiquidGlassStyle.BG_SECONDARY, relief="flat")
        info_card.pack(fill=tk.X, pady=10)
        
        content = tk.Frame(info_card, bg=LiquidGlassStyle.BG_SECONDARY)
        content.pack(fill=tk.X, padx=15, pady=15)
        
        title = tk.Label(
            content,
            text="ℹ️  Informações Importantes",
            bg=LiquidGlassStyle.BG_SECONDARY,
            fg=LiquidGlassStyle.ACCENT_PRIMARY,
            font=("Segoe UI", 11, "bold")
        )
        title.pack(anchor="w", pady=(0, 10))
        
        info_text = """• O CSV deve conter uma coluna 'Nome' com os casos de teste
• Template é opcional - será criado automaticamente se não fornecido
• Os campos personalizados serão preenchidos em todos os documentos
• Suporte automático a múltiplos encodings (UTF-8, Latin-1, etc.)
• Geração robusta com fallback automático em caso de erros"""
        
        info_label = tk.Label(
            content,
            text=info_text,
            bg=LiquidGlassStyle.BG_SECONDARY,
            fg=LiquidGlassStyle.TEXT_SECONDARY,
            font=("Segoe UI", 9),
            justify="left"
        )
        info_label.pack(anchor="w")

    def _create_fields_tab(self, parent):
        """Cria aba de campos com scroll"""
        container = tk.Frame(parent, bg=LiquidGlassStyle.BG_CARD)
        container.pack(fill=tk.BOTH, expand=True)
        
        # Canvas para scroll
        canvas = tk.Canvas(container, bg=LiquidGlassStyle.BG_CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=LiquidGlassStyle.BG_CARD)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Título
        title_frame = tk.Frame(scrollable_frame, bg=LiquidGlassStyle.BG_CARD)
        title_frame.pack(fill=tk.X, padx=20, pady=20)
        
        title_label = tk.Label(
            title_frame,
            text="📝 Dados para Preenchimento",
            bg=LiquidGlassStyle.BG_CARD,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(anchor="w")
        
        subtitle = tk.Label(
            title_frame,
            text="Preencha os campos abaixo que serão aplicados a todos os documentos",
            bg=LiquidGlassStyle.BG_CARD,
            fg=LiquidGlassStyle.TEXT_SECONDARY,
            font=("Segoe UI", 9)
        )
        subtitle.pack(anchor="w", pady=(5, 0))
        
        # Campos dinâmicos
        for i, campo_info in enumerate(self.campos_config):
            self._create_field_entry(scrollable_frame, campo_info, i)
        
        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y", pady=20)

    def _create_field_entry(self, parent, campo_info, index):
        """Cria entrada de campo individual"""
        field_card = tk.Frame(parent, bg=LiquidGlassStyle.BG_SECONDARY, relief="flat")
        field_card.pack(fill=tk.X, padx=20, pady=8)
        
        content = tk.Frame(field_card, bg=LiquidGlassStyle.BG_SECONDARY)
        content.pack(fill=tk.X, padx=15, pady=12)
        
        # Label do campo
        label = tk.Label(
            content,
            text=campo_info['label'],
            bg=LiquidGlassStyle.BG_SECONDARY,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            font=("Segoe UI", 10, "bold"),
            width=15,
            anchor="w"
        )
        label.pack(side=tk.LEFT, padx=(0, 15))
        
        # Entry do campo
        entry = tk.Entry(
            content,
            bg=LiquidGlassStyle.BG_PRIMARY,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            insertbackground=LiquidGlassStyle.ACCENT_PRIMARY,
            relief="flat",
            font=("Segoe UI", 10),
            bd=0
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, ipadx=10)
        
        self.campos_entries[campo_info['key']] = entry

    def _create_log_tab(self, parent):
        """Cria aba de log com design moderno"""
        container = tk.Frame(parent, bg=LiquidGlassStyle.BG_CARD)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title_label = tk.Label(
            container,
            text="📋 Log de Execução",
            bg=LiquidGlassStyle.BG_CARD,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(anchor="w", pady=(0, 15))
        
        # Log text com estilo
        log_frame = tk.Frame(container, bg=LiquidGlassStyle.BG_SECONDARY, relief="flat")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            bg=LiquidGlassStyle.BG_PRIMARY,
            fg=LiquidGlassStyle.TEXT_PRIMARY,
            insertbackground=LiquidGlassStyle.ACCENT_PRIMARY,
            relief="flat",
            font=("Consolas", 9),
            bd=0,
            padx=10,
            pady=10
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Barra de progresso
        progress_frame = tk.Frame(container, bg=LiquidGlassStyle.BG_CARD)
        progress_frame.pack(fill=tk.X, pady=(15, 0))
        
        progress_label = tk.Label(
            progress_frame,
            text="Progresso:",
            bg=LiquidGlassStyle.BG_CARD,
            fg=LiquidGlassStyle.TEXT_SECONDARY,
            font=("Segoe UI", 10)
        )
        progress_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            style="Glass.Horizontal.TProgressbar"
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _select_csv(self):
        """Seleciona arquivo CSV"""
        filename = filedialog.askopenfilename(
            title="Selecione o arquivo CSV",
            filetypes=[("CSV Files", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.csv_var.set(filename)
            self._log(f"📊 CSV selecionado: {os.path.basename(filename)}")

    def _select_template(self):
        """Seleciona template DOCX"""
        filename = filedialog.askopenfilename(
            title="Selecione o template DOCX",
            filetypes=[("Documentos Word", "*.docx"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.template_var.set(filename)
            self._log(f"📄 Template selecionado: {os.path.basename(filename)}")

    def _select_output(self):
        """Seleciona diretório de saída"""
        directory = filedialog.askdirectory(title="Selecione o diretório de saída")
        if directory:
            self.output_var.set(directory)
            self._log(f"📁 Diretório de saída: {directory}")

    def _limpar_campos(self):
        """Limpa todos os campos"""
        self.csv_var.set("")
        self.template_var.set("")
        self.output_var.set(self.settings.get('output_dir', 'evidencias_geradas'))
        
        for entry in self.campos_entries.values():
            entry.delete(0, tk.END)
            
        self._clear_log()
        if self.progress:
            self.progress['value'] = 0
            
        self._log("🔄 Campos limpos")

    def _clear_log(self):
        """Limpa o log"""
        if self.log_text:
            self.log_text.delete(1.0, tk.END)

    def _log(self, mensagem):
        """Adiciona mensagem ao log com cores"""
        if self.log_text:
            self.log_text.insert(tk.END, mensagem + "\n")
            self.log_text.see(tk.END)
            if self.window:
                self.window.update()

    def _iniciar_processamento(self):
        """Inicia processamento em thread separada"""
        if not self._validar_entradas():
            return
            
        self.gerar_btn.config(state="disabled")
        self._clear_log()
        self.progress['value'] = 0
        
        thread = threading.Thread(target=self._processar_documentos)
        thread.daemon = True
        thread.start()

    def _validar_entradas(self):
        """Valida as entradas do usuário"""
        if not self.csv_var.get():
            messagebox.showerror("Erro", "Selecione um arquivo CSV!")
            return False
            
        if not os.path.exists(self.csv_var.get()):
            messagebox.showerror("Erro", "Arquivo CSV não encontrado!")
            return False
            
        if self.template_var.get() and not os.path.exists(self.template_var.get()):
            messagebox.showwarning("Aviso", "Template não encontrado. Será criado automaticamente.")
            
        # Valida campos obrigatórios
        try:
            dados_fixos = self._get_dados_fixos()
        except ValueError as e:
            messagebox.showerror("Erro", str(e))
            return False
            
        return True

    def _get_dados_fixos(self):
        """Obtém dados dos campos fixos"""
        dados = {}
        for campo_key, entry in self.campos_entries.items():
            valor = entry.get().strip()
            if not valor:
                raise ValueError(f"O campo '{campo_key}' é obrigatório!")
            dados[campo_key] = valor
        return dados

    def _processar_documentos(self):
        """Processa os documentos em lote"""
        try:
            dados_fixos = self._get_dados_fixos()
            
            csv_path = self.csv_var.get()
            template_path = self.template_var.get().strip()
            output_folder = self.output_var.get().strip() or 'evidencias_geradas'
            
            self._log("🚀 INICIANDO PROCESSAMENTO EM LOTE")
            self._log("=" * 60)
            
            # Garantir template válido
            template_path = self._garantir_template_valido(template_path)
            
            # Criar pasta de saída
            try:
                Path(output_folder).mkdir(exist_ok=True)
                self._log(f"📁 Pasta de saída: {output_folder}")
            except Exception as e:
                self._log(f"⚠️ Aviso: Não foi possível criar a pasta '{output_folder}': {e}")
                self._log("📂 Usando pasta atual...")
                output_folder = '.'
            
            # Ler CSV
            self._log("📖 Lendo arquivo CSV...")
            casos_teste = self.csv_reader.read_csv(csv_path)
            
            if not casos_teste:
                self._log("❌ Nenhum caso de teste encontrado no CSV")
                messagebox.showerror("Erro", "Não foi possível ler os casos de teste do CSV")
                self.gerar_btn.config(state="normal")
                return
            
            self._log(f"📊 Encontrados {len(casos_teste)} casos de teste\n")
            
            # Determinar modo de operação
            use_default_template = not (template_path and Path(template_path).exists())
            
            if use_default_template:
                self._log("📝 Gerando documentos com template padrão...")
            else:
                self._log("📄 Usando template personalizado...")
            
            # Processar cada caso
            self._processar_casos_teste(casos_teste, dados_fixos, template_path, 
                                       output_folder, use_default_template)
            
        except Exception as e:
            self._log(f"❌ Erro inesperado: {e}")
            messagebox.showerror("Erro", f"Erro inesperado: {e}")
        finally:
            self.gerar_btn.config(state="normal")

    def _garantir_template_valido(self, template_path):
        """Garante que temos um template válido"""
        if not template_path or not Path(template_path).exists():
            self._log("🔍 Nenhum template válido encontrado, criando automaticamente...")
            if TemplateGenerator.create_example_template(self.campos_config):
                new_template_path = 'template_evidencias.docx'
                self.template_var.set(new_template_path)
                self._log("✅ Template padrão criado e configurado!")
                return new_template_path
            else:
                self._log("⚠️ Não foi possível criar template, usando geração padrão...")
                return ""
        return template_path

    def _processar_casos_teste(self, casos_teste, dados_fixos, template_path, 
                              output_folder, use_default_template):
        """Processa cada caso de teste individualmente"""
        self.progress['maximum'] = len(casos_teste)
        
        sucessos = 0
        erros = []
        arquivos_gerados = set()
        
        # Criar mapeamento de campos
        field_mapping = {campo['key']: campo['label'].rstrip(':').strip() 
                        for campo in self.campos_config}
        
        campo_nome = next(iter(dados_fixos.keys()))
        
        for i, caso_teste in enumerate(casos_teste, 1):
            try:
                self.progress['value'] = i
                self._log(f"📄 [{i}/{len(casos_teste)}] Processando: {caso_teste}")
                
                if self._gerar_documento_individual(caso_teste, dados_fixos, template_path,
                                                   output_folder, field_mapping, campo_nome,
                                                   arquivos_gerados, use_default_template, i):
                    sucessos += 1
                else:
                    erros.append(caso_teste)
                    
            except Exception as e:
                self._log(f"❌ Erro no caso '{caso_teste}': {e}")
                erros.append(caso_teste)
        
        # Resultado final
        self._mostrar_resultado_final(sucessos, len(erros), len(casos_teste), 
                                     output_folder, arquivos_gerados)

    def _gerar_documento_individual(self, caso_teste, dados_fixos, template_path,
                                   output_folder, field_mapping, campo_nome,
                                   arquivos_gerados, use_default_template, index):
        """Gera um documento individual"""
        try:
            dados_completos = dados_fixos.copy()
            dados_completos['Caso de Teste'] = caso_teste
            
            # Usar template ou criar padrão
            if not use_default_template and template_path:
                try:
                    doc = Document(template_path)
                    self.doc_processor.fill_template(doc, dados_completos, field_mapping)
                except Exception as e:
                    self._log(f"⚠️ Erro ao usar template: {e}. Usando padrão...")
                    doc = self.default_doc_generator.create_default_document(
                        dados_completos, self.campos_config)
            else:
                doc = self.default_doc_generator.create_default_document(
                    dados_completos, self.campos_config)
            
            # Gerar nome único
            nome_base = self.doc_processor.clean_filename(caso_teste)
            nome_arquivo = self._gerar_nome_arquivo_unico(
                f"Evidencia_{dados_fixos[campo_nome]}_{nome_base}.docx",
                arquivos_gerados, index)
            
            caminho_completo = Path(output_folder) / nome_arquivo
            
            # Salvar documento
            try:
                doc.save(caminho_completo)
                self._log(f"✅ Salvo: {nome_arquivo}")
                return True
            except Exception as e:
                # Fallback com nome alternativo
                nome_alternativo = f"Evidencia_{index}_{datetime.now().strftime('%H%M%S')}.docx"
                caminho_alternativo = Path(output_folder) / nome_alternativo
                doc.save(caminho_alternativo)
                self._log(f"✅ Salvo (nome alternativo): {nome_alternativo}")
                arquivos_gerados.add(nome_alternativo)
                return True
                
        except Exception as e:
            self._log(f"❌ Erro crítico ao gerar: {e}")
            return False

    def _gerar_nome_arquivo_unico(self, nome_base, arquivos_gerados, fallback_index):
        """Gera um nome de arquivo único"""
        contador = 1
        nome_final = nome_base
        
        while nome_final in arquivos_gerados:
            base, ext = os.path.splitext(nome_base)
            nome_final = f"{base}_{contador}{ext}"
            contador += 1
        
        arquivos_gerados.add(nome_final)
        return nome_final

    def _mostrar_resultado_final(self, sucessos, erros, total, output_folder, arquivos_gerados):
        """Mostra resultado final do processamento"""
        self._log("\n" + "=" * 60)
        self._log("🎉 PROCESSAMENTO CONCLUÍDO!")
        self._log("=" * 60)
        self._log(f"📊 Total processado: {total}")
        self._log(f"✅ Sucessos: {sucessos}")
        self._log(f"❌ Erros: {erros}")
        self._log(f"📁 Pasta: {Path(output_folder).absolute()}")
        
        if sucessos > 0:
            self._log(f"\n📋 Arquivos gerados:")
            for arquivo in sorted(arquivos_gerados)[:10]:
                self._log(f"  • {arquivo}")
            if len(arquivos_gerados) > 10:
                self._log(f"  • ... e mais {len(arquivos_gerados) - 10} arquivos")
        
        if erros == 0:
            messagebox.showinfo(
                "Sucesso", 
                f"✅ Todos os {sucessos} documentos foram gerados com sucesso!\n\n"
                f"📁 Pasta: {Path(output_folder).absolute()}"
            )
        else:
            messagebox.showwarning(
                "Concluído com avisos",
                f"Processo concluído com {erros} erro(s).\n"
                f"✅ {sucessos} documentos gerados.\n\n"
                f"📁 Pasta: {Path(output_folder).absolute()}"
            )


# Exemplo de uso (para testes independentes)
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    
    settings = {
        'template_dir': '.',
        'output_dir': 'evidencias_geradas'
    }
    
    module = TemplateGeneratorModule(root, settings)
    module.show()
    
    root.mainloop()