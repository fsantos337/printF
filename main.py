import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
from datetime import datetime

class PrintFApp:
    """Aplicação principal unificada completa"""
    
    def __init__(self):
        self.root = tk.Tk()
        self._setup_main_window()
        
        # Módulos
        self.modules = {}
        self.current_module = None
        
        # Configurações
        self.settings = self._load_settings()
        
        # Configurar estilo
        self._setup_styles()
        
    def _setup_main_window(self):
        """Configura a janela principal"""
        self.root.title("PrintF - Sistema Completo de Evidências")
        self.root.geometry("500x600")
        self.root.configure(bg='#f5f5f5')
        self.root.minsize(450, 550)
        
        # Ícone (se disponível)
        self._set_window_icon()
        
        # Protocolo de fechamento
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _set_window_icon(self):
        """Tenta configurar ícone da janela"""
        try:
            # Tenta carregar ícone se existir
            icon_path = os.path.join(os.path.dirname(__file__), "assets", "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass

    def _setup_styles(self):
        """Configura estilos visuais"""
        self.style = ttk.Style()
        
        # Configurar tema
        try:
            self.style.theme_use('clam')
        except:
            pass
        
        # Estilos personalizados
        self.style.configure('Title.TLabel', 
                           font=('Arial', 16, 'bold'), 
                           foreground='#2c3e50')
        
        self.style.configure('Module.TButton',
                           font=('Arial', 11, 'bold'),
                           padding=(15, 10))
        
        self.style.configure('Accent.TButton',
                           background='#3498db',
                           foreground='white',
                           focuscolor='none')

    def _load_settings(self):
        """Carrega configurações"""
        from config import APP_CONFIG
        return APP_CONFIG.load_user_settings()

    def _save_settings(self):
        """Salva configurações"""
        from config import APP_CONFIG
        APP_CONFIG.save_user_settings(self.settings)

    def create_ui(self):
        """Cria interface principal completa"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Cabeçalho
        self._create_header(main_frame)
        
        # Separador
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=20)
        
        # Módulos
        self._create_modules_grid(main_frame)
        
        # Rodapé
        self._create_footer(main_frame)

    def _create_header(self, parent):
        """Cria cabeçalho da aplicação"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Logo e título
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(fill=tk.X)
        
        # Ícone do aplicativo
        icon_label = ttk.Label(title_frame, text="🖨️", font=("Arial", 24))
        icon_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # Textos
        title_text = ttk.Label(title_frame, text="PRINTF UNIFICADO", 
                              style='Title.TLabel')
        title_text.pack(side=tk.LEFT)
        
        subtitle = ttk.Label(header_frame, 
                           text="Sistema Completo de Captura e Documentação de Evidências",
                           font=("Arial", 10), 
                           foreground="#7f8c8d")
        subtitle.pack(pady=(5, 0))
        
        # Versão
        from config import APP_CONFIG
        version_text = ttk.Label(header_frame, 
                               text=f"Versão {APP_CONFIG.VERSION}",
                               font=("Arial", 8),
                               foreground="#bdc3c7")
        version_text.pack(side=tk.RIGHT)

    def _create_modules_grid(self, parent):
        """Cria grid de módulos"""
        modules_frame = ttk.Frame(parent)
        modules_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configuração dos módulos
        modules_config = [
            {
                "title": "📷 CAPTURAR EVIDÊNCIAS",
                "key": "capture",
                "color": "#27ae60",
                "description": "Capture screenshots com um clique\n• Multi-monitor\n• Timestamp automático\n• Metadados completos",
                "hotkey": "F8"
            },
            {
                "title": "📄 GERAR TEMPLATES", 
                "key": "templates",
                "color": "#3498db",
                "description": "Crie documentos em lote\n• Templates personalizados\n• CSV automático\n• Campos dinâmicos",
                "hotkey": "F9"
            },
            {
                "title": "📋 GERAR DOCUMENTOS",
                "key": "evidence", 
                "color": "#f39c12",
                "description": "Converta evidências em DOCX\n• Navegação avançada\n• Editor completo\n• Comentários",
                "hotkey": "F10"
            },
            {
                "title": "🗑️ LIMPAR ARQUIVOS",
                "key": "cleanup",
                "color": "#e74c3c", 
                "description": "Gerencie e limpe arquivos\n• Análise detalhada\n• Filtros inteligentes\n• Exclusão segura",
                "hotkey": "F11"
            }
        ]
        
        # Criar grid 2x2
        for i, module in enumerate(modules_config):
            row = i // 2
            col = i % 2
            
            self._create_module_card(modules_frame, module, row, col)
        
        # Configurar grid
        modules_frame.grid_rowconfigure(0, weight=1)
        modules_frame.grid_rowconfigure(1, weight=1)
        modules_frame.grid_columnconfigure(0, weight=1)
        modules_frame.grid_columnconfigure(1, weight=1)

    def _create_module_card(self, parent, module_config, row, col):
        """Cria card de módulo individual"""
        card_frame = ttk.Frame(parent, relief="solid", borderwidth=1)
        card_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        card_frame.columnconfigure(0, weight=1)
        
        # Botão principal
        btn = tk.Button(card_frame, 
                       text=module_config["title"],
                       command=lambda k=module_config["key"]: self.open_module(k),
                       bg=module_config["color"],
                       fg="white",
                       font=("Arial", 11, "bold"),
                       width=20,
                       height=2,
                       relief="flat",
                       cursor="hand2",
                       anchor="w",
                       justify="left")
        btn.pack(fill=tk.X, padx=8, pady=8)
        
        # Hotkey
        hotkey_frame = ttk.Frame(card_frame)
        hotkey_frame.pack(fill=tk.X, padx=8)
        
        hotkey_label = ttk.Label(hotkey_frame, 
                                text=f"Atalho: {module_config['hotkey']}",
                                font=("Arial", 8, "bold"),
                                foreground=module_config["color"])
        hotkey_label.pack(side=tk.RIGHT)
        
        # Descrição
        desc_label = tk.Label(card_frame, 
                             text=module_config["description"],
                             font=("Arial", 9),
                             bg="white",
                             fg="#2c3e50",
                             justify="left",
                             wraplength=200)
        desc_label.pack(fill=tk.X, padx=8, pady=(0, 8))
        
        # Efeitos hover
        btn.bind("<Enter>", lambda e, b=btn, c=module_config["color"]: 
                b.config(bg=self._darken_color(c)))
        btn.bind("<Leave>", lambda e, b=btn, c=module_config["color"]: 
                b.config(bg=c))

    def _create_footer(self, parent):
        """Cria rodapé"""
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Informações do sistema
        sys_info = ttk.Frame(footer_frame)
        sys_info.pack(side=tk.LEFT)
        
        ttk.Label(sys_info, 
                 text=f"© 2024 PrintF Unificado • {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                 font=("Arial", 8),
                 foreground="gray").pack(anchor="w")
        
        # Ações globais
        actions = ttk.Frame(footer_frame)
        actions.pack(side=tk.RIGHT)
        
        ttk.Button(actions, text="⚙️ Configurações",
                  command=self._open_settings).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(actions, text="❓ Ajuda",
                  command=self._show_help).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(actions, text="❌ Sair",
                  command=self._on_closing).pack(side=tk.LEFT, padx=2)

    def _darken_color(self, color, factor=0.8):
        """Escurece cor hexadecimal (simplificado)"""
        return color

    def open_module(self, module_key):
        """Abre um módulo específico"""
        # Fecha módulo atual se existir
        if self.current_module:
            self.current_module.hide()
        
        # Importa e cria módulo dinamicamente
        if module_key not in self.modules:
            try:
                self.modules[module_key] = self._create_module(module_key)
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao carregar módulo: {e}")
                return
        
        # Abre novo módulo
        self.current_module = self.modules[module_key]
        self.current_module.show()

    def _create_module(self, module_key):
        """Cria módulo dinamicamente"""
        if module_key == "capture":
            from modules.capture import CaptureModule
            return CaptureModule(self.root, self.settings)
        elif module_key == "templates":
            from modules.template_gen import TemplateGeneratorModule
            return TemplateGeneratorModule(self.root, self.settings)
        elif module_key == "evidence":
            from modules.evidence_gen import EvidenceGeneratorModule
            return EvidenceGeneratorModule(self.root, self.settings)
        elif module_key == "cleanup":
            from modules.cleanup import CleanupModule
            return CleanupModule(self.root, self.settings)

    def _open_settings(self):
        """Abre configurações"""
        messagebox.showinfo("Configurações", "Painel de configurações em desenvolvimento!")

    def _show_help(self):
        """Mostra ajuda"""
        help_text = """
🖨️ PRINTF UNIFICADO - AJUDA RÁPIDA

📷 CAPTURAR EVIDÊNCIAS:
• F8 - Iniciar gravação
• F6 - Pausar/Retomar
• F9 - Finalizar gravação
• Clique para capturar screenshots

📄 GERAR TEMPLATES:
• Use CSV com coluna 'Nome'
• Templates DOCX personalizados
• Geração em lote automática

📋 GERAR DOCUMENTOS:
• Navegue entre evidências
• Adicione comentários
• Edite imagens com ferramentas
• Gere DOCX final

🗑️ LIMPAR ARQUIVOS:
• Analise pastas completas
• Filtre por tipo e tamanho
• Exclusão segura com confirmação

💡 DICAS:
• Mantenha suas evidências organizadas
• Use templates para padronização
• Faça backup regular dos documentos
        """
        messagebox.showinfo("Ajuda - PrintF Unificado", help_text)

    def _on_closing(self):
        """Manipula fechamento da aplicação"""
        # Fecha todos os módulos abertos
        for module in self.modules.values():
            if hasattr(module, 'hide'):
                module.hide()
        
        # Salva configurações
        self._save_settings()
        
        # Fecha aplicação
        self.root.quit()
        self.root.destroy()

    def run(self):
        """Executa a aplicação"""
        try:
            self.create_ui()
            
            # Centralizar na tela
            self.root.eval('tk::PlaceWindow . center')
            
            # Verificar dependências
            self._check_dependencies()
            
            self.root.mainloop()
            
        except Exception as e:
            messagebox.showerror("Erro Fatal", f"Falha ao iniciar aplicação: {e}")
            sys.exit(1)

    def _check_dependencies(self):
        """Verifica dependências críticas"""
        missing_deps = []
        
        try:
            from PIL import Image
        except ImportError:
            missing_deps.append("Pillow")
        
        try:
            from docx import Document
        except ImportError:
            missing_deps.append("python-docx")
        
        try:
            import pyautogui
        except ImportError:
            missing_deps.append("pyautogui")
        
        if missing_deps:
            messagebox.showwarning(
                "Dependências Ausentes",
                f"As seguintes bibliotecas não estão instaladas:\n\n"
                f"{', '.join(missing_deps)}\n\n"
                f"Algumas funcionalidades podem não estar disponíveis.\n"
                f"Execute: pip install {' '.join(missing_deps)}"
            )

if __name__ == "__main__":
    # Configurar caminho para imports
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    # Criar e executar aplicação
    app = PrintFApp()
    app.run()