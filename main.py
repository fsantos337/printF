# main.py
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
from datetime import datetime

class PrintFApp:
    """Aplicação principal unificada completa"""
    
    def __init__(self):
        self.root = tk.Tk()
        
        # Carregar configurações PRIMEIRO
        self.settings = self._load_settings()
        
        # Configurar janela principal
        self._setup_main_window()
        
        # Módulos
        self.modules = {}
        self.current_module = None
        
        # Configurar estilo
        self._setup_styles()
        
        # Bindings para responsividade
        self._setup_bindings()
        
        # Variáveis para controle de responsividade
        self.current_padding = "30"
        self.current_font_scale = 1.0
        self.current_cols = 2
        
    def _setup_main_window(self):
        """Configura a janela principal"""
        self.root.title("PrintF - Sistema Completo de Evidências")
        
        # Usar tamanho salvo ou padrão
        width = self.settings.get('window_size', {}).get('width', 904)
        height = self.settings.get('window_size', {}).get('height', 600)
        self.root.geometry(f"{width}x{height}")
        
        # Posição salva ou centralizada
        if 'window_position' in self.settings:
            x = self.settings['window_position']['x']
            y = self.settings['window_position']['y']
            self.root.geometry(f"+{x}+{y}")
        else:
            self.root.update_idletasks()
            largura = self.root.winfo_width()
            largura_tela = self.root.winfo_screenwidth()
        
            x = (largura_tela - largura) // 2
            y = 30  # Margem do topo
            self.root.geometry(f"+{x}+{y}")
        
        # Tamanho mínimo
        from config import APP_CONFIG
        self.root.minsize(
            APP_CONFIG.UI_SETTINGS['min_width'], 
            APP_CONFIG.UI_SETTINGS['min_height']
        )
        
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
        """Configura estilos visuais baseados no tema selecionado"""
        self.style_manager = None
        self.using_liquid_glass = False
        
        try:
            # Tentar importar do módulo styles
            from modules.styles import LiquidGlassStyle
            
            # Verificar se o tema está habilitado nas configurações
            theme_to_use = self.settings.get('theme', 'liquid_glass')
            if theme_to_use == 'liquid_glass':
                # Aplicar estilo Liquid Glass
                LiquidGlassStyle.apply_window_style(self.root)
                self.style_manager = LiquidGlassStyle
                self.using_liquid_glass = True
                
                # Configurar fonte fallback para sistemas sem SF Pro
                default_fonts = ['Segoe UI', 'Arial', 'Helvetica']
                self._configure_font_fallback(default_fonts)
                print("✅ Estilo Liquid Glass aplicado com sucesso!")
            else:
                self._setup_fallback_styles()
                print(f"ℹ️ Usando estilo padrão (tema: {theme_to_use})")
            
        except ImportError as e:
            # Fallback para estilo padrão
            print(f"⚠️ Liquid Glass não disponível: {e}")
            self._setup_fallback_styles()
        except Exception as e:
            print(f"⚠️ Erro ao aplicar Liquid Glass: {e}")
            self._setup_fallback_styles()

    def _configure_font_fallback(self, font_list):
        """Configura fallback de fontes"""
        for font_name in font_list:
            try:
                test_font = tk.font.Font(family=font_name, size=10)
                if test_font.actual()['family'] == font_name:
                    # Fonte disponível, usar como padrão
                    self.root.option_add('*Font', (font_name, 10))
                    break
            except:
                continue

    def _setup_fallback_styles(self):
        """Configura estilos fallback"""
        self.style = ttk.Style()
        
        # Configurar tema
        try:
            self.style.theme_use('clam')
        except:
            pass
        
        # Estilos personalizados fallback
        self.style.configure('Title.TLabel', 
                           font=('Arial', 16, 'bold'), 
                           foreground='#2c3e50',
                           background='#f5f5f5')
        
        self.style.configure('Module.TButton',
                           font=('Arial', 11, 'bold'),
                           padding=(15, 10))
        
        self.style.configure('Accent.TButton',
                           background='#3498db',
                           foreground='white',
                           focuscolor='none')
        
        # Configurar cores de fallback
        self.root.configure(bg='#f5f5f5')

    def _setup_bindings(self):
        """Configura bindings para responsividade"""
        self.root.bind('<Configure>', self._on_window_resize)
        
    def _on_window_resize(self, event):
        """Manipula redimensionamento da janela"""
        if event.widget == self.root and hasattr(self, 'main_frame'):
            # Atualizar layout responsivo
            self._update_responsive_layout(event.width, event.height)

    def _update_responsive_layout(self, width, height):
        """Atualiza layout baseado no tamanho da tela"""
        from config import APP_CONFIG
        breakpoints = APP_CONFIG.UI_SETTINGS['responsive_breakpoints']
        
        # Determinar configurações baseadas no tamanho
        if width < breakpoints['small']:
            # Layout mobile/tablet
            new_padding = "10"
            new_font_scale = 0.85
            new_cols = 1
            title_font_size = 14
            subtitle_font_size = 9
            module_font_size = 10
            wrap_length = 300
        elif width < breakpoints['medium']:
            # Layout pequeno desktop
            new_padding = "20"
            new_font_scale = 0.95
            new_cols = 2
            title_font_size = 15
            subtitle_font_size = 10
            module_font_size = 11
            wrap_length = 180
        else:
            # Layout grande desktop
            new_padding = "30"
            new_font_scale = 1.0
            new_cols = 2
            title_font_size = 16
            subtitle_font_size = 11
            module_font_size = 12
            wrap_length = 200
        
        # Aplicar mudanças se necessário
        needs_update = False
        
        if (hasattr(self, 'current_padding') and self.current_padding != new_padding or
            hasattr(self, 'current_font_scale') and self.current_font_scale != new_font_scale or
            hasattr(self, 'current_cols') and self.current_cols != new_cols):
            
            self.current_padding = new_padding
            self.current_font_scale = new_font_scale
            self.current_cols = new_cols
            needs_update = True
        
        # Atualizar UI se necessário
        if needs_update and hasattr(self, 'main_frame'):
            self._refresh_ui_layout(title_font_size, subtitle_font_size, module_font_size, wrap_length)

    def _refresh_ui_layout(self, title_font_size, subtitle_font_size, module_font_size, wrap_length):
        """Atualiza dinamicamente o layout da UI"""
        try:
            # Atualizar padding do frame principal
            self.main_frame.pack_configure(padx=self.current_padding, pady=self.current_padding)
            
            # Atualizar header
            if hasattr(self, 'title_text'):
                if self.using_liquid_glass:
                    self.title_text.configure(font=("Arial", title_font_size, "bold"))
                else:
                    self.title_text.configure(font=("Arial", title_font_size, "bold"))
            
            if hasattr(self, 'subtitle'):
                if self.using_liquid_glass:
                    self.subtitle.configure(font=("Arial", subtitle_font_size))
                else:
                    self.subtitle.configure(font=("Arial", subtitle_font_size))
            
            # Atualizar grid de módulos
            if hasattr(self, 'modules_frame') and hasattr(self, 'modules_config'):
                self._create_responsive_grid(self.modules_frame, self.modules_config, wrap_length)
                
            # Atualizar footer
            if hasattr(self, 'footer_text'):
                if self.using_liquid_glass:
                    pass  # Mantém estilo padrão para footer
                else:
                    self.footer_text.configure(font=("Arial", 8))
            
        except Exception as e:
            print(f"⚠️ Erro ao atualizar layout: {e}")

    def _load_settings(self):
        """Carrega configurações"""
        from config import APP_CONFIG
        return APP_CONFIG.load_user_settings()

    def _save_settings(self):
        """Salva configurações"""
        # Salvar tamanho e posição atual
        self.settings['window_size'] = {
            'width': self.root.winfo_width(),
            'height': self.root.winfo_height()
        }
        self.settings['window_position'] = {
            'x': self.root.winfo_x(),
            'y': self.root.winfo_y()
        }
        
        from config import APP_CONFIG
        APP_CONFIG.save_user_settings(self.settings)

    def create_ui(self):
        """Cria interface principal completa"""
        # Frame principal
        if self.using_liquid_glass and self.style_manager:
            # Usar estilo Liquid Glass
            self.main_frame = self.style_manager.create_glass_frame(self.root)
        else:
            # Fallback
            self.main_frame = ttk.Frame(self.root, padding=self.current_padding)
            self.main_frame.configure(style='TFrame')
            
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Cabeçalho
        self._create_header(self.main_frame)
        
        # Separador
        if self.using_liquid_glass and self.style_manager:
            separator = ttk.Separator(self.main_frame, orient='horizontal', style="Glass.TSeparator")
        else:
            separator = ttk.Separator(self.main_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=20)
        
        # Módulos
        self._create_modules_grid(self.main_frame)
        
        # Rodapé
        self._create_footer(self.main_frame)
        
        # Forçar atualização inicial do layout
        self.root.update()
        self._update_responsive_layout(self.root.winfo_width(), self.root.winfo_height())

    def _create_header(self, parent):
        """Cria cabeçalho da aplicação"""
        if self.using_liquid_glass and self.style_manager:
            header_frame = self.style_manager.create_glass_frame(parent)
        else:
            header_frame = ttk.Frame(parent)
            header_frame.configure(style='TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Logo e título
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(fill=tk.X)
        
        # 🔥 CORREÇÃO: Verificar e carregar logo customizada de forma relativa ao executável
        logo_loaded = False

        # Determinar o diretório base correto (funciona tanto no desenvolvimento quanto no EXE)
        if getattr(sys, 'frozen', False):
            # Se estiver rodando como executável (PyInstaller)
            base_dir = os.path.dirname(sys.executable)
        else:
            # Se estiver rodando como script Python
            base_dir = os.path.dirname(__file__)

        # Procurar a logo em vários locais possíveis
        possible_logo_paths = [
            os.path.join(base_dir, "CUSTOM-LOGO.PNG"),
            os.path.join(base_dir, "assets", "CUSTOM-LOGO.PNG"),
            os.path.join(base_dir, "images", "CUSTOM-LOGO.PNG"),
            os.path.join(os.getcwd(), "CUSTOM-LOGO.PNG")  # Diretório atual de trabalho
        ]

        custom_logo_path = None
        for path in possible_logo_paths:
            if os.path.exists(path):
                custom_logo_path = path
                break
        
        if custom_logo_path and os.path.exists(custom_logo_path):
            try:
                from PIL import Image, ImageTk
                
                # Carregar e redimensionar imagem
                pil_image = Image.open(custom_logo_path)
                # Redimensionar para 100x100 mantendo proporção
                pil_image.thumbnail((100, 100), Image.Resampling.LANCZOS)
                
                # Converter para PhotoImage do Tkinter
                self.custom_logo = ImageTk.PhotoImage(pil_image)
                
                # Criar label com imagem
                if self.using_liquid_glass and self.style_manager:
                    icon_label = ttk.Label(title_frame, image=self.custom_logo, style="Glass.TLabel")
                else:
                    icon_label = ttk.Label(title_frame, image=self.custom_logo, background='#f5f5f5')
                
                icon_label.pack(side=tk.LEFT, padx=(0, 10))
                logo_loaded = True
                print(f"✅ Logo customizada carregada: {custom_logo_path}")
                
            except Exception as e:
                print(f"⚠️ Erro ao carregar logo customizada: {e}")
                logo_loaded = False
        
        # Fallback para emoji se logo customizada não carregar
        if not logo_loaded:
            if self.using_liquid_glass and self.style_manager:
                icon_label = ttk.Label(title_frame, text="🖨️", font=("Arial", 24), style="Glass.TLabel")
            else:
                icon_label = ttk.Label(title_frame, text="🖨️", font=("Arial", 24), background='#f5f5f5')
            icon_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # Textos
        if self.using_liquid_glass and self.style_manager:
            self.title_text = self.style_manager.create_title_label(title_frame, "PRINTF")
        else:
            self.title_text = ttk.Label(title_frame, text="PRINTF", style='Title.TLabel')
        self.title_text.pack(side=tk.LEFT)
        
        if self.using_liquid_glass and self.style_manager:
            self.subtitle = ttk.Label(header_frame, 
                               text="Sistema Completo de Captura e Documentação de Evidências",
                               style="Subtitle.TLabel")
        else:
            self.subtitle = ttk.Label(header_frame, 
                               text="Sistema Completo de Captura e Documentação de Evidências",
                               style="Subtitle.TLabel" if hasattr(self, 'style') else None,
                               font=("Arial", 10))
        self.subtitle.pack(pady=(5, 0))
        
        # Versão
        from config import APP_CONFIG
        if self.using_liquid_glass and self.style_manager:
            version_text = ttk.Label(header_frame, 
                                   text=f"Versão {APP_CONFIG.VERSION}",
                                   style="Subtitle.TLabel")
        else:
            version_text = ttk.Label(header_frame, 
                                   text=f"Versão {APP_CONFIG.VERSION}",
                                   style="Subtitle.TLabel" if hasattr(self, 'style') else None,
                                   font=("Arial", 8))
        version_text.pack(side=tk.RIGHT)

    def _create_modules_grid(self, parent):
        """Cria grid de módulos responsivo"""
        if self.using_liquid_glass and self.style_manager:
            self.modules_frame = self.style_manager.create_glass_frame(parent)
        else:
            self.modules_frame = ttk.Frame(parent)
            self.modules_frame.configure(style='TFrame')
        self.modules_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configuração dos módulos (SEM HOTKEYS)
        self.modules_config = [
            {
                "title": "📷 CAPTURAR EVIDÊNCIAS",
                "key": "capture",
                "color": "#27ae60",
                "description": "Capture screenshots com um clique\n• Multi-monitor\n• Timestamp automático\n• Metadados completos"
            },
            {
                "title": "📄 GERAR TEMPLATES", 
                "key": "templates",
                "color": "#3498db",
                "description": "Crie documentos em lote\n• Templates personalizados\n• CSV automático\n• Campos dinâmicos"
            },
            {
                "title": "📋 GERAR DOCUMENTOS",
                "key": "evidence", 
                "color": "#f39c12",
                "description": "Converta evidências em DOCX\n• Navegação avançada\n• Editor completo\n• Comentários"
            },
            {
                "title": "🗑️ LIMPAR ARQUIVOS",
                "key": "cleanup",
                "color": "#e74c3c", 
                "description": "Gerencie e limpe arquivos\n• Análise detalhada\n• Filtros inteligentes\n• Exclusão segura"
            }
        ]
        
        # Criar grid responsivo inicial
        self._create_responsive_grid(self.modules_frame, self.modules_config, 200)

    def _create_responsive_grid(self, parent, modules_config, wrap_length=200):
        """Cria grid responsivo baseado no tamanho da tela"""
        width = self.root.winfo_width()
        
        # Limpar grid anterior se existir
        for widget in parent.winfo_children():
            widget.destroy()
        
        # Criar novo grid
        for i, module in enumerate(modules_config):
            if self.current_cols == 1:
                # Layout single column - usa pack
                card = self._create_module_card(parent, module, wrap_length)
                card.pack(fill=tk.X, padx=5, pady=5)
            else:
                # Layout grid - usa grid
                row = i // self.current_cols
                col = i % self.current_cols
                
                card = self._create_module_card(parent, module, wrap_length)
                card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # Configurar weights para expansão responsiva
        if self.current_cols > 1:
            row_count = (len(modules_config) + self.current_cols - 1) // self.current_cols
            for i in range(row_count):
                parent.grid_rowconfigure(i, weight=1)
            for j in range(self.current_cols):
                parent.grid_columnconfigure(j, weight=1)

    def _create_module_card(self, parent, module_config, wrap_length):
        """Cria card de módulo individual com wrap length dinâmico"""
        if self.using_liquid_glass and self.style_manager:
            card_frame = self.style_manager.create_card(parent)
        else:
            card_frame = tk.Frame(parent, relief="solid", borderwidth=1, bg='white')
        
        # Configurar card para expandir
        card_frame.pack_propagate(False) if self.current_cols > 1 else card_frame.pack_propagate(True)
        
        # Botão principal
        if self.using_liquid_glass and self.style_manager:
            btn = self.style_manager.create_accent_button(
                card_frame, 
                module_config["title"],
                command=lambda k=module_config["key"]: self.open_module(k)
            )
            btn.pack(fill=tk.X, padx=8, pady=8)
        else:
            btn = tk.Button(card_frame, 
                           text=module_config["title"],
                           command=lambda k=module_config["key"]: self.open_module(k),
                           bg=module_config["color"],
                           fg="white",
                           font=("Arial", 11, "bold"),
                           relief="flat",
                           cursor="hand2",
                           anchor="w",
                           justify="left")
            btn.pack(fill=tk.X, padx=8, pady=8)
            
            # Efeitos hover para estilo padrão
            btn.bind("<Enter>", lambda e, b=btn, c=module_config["color"]: 
                    b.config(bg=self._darken_color(c)))
            btn.bind("<Leave>", lambda e, b=btn, c=module_config["color"]: 
                    b.config(bg=c))
        
        # Descrição com wrap length dinâmico
        if self.using_liquid_glass and self.style_manager:
            desc_label = ttk.Label(card_frame, 
                                 text=module_config["description"],
                                 style="Glass.TLabel",
                                 justify="left",
                                 wraplength=wrap_length)
        else:
            desc_label = tk.Label(card_frame, 
                                 text=module_config["description"],
                                 font=("Arial", 9),
                                 bg="white",
                                 fg="#2c3e50",
                                 justify="left",
                                 wraplength=wrap_length)
        desc_label.pack(fill=tk.X, padx=8, pady=(0, 8))

        return card_frame

    def _create_footer(self, parent):
        """Cria rodapé"""
        if self.using_liquid_glass and self.style_manager:
            footer_frame = self.style_manager.create_glass_frame(parent)
        else:
            footer_frame = tk.Frame(parent, bg='#f5f5f5')
        footer_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Informações do sistema
        sys_info = tk.Frame(footer_frame, 
                           bg=self.style_manager.BG_CARD if self.using_liquid_glass else '#f5f5f5')
        sys_info.pack(side=tk.LEFT)
        
        if self.using_liquid_glass and self.style_manager:
            self.footer_text = ttk.Label(sys_info, 
                                  text=f"©PrintF Unificado • {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                                  style="Subtitle.TLabel")
        else:
            self.footer_text = tk.Label(sys_info, 
                                  text=f"©PrintF Unificado • {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                                  font=("Arial", 8),
                                  foreground="gray",
                                  bg='#f5f5f5')
        self.footer_text.pack(anchor="w")
        
        # Ações globais
        actions = tk.Frame(footer_frame, 
                          bg=self.style_manager.BG_CARD if self.using_liquid_glass else '#f5f5f5')
        actions.pack(side=tk.RIGHT)
        
        if self.using_liquid_glass and self.style_manager:
            
            ttk.Button(actions, text="❓ Ajuda", style="Glass.TButton",
                      command=self._show_help).pack(side=tk.LEFT, padx=2)
            
            ttk.Button(actions, text="❌ Sair", style="Back.TButton",
                      command=self._on_closing).pack(side=tk.LEFT, padx=2)
        else:
            tk.Button(actions, text="⚙️ Configurações",
                     command=self._open_settings,
                     bg='#3498db', fg='white', font=("Arial", 9),
                     relief="flat").pack(side=tk.LEFT, padx=2)
            
            tk.Button(actions, text="❓ Ajuda",
                     command=self._show_help,
                     bg='#95a5a6', fg='white', font=("Arial", 9),
                     relief="flat").pack(side=tk.LEFT, padx=2)
            
            tk.Button(actions, text="❌ Sair",
                     command=self._on_closing,
                     bg='#e74c3c', fg='white', font=("Arial", 9),
                     relief="flat").pack(side=tk.LEFT, padx=2)

    def _darken_color(self, color, factor=0.8):
        """Escurece cor hexadecimal (simplificado)"""
        # Conversão simples para escurecer cor
        try:
            if color.startswith('#'):
                r = int(color[1:3], 16)
                g = int(color[3:5], 16)
                b = int(color[5:7], 16)
                r = int(r * factor)
                g = int(g * factor)
                b = int(b * factor)
                return f'#{r:02x}{g:02x}{b:02x}'
        except:
            pass
        return color

    def open_module(self, module_key):
        """Abre um módulo específico - CORREÇÃO COMPLETA E MELHORADA"""
        # Fecha módulo atual se existir
        if self.current_module:
            self.current_module.hide()
        
        # Importa e cria módulo dinamicamente
        if module_key not in self.modules:
            try:
                module = self._create_module(module_key)
                if module is None:
                    return
                self.modules[module_key] = module
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao carregar módulo {module_key}: {e}")
                return
        
        # Abre novo módulo
        self.current_module = self.modules[module_key]
        
        # 🔥 CORREÇÃO CRÍTICA: Sequência correta para garantir que o módulo fique visível
        try:
            # Primeiro mostra o módulo
            self.current_module.show()
            
            # 🔥 NOVA CORREÇÃO: Aguardar um pouco para garantir renderização completa
            self.root.after(500, self._focus_module_only)  # Aumentei o tempo para 500ms
            
        except Exception as e:
            print(f"❌ Erro ao abrir módulo {module_key}: {e}")
            # Tentar recriar o módulo
            try:
                module = self._create_module(module_key)
                if module:
                    self.modules[module_key] = module
                    self.current_module = module
                    self.current_module.show()
                    self.root.after(500, self._focus_module_only)
                else:
                    messagebox.showerror("Erro", f"Falha crítica ao abrir {module_key}")
                    return
            except Exception as e2:
                messagebox.showerror("Erro", f"Falha crítica ao abrir {module_key}: {e2}")
                return

    # 🔥 NOVO MÉTODO: Foca apenas no módulo SEM minimizar a janela principal
    def _focus_module_only(self):
        """Apenas foca no módulo sem minimizar a janela principal"""
        try:
            if self.current_module and hasattr(self.current_module, 'root'):
                # Garante que o módulo está visível e com foco
                self.current_module.root.deiconify()
                self.current_module.root.lift()
                self.current_module.root.focus_force()
                self.current_module.root.attributes('-topmost', True)
                
                # Remove o topmost após um breve período
                self.current_module.root.after(1000, lambda: 
                    self.current_module.root.attributes('-topmost', False) 
                    if hasattr(self.current_module, 'root') else None)
                
        except Exception as e:
            print(f"⚠️ Erro ao focar módulo: {e}")

    def _create_module(self, module_key):
        """Cria módulo dinamicamente"""
        try:
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
        except Exception as e:
            print(f"❌ Erro ao criar módulo {module_key}: {e}")
            messagebox.showerror("Erro", f"Falha ao carregar módulo {module_key}: {e}")
            return None

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