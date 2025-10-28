# styles.py

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font as tkfont

class LiquidGlassStyle:
    """Estilo Liquid Glass para a aplicação - Design inspirado no Windows 11"""
    
    # Cores do tema Windows 11 Dark Mode com transparência reduzida
    BG_PRIMARY = "#202020"           # Fundo principal escuro
    BG_SECONDARY = "#2D2D2D"         # Fundo secundário
    BG_CARD = "#383838"              # Cartões e containers
    BG_HOVER = "#404040"             # Hover states
    BG_GLASS = "#383838"             # Efeito glass (sem transparência)
    
    # Cores de destaque (Windows 11 palette)
    ACCENT_PRIMARY = "#0078D4"       # Azul Windows
    ACCENT_SECONDARY = "#9A0089"     # Roxo Windows
    ACCENT_SUCCESS = "#107C10"       # Verde Windows
    ACCENT_WARNING = "#D83B01"       # Laranja Windows
    ACCENT_ERROR = "#D13438"         # Vermelho Windows
    
    # Texto - CORRIGIDO: Textos claros para fundo escuro
    TEXT_PRIMARY = "#FFFFFF"         # Texto primário (BRANCO)
    TEXT_SECONDARY = "#E0E0E0"       # Texto secundário (CINZA CLARO)
    TEXT_MUTED = "#A0A0A0"           # Texto muted (CINZA MÉDIO)
    TEXT_ACCENT = "#0078D4"          # Texto de destaque
    
    # Bordas e separadores
    BORDER_COLOR = "#484848"
    SEPARATOR_COLOR = "#484848"
    
    # Efeitos
    GLASS_ALPHA = 0.95               # Transparência reduzida (95% opaco)
    BORDER_RADIUS = 8                # Bordas mais suaves como Windows 11
    SHADOW_COLOR = "#00000030"
    
    # Fontes (Windows 11 Fonts)
    FONT_PRIMARY = ("Segoe UI", 10)
    FONT_SECONDARY = ("Segoe UI", 9)
    FONT_TITLE = ("Segoe UI", 16, "bold")
    FONT_HEADER = ("Segoe UI", 12, "bold")
    FONT_ACCENT = ("Segoe UI", 10, "bold")
    
    @classmethod
    def configure_styles(cls):
        """Configura todos os estilos ttk com tema Liquid Glass"""
        style = ttk.Style()
        
        # Tenta usar o tema 'clam' como base para estilos ttk modernos
        try:
            style.theme_use('clam')
        except:
            try:
                style.theme_use('alt')
            except:
                pass

        # Configurações gerais - CORRIGIDO: textos claros e sem transparência
        style.configure(".", 
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,  # Texto BRANCO
                       fieldbackground=cls.BG_SECONDARY,
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,  # Texto BRANCO
                       insertcolor=cls.TEXT_PRIMARY,       # Cursor BRANCO
                       troughcolor=cls.BG_SECONDARY,
                       focuscolor=cls.ACCENT_PRIMARY + "40")
        
        # Frame com efeito glass (sem transparência)
        style.configure("Glass.TFrame",
                       background=cls.BG_CARD,
                       relief="flat",
                       borderwidth=0)
        
        # Labels - CORRIGIDO: textos claros
        style.configure("Glass.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       font=cls.FONT_PRIMARY)
        
        style.configure("Title.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       font=cls.FONT_TITLE)
        
        style.configure("Subtitle.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_SECONDARY,  # CINZA CLARO
                       font=cls.FONT_SECONDARY)
        
        style.configure("Header.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       font=cls.FONT_HEADER)
        
        style.configure("Accent.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.ACCENT_PRIMARY,
                       font=cls.FONT_ACCENT)
        
        style.configure("Success.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.ACCENT_SUCCESS,
                       font=cls.FONT_ACCENT)
        
        style.configure("Warning.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.ACCENT_WARNING,
                       font=cls.FONT_ACCENT)
        
        style.configure("Error.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.ACCENT_ERROR,
                       font=cls.FONT_ACCENT)

        # Entry - CORRIGIDO: texto claro e fundo sólido
        style.configure("Glass.TEntry",
                       fieldbackground=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       bordercolor=cls.BORDER_COLOR,
                       lightcolor=cls.BORDER_COLOR,
                       darkcolor=cls.BORDER_COLOR,
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,  # BRANCO
                       insertcolor=cls.TEXT_PRIMARY,       # BRANCO
                       borderwidth=1,
                       relief="flat",
                       padding=(8, 6))
        
        style.map("Glass.TEntry",
                 bordercolor=[("focus", cls.ACCENT_PRIMARY),
                            ("hover", cls.ACCENT_PRIMARY + "80")],
                 lightcolor=[("focus", cls.ACCENT_PRIMARY)],
                 darkcolor=[("focus", cls.ACCENT_PRIMARY)])

        # Buttons - CORRIGIDO: textos claros e fundos sólidos
        style.configure("Accent.TButton",
                       background=cls.ACCENT_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       borderwidth=0,
                       focuscolor=cls.ACCENT_PRIMARY + "40",
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Accent.TButton",
                 background=[("active", "#106EBE"),
                           ("pressed", "#005A9E")],
                 relief=[("pressed", "sunken")])
        
        style.configure("Glass.TButton",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       borderwidth=1,
                       bordercolor=cls.BORDER_COLOR,
                       font=cls.FONT_PRIMARY,
                       relief="flat",
                       padding=(15, 8))
        
        style.map("Glass.TButton",
                 background=[("active", cls.BG_HOVER),
                           ("pressed", cls.ACCENT_PRIMARY + "20")],
                 bordercolor=[("active", cls.ACCENT_PRIMARY + "80")])
        
        style.configure("Success.TButton",
                       background=cls.ACCENT_SUCCESS,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       borderwidth=0,
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Success.TButton",
                 background=[("active", "#0E6E0E"),
                           ("pressed", "#0A5A0A")])
        
        style.configure("Warning.TButton",
                       background=cls.ACCENT_WARNING,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       borderwidth=0,
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Warning.TButton",
                 background=[("active", "#B83201"),
                           ("pressed", "#9A2A01")])
        
        style.configure("Error.TButton",
                       background=cls.ACCENT_ERROR,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       borderwidth=0,
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Error.TButton",
                 background=[("active", "#C12A2E"),
                           ("pressed", "#A32024")])

        style.map("Back.TButton",
                 background=[("active", "#e7b13c"),
                           ("pressed", "#A32024")])                   

        # Checkbutton/Radiobutton - CORRIGIDO: textos claros
        style.configure("Glass.TRadiobutton",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       indicatorcolor=cls.BG_SECONDARY,
                       indicatorrelief="raised",
                       indicatordiameter=12,
                       font=cls.FONT_SECONDARY)
        
        style.configure("Glass.TCheckbutton",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       indicatorcolor=cls.BG_SECONDARY,
                       indicatorrelief="raised",
                       indicatordiameter=12,
                       font=cls.FONT_SECONDARY)
        
        style.map("Glass.TRadiobutton",
                 indicatorcolor=[("selected", cls.ACCENT_PRIMARY),
                               ("active", cls.BG_HOVER)],
                 background=[("active", cls.BG_HOVER)])
        
        style.map("Glass.TCheckbutton",
                 indicatorcolor=[("selected", cls.ACCENT_PRIMARY),
                               ("active", cls.BG_HOVER)],
                 background=[("active", cls.BG_HOVER)])

        # Notebook - CORRIGIDO: textos claros
        style.configure("Glass.TNotebook",
                       background=cls.BG_PRIMARY,
                       borderwidth=0,
                       tabmargins=(2, 5, 2, 0))
        
        style.configure("Glass.TNotebook.Tab",
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_SECONDARY,  # CINZA CLARO
                       padding=(20, 8),
                       borderwidth=0,
                       font=cls.FONT_PRIMARY,
                       focuscolor=cls.BG_PRIMARY)
        
        style.map("Glass.TNotebook.Tab",
                 background=[("selected", cls.BG_CARD),
                           ("active", cls.BG_HOVER)],
                 foreground=[("selected", cls.ACCENT_PRIMARY),
                           ("active", cls.TEXT_PRIMARY)])  # BRANCO no hover

        # Separator
        style.configure("Glass.TSeparator",
                       background=cls.SEPARATOR_COLOR)

        # LabelFrame - CORRIGIDO: textos claros
        style.configure("Glass.TLabelframe",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       font=cls.FONT_HEADER,
                       borderwidth=1,
                       relief="solid",
                       labelmargins=(10, 5, 10, 5))
        
        style.configure("Glass.TLabelframe.Label",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       font=cls.FONT_HEADER)

        # Scrollbar
        style.configure("Glass.Vertical.TScrollbar",
                       background=cls.BG_SECONDARY,
                       darkcolor=cls.BG_SECONDARY,
                       lightcolor=cls.BG_SECONDARY,
                       troughcolor=cls.BG_PRIMARY,
                       bordercolor=cls.BG_PRIMARY,
                       arrowcolor=cls.TEXT_SECONDARY,  # CINZA CLARO
                       gripcount=0)
        
        style.configure("Glass.Horizontal.TScrollbar",
                       background=cls.BG_SECONDARY,
                       darkcolor=cls.BG_SECONDARY,
                       lightcolor=cls.BG_SECONDARY,
                       troughcolor=cls.BG_PRIMARY,
                       bordercolor=cls.BG_PRIMARY,
                       arrowcolor=cls.TEXT_SECONDARY,  # CINZA CLARO
                       gripcount=0)
        
        style.map("Glass.Vertical.TScrollbar",
                 background=[("active", cls.BG_HOVER),
                           ("pressed", cls.ACCENT_PRIMARY)])
        
        style.map("Glass.Horizontal.TScrollbar",
                 background=[("active", cls.BG_HOVER),
                           ("pressed", cls.ACCENT_PRIMARY)])

        # Progressbar
        style.configure("Glass.Horizontal.TProgressbar",
                       background=cls.ACCENT_PRIMARY,
                       troughcolor=cls.BG_SECONDARY,
                       bordercolor=cls.BG_SECONDARY,
                       lightcolor=cls.ACCENT_PRIMARY,
                       darkcolor=cls.ACCENT_PRIMARY,
                       borderwidth=0,
                       thickness=8)
        
        style.configure("Glass.Vertical.TProgressbar",
                       background=cls.ACCENT_PRIMARY,
                       troughcolor=cls.BG_SECONDARY,
                       bordercolor=cls.BG_SECONDARY,
                       lightcolor=cls.ACCENT_PRIMARY,
                       darkcolor=cls.ACCENT_PRIMARY,
                       borderwidth=0,
                       thickness=8)

        # Treeview - CORRIGIDO: textos claros
        style.configure("Glass.Treeview",
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       fieldbackground=cls.BG_SECONDARY,
                       borderwidth=0,
                       relief="flat",
                       rowheight=28)
        
        style.configure("Glass.Treeview.Heading",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       relief="flat",
                       borderwidth=0,
                       font=cls.FONT_ACCENT)
        
        style.map("Glass.Treeview",
                 background=[("selected", cls.ACCENT_PRIMARY)],
                 foreground=[("selected", cls.TEXT_PRIMARY)])  # BRANCO no selecionado
        
        style.map("Glass.Treeview.Heading",
                 background=[("active", cls.BG_HOVER)])

        # Combobox - CORRIGIDO: textos claros
        style.configure("Glass.TCombobox",
                       fieldbackground=cls.BG_SECONDARY,
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,  # BRANCO
                       insertcolor=cls.TEXT_PRIMARY,       # BRANCO
                       borderwidth=1,
                       relief="flat",
                       padding=(8, 6))
        
        style.map("Glass.TCombobox",
                 fieldbackground=[("readonly", cls.BG_SECONDARY)],
                 background=[("readonly", cls.BG_SECONDARY)],
                 bordercolor=[("focus", cls.ACCENT_PRIMARY),
                            ("hover", cls.ACCENT_PRIMARY + "80")],
                 lightcolor=[("focus", cls.ACCENT_PRIMARY)],
                 darkcolor=[("focus", cls.ACCENT_PRIMARY)])

        # Scale
        style.configure("Glass.Horizontal.TScale",
                       background=cls.BG_PRIMARY,
                       troughcolor=cls.BG_SECONDARY,
                       bordercolor=cls.BORDER_COLOR,
                       sliderrelief="flat",
                       borderwidth=0)
        
        style.configure("Glass.Vertical.TScale",
                       background=cls.BG_PRIMARY,
                       troughcolor=cls.BG_SECONDARY,
                       bordercolor=cls.BORDER_COLOR,
                       sliderrelief="flat",
                       borderwidth=0)
        
        style.map("Glass.Horizontal.TScale",
                 background=[("active", cls.ACCENT_PRIMARY)],
                 troughcolor=[("active", cls.BG_HOVER)])
        
        style.map("Glass.Vertical.TScale",
                 background=[("active", cls.ACCENT_PRIMARY)],
                 troughcolor=[("active", cls.BG_HOVER)])

        # Spinbox - CORRIGIDO: textos claros
        style.configure("Glass.TSpinbox",
                       fieldbackground=cls.BG_SECONDARY,
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,  # BRANCO
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,  # BRANCO
                       insertcolor=cls.TEXT_PRIMARY,       # BRANCO
                       borderwidth=1,
                       relief="flat",
                       padding=(8, 6))
        
        style.map("Glass.TSpinbox",
                 bordercolor=[("focus", cls.ACCENT_PRIMARY),
                            ("hover", cls.ACCENT_PRIMARY + "80")],
                 lightcolor=[("focus", cls.ACCENT_PRIMARY)],
                 darkcolor=[("focus", cls.ACCENT_PRIMARY)])

        # Sizegrip
        style.configure("Glass.TSizegrip",
                       background=cls.BG_SECONDARY,
                       relief="flat")

    @classmethod
    def create_glass_frame(cls, parent, **kwargs):
        """Cria um frame com efeito glass"""
        frame = tk.Frame(parent, 
                        bg=cls.BG_CARD,
                        relief="flat",
                        bd=0,
                        **kwargs)
        return frame

    @classmethod
    def create_accent_button(cls, parent, text, command, **kwargs):
        """Cria um botão de destaque"""
        btn = ttk.Button(parent, 
                        text=text, 
                        command=command,
                        style="Accent.TButton",
                        **kwargs)
        return btn

    @classmethod
    def create_glass_button(cls, parent, text, command, **kwargs):
        """Cria um botão glass"""
        btn = ttk.Button(parent, 
                        text=text, 
                        command=command,
                        style="Glass.TButton",
                        **kwargs)
        return btn

    @classmethod
    def create_title_label(cls, parent, text, **kwargs):
        """Cria um label de título"""
        label = ttk.Label(parent,
                         text=text,
                         style="Title.TLabel",
                         **kwargs)
        return label

    @classmethod
    def create_glass_entry(cls, parent, **kwargs):
        """Cria um entry com estilo glass"""
        entry = ttk.Entry(parent,
                         style="Glass.TEntry",
                         **kwargs)
        return entry

    @classmethod
    def apply_window_style(cls, window):
        """Aplica o estilo Liquid Glass a uma janela"""
        window.configure(bg=cls.BG_PRIMARY)
        
        # Configura transparência reduzida (95% opaco)
        try:
            window.wm_attributes('-alpha', cls.GLASS_ALPHA)
            # Remove transparentcolor para evitar elementos totalmente transparentes
            if window.tk.call('tk', 'windowingsystem') == 'win32':
                # Usa um alpha mais alto para evitar transparência total
                window.wm_attributes('-alpha', 0.98)
        except:
            # Fallback se não suportar transparência
            pass

    @classmethod
    def create_card(cls, parent, **kwargs):
        """Cria um card com efeito glass"""
        card = tk.Frame(parent,
                       bg=cls.BG_CARD,
                       relief="flat",
                       bd=0,
                       **kwargs)
        return card

    @classmethod
    def create_progressbar(cls, parent, **kwargs):
        """Cria uma barra de progresso com estilo glass"""
        progress = ttk.Progressbar(parent,
                                  style="Glass.Horizontal.TProgressbar",
                                  **kwargs)
        return progress

    @classmethod
    def create_scrollbar(cls, parent, **kwargs):
        """Cria uma scrollbar com estilo glass"""
        scrollbar = ttk.Scrollbar(parent,
                                 style="Glass.Vertical.TScrollbar",
                                 **kwargs)
        return scrollbar


# Configurar estilos automaticamente ao importar
LiquidGlassStyle.configure_styles()