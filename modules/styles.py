# styles.py

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font as tkfont

class LiquidGlassStyle:
    """Estilo Liquid Glass para a aplicação - Design inspirado na Apple"""
    
    # Cores do tema Liquid Glass (Apple-inspired)
    BG_PRIMARY = "#0a0e14"        # Fundo principal escuro
    BG_SECONDARY = "#1a1f2e"      # Fundo secundário
    BG_CARD = "#252a3a"           # Cartões e containers
    BG_HOVER = "#2d3448"          # Hover states
    BG_GLASS = "rgba(37, 42, 58, 0.8)"  # Efeito glass
    
    # Cores de destaque (Apple palette)
    ACCENT_PRIMARY = "#007AFF"    # Azul Apple
    ACCENT_SECONDARY = "#5856D6"  # Roxo Apple
    ACCENT_SUCCESS = "#34C759"    # Verde Apple
    ACCENT_WARNING = "#FF9500"    # Laranja Apple
    ACCENT_ERROR = "#FF3B30"      # Vermelho Apple
    
    # Texto
    TEXT_PRIMARY = "#FFFFFF"      # Texto primário
    TEXT_SECONDARY = "#8E8E93"    # Texto secundário
    TEXT_MUTED = "#48484A"        # Texto muted
    TEXT_ACCENT = "#007AFF"       # Texto de destaque
    
    # Bordas e separadores
    BORDER_COLOR = "#38383A"
    SEPARATOR_COLOR = "#38383A"
    
    # Efeitos
    GLASS_ALPHA = 0.15
    BORDER_RADIUS = 12
    SHADOW_COLOR = "#00000020"
    
    # Fontes (Apple System Fonts)
    FONT_PRIMARY = ("SF Pro Display", 10)
    FONT_SECONDARY = ("SF Pro Text", 9)
    FONT_TITLE = ("SF Pro Display", 16, "bold")
    FONT_HEADER = ("SF Pro Display", 12, "bold")
    FONT_ACCENT = ("SF Pro Text", 10, "bold")
    
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

        # Configurações gerais
        style.configure(".", 
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       fieldbackground=cls.BG_SECONDARY,
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,
                       insertcolor=cls.TEXT_PRIMARY,
                       troughcolor=cls.BG_SECONDARY,
                       focuscolor=cls.ACCENT_PRIMARY + "20")
        
        # Frame com efeito glass
        style.configure("Glass.TFrame",
                       background=cls.BG_CARD,
                       relief="flat",
                       borderwidth=0)
        
        # Labels
        style.configure("Glass.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,
                       font=cls.FONT_PRIMARY)
        
        style.configure("Title.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       font=cls.FONT_TITLE)
        
        style.configure("Subtitle.TLabel",
                       background=cls.BG_PRIMARY,
                       foreground=cls.TEXT_SECONDARY,
                       font=cls.FONT_SECONDARY)
        
        style.configure("Header.TLabel",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,
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

        # Entry
        style.configure("Glass.TEntry",
                       fieldbackground=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,
                       bordercolor=cls.BORDER_COLOR,
                       lightcolor=cls.BORDER_COLOR,
                       darkcolor=cls.BORDER_COLOR,
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,
                       insertcolor=cls.TEXT_PRIMARY,
                       borderwidth=1,
                       relief="flat",
                       padding=(8, 6))
        
        style.map("Glass.TEntry",
                 bordercolor=[("focus", cls.ACCENT_PRIMARY),
                            ("hover", cls.BG_HOVER)],
                 lightcolor=[("focus", cls.ACCENT_PRIMARY)],
                 darkcolor=[("focus", cls.ACCENT_PRIMARY)])

        # Buttons
        style.configure("Accent.TButton",
                       background=cls.ACCENT_PRIMARY,
                       foreground=cls.TEXT_PRIMARY,
                       borderwidth=0,
                       focuscolor=cls.ACCENT_PRIMARY + "20",
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Accent.TButton",
                 background=[("active", cls.ACCENT_SECONDARY),
                           ("pressed", cls.ACCENT_SECONDARY)],
                 relief=[("pressed", "flat")])
        
        style.configure("Glass.TButton",
                       background=cls.BG_HOVER,
                       foreground=cls.TEXT_PRIMARY,
                       borderwidth=0,
                       font=cls.FONT_PRIMARY,
                       relief="flat",
                       padding=(15, 8))
        
        style.map("Glass.TButton",
                 background=[("active", cls.BG_SECONDARY),
                           ("pressed", cls.ACCENT_PRIMARY + "20")])
        
        style.configure("Success.TButton",
                       background=cls.ACCENT_SUCCESS,
                       foreground=cls.TEXT_PRIMARY,
                       borderwidth=0,
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Success.TButton",
                 background=[("active", "#30a850"),
                           ("pressed", "#30a850")])
        
        style.configure("Warning.TButton",
                       background=cls.ACCENT_WARNING,
                       foreground=cls.TEXT_PRIMARY,
                       borderwidth=0,
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Warning.TButton",
                 background=[("active", "#e68a00"),
                           ("pressed", "#e68a00")])
        
        style.configure("Error.TButton",
                       background=cls.ACCENT_ERROR,
                       foreground=cls.TEXT_PRIMARY,
                       borderwidth=0,
                       font=cls.FONT_ACCENT,
                       relief="flat",
                       padding=(20, 10))
        
        style.map("Error.TButton",
                 background=[("active", "#e63530"),
                           ("pressed", "#e63530")])

        # Checkbutton/Radiobutton
        style.configure("Glass.TRadiobutton",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,
                       indicatorcolor=cls.BG_SECONDARY,
                       indicatorrelief="raised",
                       indicatordiameter=12,
                       font=cls.FONT_SECONDARY)
        
        style.configure("Glass.TCheckbutton",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,
                       indicatorcolor=cls.BG_SECONDARY,
                       indicatorrelief="raised",
                       indicatordiameter=12,
                       font=cls.FONT_SECONDARY)
        
        style.map("Glass.TRadiobutton",
                 indicatorcolor=[("selected", cls.ACCENT_PRIMARY),
                               ("active", cls.BG_HOVER)])
        
        style.map("Glass.TCheckbutton",
                 indicatorcolor=[("selected", cls.ACCENT_PRIMARY),
                               ("active", cls.BG_HOVER)])

        # Notebook
        style.configure("Glass.TNotebook",
                       background=cls.BG_PRIMARY,
                       borderwidth=0,
                       tabmargins=(2, 5, 2, 0))
        
        style.configure("Glass.TNotebook.Tab",
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_SECONDARY,
                       padding=(20, 10),
                       borderwidth=0,
                       font=cls.FONT_PRIMARY,
                       focuscolor=cls.BG_PRIMARY)
        
        style.map("Glass.TNotebook.Tab",
                 background=[("selected", cls.BG_CARD),
                           ("active", cls.BG_HOVER)],
                 foreground=[("selected", cls.ACCENT_PRIMARY),
                           ("active", cls.TEXT_PRIMARY)])

        # Separator
        style.configure("Glass.TSeparator",
                       background=cls.SEPARATOR_COLOR)

        # LabelFrame
        style.configure("Glass.TLabelframe",
                       background=cls.BG_CARD,
                       foreground=cls.ACCENT_PRIMARY,
                       font=cls.FONT_HEADER,
                       borderwidth=1,
                       relief="solid",
                       labelmargins=(10, 5, 10, 5))
        
        style.configure("Glass.TLabelframe.Label",
                       background=cls.BG_CARD,
                       foreground=cls.ACCENT_PRIMARY,
                       font=cls.FONT_HEADER)

        # Scrollbar
        style.configure("Glass.Vertical.TScrollbar",
                       background=cls.BG_SECONDARY,
                       darkcolor=cls.BG_SECONDARY,
                       lightcolor=cls.BG_SECONDARY,
                       troughcolor=cls.BG_PRIMARY,
                       bordercolor=cls.BG_PRIMARY,
                       arrowcolor=cls.TEXT_SECONDARY,
                       gripcount=0)
        
        style.configure("Glass.Horizontal.TScrollbar",
                       background=cls.BG_SECONDARY,
                       darkcolor=cls.BG_SECONDARY,
                       lightcolor=cls.BG_SECONDARY,
                       troughcolor=cls.BG_PRIMARY,
                       bordercolor=cls.BG_PRIMARY,
                       arrowcolor=cls.TEXT_SECONDARY,
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

        # Treeview
        style.configure("Glass.Treeview",
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,
                       fieldbackground=cls.BG_SECONDARY,
                       borderwidth=0,
                       relief="flat",
                       rowheight=25)
        
        style.configure("Glass.Treeview.Heading",
                       background=cls.BG_CARD,
                       foreground=cls.TEXT_PRIMARY,
                       relief="flat",
                       borderwidth=0,
                       font=cls.FONT_ACCENT)
        
        style.map("Glass.Treeview",
                 background=[("selected", cls.ACCENT_PRIMARY)],
                 foreground=[("selected", cls.TEXT_PRIMARY)])
        
        style.map("Glass.Treeview.Heading",
                 background=[("active", cls.BG_HOVER)])

        # Combobox
        style.configure("Glass.TCombobox",
                       fieldbackground=cls.BG_SECONDARY,
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,
                       insertcolor=cls.TEXT_PRIMARY,
                       borderwidth=1,
                       relief="flat",
                       padding=(8, 6))
        
        style.map("Glass.TCombobox",
                 fieldbackground=[("readonly", cls.BG_SECONDARY)],
                 background=[("readonly", cls.BG_SECONDARY)],
                 bordercolor=[("focus", cls.ACCENT_PRIMARY),
                            ("hover", cls.BG_HOVER)],
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

        # Spinbox
        style.configure("Glass.TSpinbox",
                       fieldbackground=cls.BG_SECONDARY,
                       background=cls.BG_SECONDARY,
                       foreground=cls.TEXT_PRIMARY,
                       selectbackground=cls.ACCENT_PRIMARY,
                       selectforeground=cls.TEXT_PRIMARY,
                       insertcolor=cls.TEXT_PRIMARY,
                       borderwidth=1,
                       relief="flat",
                       padding=(8, 6))
        
        style.map("Glass.TSpinbox",
                 bordercolor=[("focus", cls.ACCENT_PRIMARY),
                            ("hover", cls.BG_HOVER)],
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
        
        # Tenta configurar a transparência (efeito glass)
        try:
            if window.tk.call('tk', 'windowingsystem') == 'win32':
                window.wm_attributes('-transparentcolor', cls.BG_PRIMARY)
            elif window.tk.call('tk', 'windowingsystem') == 'aqua':
                window.wm_attributes('-transparent', True)
        except:
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