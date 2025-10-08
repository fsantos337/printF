import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, colorchooser, ttk
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import pyautogui
from pynput import mouse, keyboard
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageGrab 
from datetime import datetime
import math
import re
import screeninfo
import glob
import json
import uuid
import time
import ctypes
from ctypes import wintypes, byref
import tkinter.font as tkfont

# 🔥 CONTROLE AUTOMÁTICO DA BARRA DE TAREFAS
try:
    import win32gui
    import win32con
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    win32gui = None
    win32con = None
    win32api = None
    WIN32_AVAILABLE = False
    
# 🔥 ADICIONAR MSS PARA CAPTURA MULTI-MONITOR
try:
    import mss
    MSS_AVAILABLE = True
except ImportError:
    mss = None
    MSS_AVAILABLE = False

# ------------------ Gravador e Docx ------------------
class CaptureModule:
    def __init__(self, parent=None, settings=None):
        """Inicializa o módulo de captura para trabalhar com a main.py"""
        self.parent = parent
        self.settings = settings or {}
        self.root = None  # Será definido quando mostrar a interface
        
        # 🔥 ADICIONADO: Gerenciador de estilo
        self.style_manager = None
        self.using_liquid_glass = False
        self._setup_styles()
        
        self.gravando = False
        self.pausado = False
        self.output_dir = os.getcwd()
        self.listener_mouse = None
        self.prints = []
        self.doc = None
        self.evidencia_count = 0
        self.demanda = ""
        self.tipo_demanda = ""
        self.chamado = ""
        self.titulo = ""
        self.using_template = False
        self.template_path = None
        self.color_chooser_window = None
        self.current_index = 0
        self.evidence_dir = None
        self.metadata_path = None
        self.metadata = {"evidencias": [], "proximo_id": 1}
        self.popup = None
        self.current_img_label = None
        self.current_img_tk = None
        self.comment_entry = None
        self.manter_evidencias = None
        self.modo_captura = "ocultar"  # Valores: "manter", "ocultar"
        
        # Listener de teclado para atalhos
        self.listener_keyboard = None

    def _setup_styles(self):
        """Configura estilos visuais baseados no tema selecionado"""
        try:
            # 🔥 CORREÇÃO: Tentar importar de múltiplas formas
            try:
                # Tentar importar do módulo styles diretamente
                from styles import LiquidGlassStyle
                self.style_manager = LiquidGlassStyle
                self.using_liquid_glass = True
                print("✅ Estilo Liquid Glass carregado do módulo styles!")
                
            except ImportError:
                try:
                    # Tentar importar do diretório modules
                    from modules.styles import LiquidGlassStyle
                    self.style_manager = LiquidGlassStyle
                    self.using_liquid_glass = True
                    print("✅ Estilo Liquid Glass carregado do módulo modules.styles!")
                    
                except ImportError:
                    # Tentar importar relativo
                    import importlib.util
                    spec = importlib.util.spec_from_file_location("styles", "styles.py")
                    if spec and spec.loader:
                        styles_module = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(styles_module)
                        self.style_manager = styles_module.LiquidGlassStyle
                        self.using_liquid_glass = True
                        print("✅ Estilo Liquid Glass carregado de styles.py!")
                    else:
                        raise ImportError("Não foi possível carregar styles.py")
            
            # Verificar se o tema está habilitado nas configurações
            theme_to_use = self.settings.get('theme', 'liquid_glass')
            if theme_to_use == 'liquid_glass' and self.style_manager:
                self.using_liquid_glass = True
                print("✅ Estilo Liquid Glass configurado no módulo de captura!")
            else:
                self.using_liquid_glass = False
                print(f"ℹ️ Usando estilo padrão no módulo de captura (tema: {theme_to_use})")
            
        except ImportError as e:
            # Fallback para estilo padrão
            print(f"⚠️ Liquid Glass não disponível no módulo de captura: {e}")
            self.using_liquid_glass = False
        except Exception as e:
            print(f"⚠️ Erro ao configurar Liquid Glass no módulo de captura: {e}")
            self.using_liquid_glass = False

    def _apply_style_to_window(self, window):
        """Aplica o estilo Liquid Glass a uma janela se disponível"""
        if self.using_liquid_glass and self.style_manager:
            try:
                self.style_manager.apply_window_style(window)
                return True
            except Exception as e:
                print(f"⚠️ Erro ao aplicar estilo à janela: {e}")
                self.using_liquid_glass = False
        return False

    def _create_styled_frame(self, parent, **kwargs):
        """Cria um frame com estilo apropriado"""
        if self.using_liquid_glass and self.style_manager:
            try:
                return self.style_manager.create_glass_frame(parent, **kwargs)
            except Exception as e:
                print(f"⚠️ Erro ao criar frame estilizado: {e}")
                self.using_liquid_glass = False
        
        # Fallback para frame padrão
        frame = tk.Frame(parent, **kwargs)
        if not self.using_liquid_glass:
            frame.configure(bg='#f5f5f5')
        return frame

    def _create_styled_button(self, parent, text, command, style_type="glass", **kwargs):
        """Cria um botão com estilo apropriado"""
        if self.using_liquid_glass and self.style_manager:
            try:
                if style_type == "accent":
                    return self.style_manager.create_accent_button(parent, text, command, **kwargs)
                else:
                    return self.style_manager.create_glass_button(parent, text, command, **kwargs)
            except Exception as e:
                print(f"⚠️ Erro ao criar botão estilizado: {e}")
                self.using_liquid_glass = False
        
        # Fallback para botão padrão
        btn = tk.Button(parent, text=text, command=command, **kwargs)
        if style_type == "accent":
            btn.configure(bg="#3498db", fg="white", font=("Arial", 11, "bold"), relief="flat")
        else:
            btn.configure(bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10), relief="flat")
        return btn

    def _create_styled_label(self, parent, text, style_type="glass", **kwargs):
        """Cria um label com estilo apropriado"""
        if self.using_liquid_glass and self.style_manager:
            try:
                if style_type == "title":
                    return self.style_manager.create_title_label(parent, text, **kwargs)
                else:
                    return ttk.Label(parent, text=text, style="Glass.TLabel", **kwargs)
            except Exception as e:
                print(f"⚠️ Erro ao criar label estilizado: {e}")
                self.using_liquid_glass = False
        
        # Fallback para label padrão
        label = tk.Label(parent, text=text, **kwargs)
        if style_type == "title":
            label.configure(font=("Arial", 16, "bold"), fg="#2c3e50", bg='#f5f5f5')
        else:
            label.configure(font=("Arial", 10), fg="#2c3e50", bg='#f5f5f5')
        return label

    def _create_styled_entry(self, parent, **kwargs):
        """Cria um entry com estilo apropriado"""
        if self.using_liquid_glass and self.style_manager:
            try:
                return self.style_manager.create_glass_entry(parent, **kwargs)
            except Exception as e:
                print(f"⚠️ Erro ao criar entry estilizado: {e}")
                self.using_liquid_glass = False
        
        # Fallback para entry padrão
        return tk.Entry(parent, **kwargs)

    def show(self):
        """Mostra a interface do módulo de captura"""
        if not self.root:
            self._create_interface()
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_set()

    def hide(self):
        """Esconde a interface do módulo de forma segura"""
        if self.root:
            try:
                # Parar qualquer gravação em andamento
                if self.gravando:
                    self.finalizar()
                
                # Parar listeners
                if hasattr(self, 'listener_keyboard') and self.listener_keyboard:
                    try:
                        self.listener_keyboard.stop()
                    except:
                        pass
                    self.listener_keyboard = None
                
                if self.listener_mouse:
                    try:
                        self.listener_mouse.stop()
                    except:
                        pass
                    self.listener_mouse = None
                
                # Fechar janelas secundárias
                if self.popup and self.popup.winfo_exists():
                    try:
                        self.popup.destroy()
                    except:
                        pass
                    self.popup = None
                
                # Liberar grabs
                try:
                    self.root.grab_release()
                except:
                    pass
                    
                # Esconder a janela
                self.root.withdraw()
                
                # 🔥 CORREÇÃO EXTRA: Voltar o foco para a janela principal
                if self.parent and self.parent.winfo_exists():
                    try:
                        self.parent.deiconify()
                        self.parent.lift()
                        self.parent.focus_force()
                    except:
                        pass
                    
            except Exception as e:
                print(f"Erro ao esconder módulo: {e}")
                # Fallback: destruir completamente se houver problemas
                try:
                    self.root.destroy()
                    self.root = None
                except:
                    pass

    def _create_interface(self):
        """Cria a interface do módulo de captura"""
        self.root = tk.Toplevel(self.parent)
        self.root.title("PrintF - Capturar Evidências")
        self.root.geometry("500x400")
        
        # 🔥 APLICAR ESTILO À JANELA
        self._apply_style_to_window(self.root)
        
        # 🔥 CORREÇÃO: Usar protocolo correto para fechar
        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.resizable(False, False)

        # 🔥 CRIAR FRAME PRINCIPAL COM ESTILO
        main_frame = self._create_styled_frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Interface com estilos aplicados
        self._create_styled_label(main_frame, text="📷 PrintF - Capturar Evidências", 
                                style_type="title").pack(pady=20)
        
        self._create_styled_button(main_frame, text="▶ Iniciar Gravação (F8)", 
                                 command=self.iniciar, style_type="accent").pack(pady=8, fill=tk.X)
        
        self._create_styled_button(main_frame, text="⏸ Pausar Gravação (F6)", 
                                 command=self.pausar, style_type="glass").pack(pady=8, fill=tk.X)
        
        self._create_styled_button(main_frame, text="▶ Retomar Gravação (F7)", 
                                 command=self.retomar, style_type="glass").pack(pady=8, fill=tk.X)
        
        self._create_styled_button(main_frame, text="⏹ Finalizar Gravação (F9)", 
                                 command=self.finalizar, style_type="glass").pack(pady=8, fill=tk.X)
        
        # 🔥 BOTÃO VOLTAR COM ESTILO DE ERRO (vermelho)
        if self.using_liquid_glass and self.style_manager:
            try:
                voltar_btn = ttk.Button(main_frame, text="⬅ Voltar ao Menu", 
                                      command=self.hide,
                                      style="Error.TButton")
                voltar_btn.pack(pady=15, fill=tk.X)
            except:
                # Fallback se o estilo Error não estiver disponível
                self._create_styled_button(main_frame, text="⬅ Voltar ao Menu", 
                                         command=self.hide, style_type="glass").pack(pady=15, fill=tk.X)
        else:
            voltar_btn = tk.Button(main_frame, text="⬅ Voltar ao Menu", 
                                 command=self.hide,
                                 bg="#e74c3c", fg="white", font=("Arial", 11), relief="flat")
            voltar_btn.pack(pady=15, fill=tk.X)

        # Configurar atalhos
        self._setup_shortcuts()

    def _setup_shortcuts(self):
        """Configura atalhos de teclado"""
        def on_press(key):
            try:
                if key == keyboard.Key.f6:
                    self.root.after(0, self.pausar)
                elif key == keyboard.Key.f7:
                    self.root.after(0, self.retomar)
                elif key == keyboard.Key.f8:
                    self.root.after(0, self.iniciar)
                elif key == keyboard.Key.f9:
                    self.root.after(0, self.finalizar)
            except Exception as e:
                print(f"Erro no atalho: {e}")

        # 🔥 CORREÇÃO: Parar listener anterior se existir
        if hasattr(self, 'listener_keyboard') and self.listener_keyboard:
            try:
                self.listener_keyboard.stop()
            except:
                pass
                
        self.listener_keyboard = keyboard.Listener(on_press=on_press, suppress=False)
        self.listener_keyboard.start()

    def _salvar_metadata(self):
        """Salva os metadados no arquivo JSON"""
        if self.metadata_path:
            with open(self.metadata_path, 'w', encoding='utf-8') as f:
                json.dump(self.metadata, f, indent=2, ensure_ascii=False)

    def carregar_evidencias(self, dir_path):
        """Carrega as evidências baseadas nos metadados"""
        self.metadata_path = os.path.join(dir_path, "evidencias_metadata.json")
        
        if os.path.exists(self.metadata_path):
            try:
                with open(self.metadata_path, 'r', encoding='utf-8') as f:
                    self.metadata = json.load(f)
            except:
                self.metadata = {"evidencias": [], "proximo_id": 1}
        
        # Carrega evidências ativas (não excluídas)
        evidencias_ativas = []
        for evidencia in self.metadata["evidencias"]:
            if not evidencia.get("excluida", False):
                caminho = os.path.join(dir_path, evidencia["arquivo"])
                if os.path.exists(caminho):
                    evidencias_ativas.append(caminho)
        
        return evidencias_ativas

    def recarregar_evidencias(self):
        """Recarrega a lista de evidências"""
        if self.evidence_dir:
            self.prints = self.carregar_evidencias(self.evidence_dir)
            return True
        return False

    # 🔥 MÉTODOS DE CAPTURA SIMPLIFICADOS E OTIMIZADOS
    def capture_inteligente(self, x, y):
        """
        Captura a tela baseado no modo selecionado pelo usuário
        """
        if self.modo_captura == "manter":
            # Modo manter: captura tela COMPLETA (incluindo barra de tarefas)
            return self.capture_tela_completa_mss(x, y)
        else:
            # Modo ocultar: captura apenas área de trabalho (sem barra)
            return self.capture_work_area_pyautogui(x, y)

    def capture_tela_completa_mss(self, x, y):
        """
        Captura a tela completa INCLUINDO a barra de tarefas.
        Funciona no primário e secundário, mesmo com coordenadas negativas.
        """
        try:
            # 🔥 ESTRATÉGIA 1: Win32 API para captura precisa de monitor específico
            if WIN32_AVAILABLE:
                try:
                    # Encontrar o monitor que contém o ponto (x, y)
                    monitor_handle = win32api.MonitorFromPoint((x, y), win32con.MONITOR_DEFAULTTONEAREST)
                    monitor_info = win32gui.GetMonitorInfo(monitor_handle)
                    
                    # Área completa do monitor (inclui barra)
                    monitor_area = monitor_info["Monitor"]  # (left, top, right, bottom)
                    
                    # Capturar usando MSS para melhor compatibilidade com múltiplos monitores
                    if MSS_AVAILABLE:
                        with mss.mss() as sct:
                            monitor_mss = {
                                "left": monitor_area[0],
                                "top": monitor_area[1], 
                                "width": monitor_area[2] - monitor_area[0],
                                "height": monitor_area[3] - monitor_area[1]
                            }
                            screenshot = sct.grab(monitor_mss)
                            img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                            
                            rel_x = x - monitor_area[0]
                            rel_y = y - monitor_area[1]
                            
                            metodo_utilizado = f"Win32 + MSS Monitor Completo {monitor_area}"
                            print(f"✅ CAPTURA WIN32+MSS - Monitor {monitor_area} | Coord: ({rel_x},{rel_y})")
                            
                            return img, (rel_x, rel_y), metodo_utilizado
                    else:
                        # Fallback para ImageGrab se MSS não disponível
                        screenshot = ImageGrab.grab(bbox=monitor_area)
                        rel_x = x - monitor_area[0]
                        rel_y = y - monitor_area[1]
                        
                        metodo_utilizado = f"Win32 Monitor Completo {monitor_area}"
                        print(f"✅ CAPTURA WIN32 - Monitor {monitor_area} | Coord: ({rel_x},{rel_y})")
                        
                        return screenshot, (rel_x, rel_y), metodo_utilizado
                        
                except Exception as e:
                    print(f"⚠️  Win32 falhou (capturando com alternativa): {e}")

            # 🔥 ESTRATÉGIA 2: MSS como alternativa principal
            if MSS_AVAILABLE:
                try:
                    with mss.mss() as sct:
                        monitor_encontrado = None
                        
                        # Procurar em todos os monitores (exceto o virtual)
                        for monitor in sct.monitors[1:]:
                            if (monitor["left"] <= x < monitor["left"] + monitor["width"] and
                                monitor["top"] <= y < monitor["top"] + monitor["height"]):
                                monitor_encontrado = monitor
                                break

                        # Fallback para primeiro monitor se não encontrou
                        if not monitor_encontrado:
                            monitor_encontrado = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                            print(f"⚠️  Monitor não encontrado para coordenadas ({x},{y}), usando monitor {monitor_encontrado} como fallback")

                        # Capturar a tela completa do monitor encontrado
                        screenshot = sct.grab(monitor_encontrado)
                        img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

                        # Calcular coordenadas relativas ao monitor
                        rel_x = x - monitor_encontrado["left"]
                        rel_y = y - monitor_encontrado["top"]

                        metodo_utilizado = f"MSS Monitor Completo {monitor_encontrado['width']}x{monitor_encontrado['height']}"
                        print(f"✅ CAPTURA MSS - Monitor {monitor_encontrado} | Coord: ({rel_x},{rel_y})")

                        return img, (rel_x, rel_y), metodo_utilizado
                        
                except Exception as e:
                    print(f"⚠️  MSS falhou (capturando com alternativa): {e}")

            # 🔥 ESTRATÉGIA 3: Fallback com ImageGrab
            try:
                screenshot = ImageGrab.grab()
                metodo_utilizado = "Fallback - ImageGrab (tela completa)"
                print(f"⚠️  Usando fallback ImageGrab para coordenadas ({x},{y})")
                return screenshot, (x, y), metodo_utilizado
                
            except Exception as e:
                print(f"⚠️  ImageGrab falhou: {e}")

            # 🔥 ESTRATÉGIA 4: Fallback final com pyautogui
            try:
                screenshot = pyautogui.screenshot()
                metodo_utilizado = "Fallback - pyautogui (apenas primário)"
                print(f"⚠️  Usando fallback pyautogui para coordenadas ({x},{y})")
                return screenshot, (x, y), metodo_utilizado
                
            except Exception as e:
                print(f"❌ Todos os métodos de captura falharam: {e}")
                raise

        except Exception as e:
            print(f"❌ Falha crítica na captura completa: {e}")
            # Último recurso - retorna imagem preta ou levanta exceção
            try:
                # Tenta criar uma imagem preta como fallback extremo
                img = Image.new('RGB', (100, 100), color='black')
                return img, (0, 0), f"Fallback Extremo - Erro: {str(e)}"
            except:
                raise Exception(f"Falha completa na captura de tela: {str(e)}")

    def capture_work_area_pyautogui(self, x, y):
        """
        Captura apenas a área de trabalho (SEM barra de tarefas) usando pyautogui
        Este método é usado no modo "ocultar"
        """
        try:
            # Captura com pyautogui (já captura sem a barra automaticamente)
            screenshot = pyautogui.screenshot()
            
            metodo_utilizado = "PyAutoGUI - Área de Trabalho (sem barra)"
            print(f"✅ CAPTURA PYAUTOGUI - Área de Trabalho | Coord: ({x},{y})")
            
            return screenshot, (x, y), metodo_utilizado
            
        except Exception as e:
            print(f"❌ Falha na captura com pyautogui: {e}")
            # Fallback extremo
            screenshot = pyautogui.screenshot()
            return screenshot, (x, y), f"Fallback - Erro: {str(e)}"

    def estimativa_segura_barra_tarefas(self, altura_tela):
        """
        Estimativa conservadora da altura da barra de tarefas
        """
        if altura_tela >= 2160:  # 4K
            return 100
        elif altura_tela >= 1440:  # QHD
            return 80
        elif altura_tela >= 1080:  # Full HD
            return 70
        else:  # Resoluções menores
            return 60
        
    # 🔥 NOVA FUNÇÃO: APLICAR TIMESTAMP MODERNO COM FUNDO
    def aplicar_timestamp_moderno(self, caminho_imagem, evidencia_meta):
        """Aplica o timestamp com fundo semi-transparente e texto branco"""
        img = Image.open(caminho_imagem).convert("RGBA")
        draw = ImageDraw.Draw(img)
        
        # Calcular posição em pixels
        img_width, img_height = img.size
        pos_x = int(evidencia_meta["timestamp_posicao"]["x"] * img_width)
        pos_y = int(evidencia_meta["timestamp_posicao"]["y"] * img_height)
        
        # Configurações do texto
        texto = evidencia_meta["timestamp_texto"]
        texto_cor = evidencia_meta["timestamp_cor"]  # Branco
        fundo_cor = evidencia_meta.get("timestamp_fundo", "#000000B2")  # Preto 70%
        tamanho = evidencia_meta["timestamp_tamanho"]
        
        # Converter cor de fundo para RGBA
        if fundo_cor.startswith("#") and len(fundo_cor) == 9:  # Formato #RRGGBBAA
            r = int(fundo_cor[1:3], 16)
            g = int(fundo_cor[3:5], 16)
            b = int(fundo_cor[5:7], 16)
            a = int(fundo_cor[7:9], 16)
            fundo_rgba = (r, g, b, a)
        else:
            fundo_rgba = (0, 0, 0, 178)  # Fallback: preto 70%
        
        # Usar fonte
        try:
            font = ImageFont.truetype("arial.ttf", tamanho)
        except:
            font = ImageFont.load_default()
        
        # 🔥 CALCULAR TAMANHO DO TEXTO PARA CRIAR FUNDO
        bbox = draw.textbbox((0, 0), texto, font=font)
        texto_largura = bbox[2] - bbox[0]
        texto_altura = bbox[3] - bbox[1]
        
        # 🔥 DEFINIR PADDING E CANTOS ARREDONDADOS
        padding = 10
        borda_radius = 8
        
        # Coordenadas do fundo
        fundo_x1 = pos_x - padding
        fundo_y1 = pos_y - padding
        fundo_x2 = pos_x + texto_largura + padding
        fundo_y2 = pos_y + texto_altura + padding
        
        # 🔥 DESENHAR FUNDO COM CANTOS ARREDONDADOS
        # Criar máscara para cantos arredondados
        mask = Image.new("L", (fundo_x2 - fundo_x1, fundo_y2 - fundo_y1), 0)
        mask_draw = ImageDraw.Draw(mask)
        mask_draw.rounded_rectangle(
            [0, 0, fundo_x2 - fundo_x1, fundo_y2 - fundo_y1],
            radius=borda_radius,
            fill=255
        )
        
        # Aplicar fundo semi-transparente
        fundo_img = Image.new("RGBA", (fundo_x2 - fundo_x1, fundo_y2 - fundo_y1), fundo_rgba)
        img.paste(fundo_img, (fundo_x1, fundo_y1), mask)
        
        # 🔥 DESENHAR TEXTO BRANCO (SEM BORDAS PRETAS)
        draw.text((pos_x, pos_y), texto, fill=texto_cor, font=font)
        
        # Salvar a imagem
        img.save(caminho_imagem)

    # ---------- Nova janela de configuração ----------
    def mostrar_janela_configuracao(self):
        config_window = tk.Toplevel(self.root)
        config_window.title("Configuração de Gravação")
        config_window.geometry("600x600")
        config_window.resizable(False, False)
        
        # 🔥 APLICAR ESTILO À JANELA
        self._apply_style_to_window(config_window)
        
        # 🔥 CORREÇÃO: Usar transient mas SEM grab_set
        config_window.transient(self.root)
        
        # 🔥 CRIAR FRAME PRINCIPAL COM ESTILO
        main_frame = self._create_styled_frame(config_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self._create_styled_label(main_frame, text="PrintF - Configuração de Gravação", 
                                style_type="title").pack(pady=10)
        
        # Seleção de template
        self._create_styled_label(main_frame, text="Selecione o template DOCX:").pack(anchor="w", pady=(10, 5))
        
        template_frame = self._create_styled_frame(main_frame)
        template_frame.pack(fill=tk.X, pady=5)
        
        self.template_var = tk.StringVar()
        template_entry = self._create_styled_entry(template_frame, textvariable=self.template_var, width=40)
        template_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        def selecionar_template():
            template_path = filedialog.askopenfilename(
                title="Selecione o template DOCX",
                filetypes=[("Documentos Word", "*.docx")]
            )
            if template_path:
                self.template_var.set(template_path)
        
        self._create_styled_button(template_frame, text="Procurar", command=selecionar_template).pack(side=tk.RIGHT)
        
        # Seleção de diretório de destino
        self._create_styled_label(main_frame, text="Selecione o diretório de destino:").pack(anchor="w", pady=(20, 5))
        
        dir_frame = self._create_styled_frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        self.dir_var = tk.StringVar()
        dir_entry = self._create_styled_entry(dir_frame, textvariable=self.dir_var, width=40)
        dir_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        def selecionar_diretorio():
            dir_path = filedialog.askdirectory(title="Selecione o diretório para salvar")
            if dir_path:
                self.dir_var.set(dir_path)
        
        self._create_styled_button(dir_frame, text="Procurar", command=selecionar_diretorio).pack(side=tk.RIGHT)
        
        # 🔥 NOVO: Seleção do modo de captura (APENAS 2 OPÇÕES)
        self._create_styled_label(main_frame, text="Modo de Captura da Barra de Tarefas:", 
                                style_type="title").pack(anchor="w", pady=(20, 10))
        
        # Variável para os RadioButtons
        self.modo_captura_var = tk.StringVar(value="ocultar")  # Valor padrão
        
        # Frame para os RadioButtons
        modo_frame = self._create_styled_frame(main_frame)
        modo_frame.pack(fill=tk.X, pady=5)
        
        # RadioButton 1: Manter barra completa
        if self.using_liquid_glass and self.style_manager:
            rb1 = ttk.Radiobutton(
                modo_frame, 
                text="Manter barra de tarefas (data/hora visível na barra do Windows)",
                variable=self.modo_captura_var, 
                value="manter",
                style="Glass.TRadiobutton"
            )
        else:
            rb1 = tk.Radiobutton(
                modo_frame, 
                text="Manter barra de tarefas (data/hora visível na barra do Windows)",
                variable=self.modo_captura_var, 
                value="manter",
                bg='#f5f5f5'
            )
        rb1.pack(anchor="w", pady=2)
        
        # RadioButton 2: Ocultar barra
        if self.using_liquid_glass and self.style_manager:
            rb2 = ttk.Radiobutton(
                modo_frame, 
                text="Ocultar barra de tarefas (data/hora será adicionada na imagem)",
                variable=self.modo_captura_var, 
                value="ocultar",
                style="Glass.TRadiobutton"
            )
        else:
            rb2 = tk.Radiobutton(
                modo_frame, 
                text="Ocultar barra de tarefas (data/hora será adicionada na imagem)",
                variable=self.modo_captura_var, 
                value="ocultar",
                bg='#f5f5f5'
            )
        rb2.pack(anchor="w", pady=2)
        
        # Checkbox para manter evidências
        self._create_styled_label(main_frame, text="Opções de saída:", style_type="title").pack(anchor="w", pady=(20, 10))
        
        # Variável para o checkbox - valor padrão True (marcado)
        self.manter_evidencias_var = tk.BooleanVar(value=True)
        
        # Checkbox
        checkbox_frame = self._create_styled_frame(main_frame)
        checkbox_frame.pack(fill=tk.X, pady=5)
        
        if self.using_liquid_glass and self.style_manager:
            manter_checkbox = ttk.Checkbutton(
                checkbox_frame, 
                text="Manter arquivos de evidência (prints) na pasta após gerar o DOCX",
                variable=self.manter_evidencias_var,
                style="Glass.TCheckbutton"
            )
        else:
            manter_checkbox = tk.Checkbutton(
                checkbox_frame, 
                text="Manter arquivos de evidência (prints) na pasta após gerar o DOCX",
                variable=self.manter_evidencias_var,
                bg='#f5f5f5'
            )
        manter_checkbox.pack(anchor="w")
        
        # Label informativa
        if self.using_liquid_glass and self.style_manager:
            info_label = ttk.Label(
                main_frame, 
                text="Se desmarcado, os arquivos de print serão excluídos após a geração do DOCX.", 
                style="Subtitle.TLabel",
                justify=tk.LEFT
            )
        else:
            info_label = tk.Label(
                main_frame, 
                text="Se desmarcado, os arquivos de print serão excluídos após a geração do DOCX.", 
                font=("Arial", 9), 
                foreground="gray",
                justify=tk.LEFT,
                bg='#f5f5f5'
            )
        info_label.pack(anchor="w", pady=(5, 15))
        
        # Frame para os botões na parte inferior
        btn_frame = self._create_styled_frame(main_frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))
        
        def iniciar_com_config():
            if not self.template_var.get() or not self.dir_var.get():
                messagebox.showerror("Erro", "Por favor, selecione o template e o diretório de destino.")
                return
            
            if not os.path.exists(self.template_var.get()):
                messagebox.showerror("Erro", "O arquivo de template selecionado não existe.")
                return
            
            # 🔥 Armazena a escolha do modo de captura
            self.modo_captura = self.modo_captura_var.get()
            
            # 🔥 VERIFICAÇÃO ADICIONAL: Limpar qualquer estado residual
            self.gravando = False
            self.pausado = False
            self.prints = []
            
            # VALIDAÇÃO SIMPLES: BLOQUEAR APENAS SE TIVER ARQUIVOS
            dir_path = self.dir_var.get()
            
            if os.path.exists(dir_path):
                try:
                    # Verificar se existe algum arquivo na pasta raiz
                    for item in os.listdir(dir_path):
                        item_path = os.path.join(dir_path, item)
                        # Ignorar ocultos e verificar apenas arquivos (não pastas)
                        if not item.startswith('.') and os.path.isfile(item_path):
                            messagebox.showerror(
                                "Arquivos na Pasta", 
                                f"A pasta selecionada contém arquivos.\n\n"
                                f"Para evitar misturar evidências, a pasta deve estar vazia "
                                f"ou conter apenas outras pastas.\n\n"                                
                            )
                            return
                            
                except PermissionError:
                    messagebox.showerror("Erro de Permissão", "Sem permissão para acessar a pasta selecionada.")
                    return
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao verificar a pasta: {str(e)}")
                    return
            
            # Armazenar a escolha do usuário
            self.manter_evidencias = self.manter_evidencias_var.get()
            
            self.template_path = self.template_var.get()
            self.output_dir = self.dir_var.get()
            self.evidence_dir = self.dir_var.get()
            
            # 🔥 CORREÇÃO: Fechar corretamente
            config_window.destroy()  # Usar destroy em vez de apenas fechar
              
            self.iniciar_gravacao()
        
        # Centralizar os botões horizontalmente
        button_container = self._create_styled_frame(btn_frame)
        button_container.pack(expand=True)
        
        self._create_styled_button(button_container, text="Iniciar Gravação", 
                                 command=iniciar_com_config, style_type="accent").pack(side=tk.LEFT, padx=10)
        
        def cancelar_config():
            config_window.destroy()
            
        self._create_styled_button(button_container, text="Cancelar", 
                                 command=cancelar_config, style_type="glass").pack(side=tk.LEFT, padx=10)
        
        # Forçar atualização da interface e ajustar tamanho se necessário
        config_window.update_idletasks()
        
        # Se a janela for muito grande para a tela, ajustar
        screen_width = config_window.winfo_screenwidth()
        screen_height = config_window.winfo_screenheight()
        
        if config_window.winfo_height() > screen_height:
            config_window.geometry(f"600x{screen_height-100}")
        
        # 🔥 CORREÇÃO: Não usar wait_window que pode travar
        return self.template_path is not None and self.output_dir is not None

    def iniciar(self):
        """Inicia o processo de configuração da gravação"""
        # 🔥 CORREÇÃO: Resetar estado ANTES de iniciar
        self.gravando = False
        self.pausado = False
        self.prints = []
        self.evidencia_count = 0
        
        # Mostrar janela de configuração
        if self.mostrar_janela_configuracao():
            print("✅ Configuração concluída, iniciando gravação...")
        else:
            print("❌ Configuração cancelada pelo usuário")

    def pausar(self):
        if self.gravando and not self.pausado:
            self.pausado = True
            messagebox.showinfo("Gravação Pausada", "Gravação pausada. Clique em Retomar para continuar.")
        else:
            messagebox.showwarning("Aviso", "Gravação não está ativa ou já está pausada.")

    def retomar(self):
        if self.gravando and self.pausado:
            self.pausado = False
            messagebox.showinfo("Gravação Retomada", "Gravação retomada. Continue clicando para capturar telas.")
        else:
            messagebox.showwarning("Aviso", "Gravação não está pausada.")

    def finalizar(self):
        if not self.gravando:
            messagebox.showwarning("Aviso", "Nenhuma gravação em andamento.")
            return

        # Parar listener do mouse
        if self.listener_mouse:
            self.listener_mouse.stop()
            self.listener_mouse = None

        self.gravando = False
        self.pausado = False

        # Fechar popup se estiver aberto
        if self.popup and self.popup.winfo_exists():
            self.popup.destroy()
            self.popup = None

        # Gerar documento
        if self.prints:
            try:
                self.gerar_documento()
                messagebox.showinfo("Sucesso", f"Documento gerado com sucesso em:\n{self.output_dir}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar documento: {e}")
        else:
            messagebox.showwarning("Aviso", "Nenhuma evidência capturada.")

        # Limpar estado
        self.prints = []
        self.evidencia_count = 0

    def iniciar_gravacao(self):
        """Inicia a gravação após configuração"""
        # Criar diretório de evidências se não existir
        if not os.path.exists(self.evidence_dir):
            os.makedirs(self.evidence_dir)

        # Inicializar metadados
        self.metadata_path = os.path.join(self.evidence_dir, "evidencias_metadata.json")
        if os.path.exists(self.metadata_path):
            try:
                with open(self.metadata_path, 'r', encoding='utf-8') as f:
                    self.metadata = json.load(f)
            except:
                self.metadata = {"evidencias": [], "proximo_id": 1}
        else:
            self.metadata = {"evidencias": [], "proximo_id": 1}

        # Carregar template
        try:
            if os.path.exists(self.template_path):
                self.doc = Document(self.template_path)
                self.using_template = True
                print("Template carregado com sucesso!")
            else:
                self.doc = Document()
                self.using_template = False
                print("Template não encontrado. Criando documento vazio.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar template: {str(e)}")
            self.doc = Document()
            self.using_template = False

        # Iniciar gravação
        self.gravando = True
        self.pausado = False

        # Configurar listener do mouse
        def on_click(x, y, button, pressed):
            if pressed and button == mouse.Button.left and self.gravando and not self.pausado:
                self.capturar_tela(x, y)

        if self.listener_mouse:
            self.listener_mouse.stop()

        self.listener_mouse = mouse.Listener(on_click=on_click)
        self.listener_mouse.start()

        # Mostrar feedback
        self.mostrar_janela_feedback()
        messagebox.showinfo("Gravação Iniciada", 
                          "✅ Gravação iniciada com sucesso!\n\n"
                          "Clique com o botão esquerdo do mouse para capturar telas.\n\n"
                          "Atalhos disponíveis:\n"
                          "• F6: Pausar gravação\n"
                          "• F7: Retomar gravação\n" 
                          "• F9: Finalizar gravação")

    def capturar_tela(self, x, y):
        """Captura a tela e salva a evidência"""
        try:
            # 🔥 CAPTURA INTELIGENTE BASEADA NO MODO SELECIONADO
            screenshot, (rel_x, rel_y), metodo_utilizado = self.capture_inteligente(x, y)
            
            # Gerar nome único para o arquivo
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"evidencia_{self.metadata['proximo_id']:04d}_{timestamp}.png"
            filepath = os.path.join(self.evidence_dir, filename)
            
            # Salvar a imagem
            screenshot.save(filepath, "PNG")
            
            # 🔥 ADICIONAR METADADOS DA EVIDÊNCIA
            evidencia_meta = {
                "id": self.metadata['proximo_id'],
                "arquivo": filename,
                "timestamp": timestamp,
                "coordenadas": {"x": x, "y": y},
                "coordenadas_relativas": {"x": rel_x, "y": rel_y},
                "metodo_captura": metodo_utilizado,
                "modo_captura": self.modo_captura,
                "comentario": "",
                "excluida": False,
                "timestamp_texto": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "timestamp_cor": "#FFFFFF",  # Branco
                "timestamp_tamanho": 16,
                "timestamp_posicao": {"x": 0.02, "y": 0.02},  # Canto superior esquerdo
                "timestamp_fundo": "#000000B2"  # Preto 70%
            }
            
            # 🔥 APLICAR TIMESTAMP MODERNO
            self.aplicar_timestamp_moderno(filepath, evidencia_meta)
            
            # Atualizar metadados
            self.metadata["evidencias"].append(evidencia_meta)
            self.metadata["proximo_id"] += 1
            self._salvar_metadata()
            
            # Adicionar à lista de prints
            self.prints.append(filepath)
            self.evidencia_count += 1
            
            # Atualizar feedback
            if self.popup and self.popup.winfo_exists():
                try:
                    self.current_index = len(self.prints) - 1
                    self.atualizar_popup()
                except Exception as e:
                    print(f"Erro ao atualizar popup: {e}")
            
            print(f"✅ Captura {self.evidencia_count} salva: {filename}")
            
        except Exception as e:
            print(f"❌ Erro ao capturar tela: {e}")
            messagebox.showerror("Erro", f"Erro ao capturar tela: {e}")

    def mostrar_janela_feedback(self):
        """Mostra janela de feedback durante a gravação"""
        if self.popup and self.popup.winfo_exists():
            self.popup.destroy()
            
        self.popup = tk.Toplevel(self.root)
        self.popup.title("PrintF - Gravando...")
        self.popup.geometry("400x300")
        self.popup.resizable(False, False)
        
        # 🔥 APLICAR ESTILO À JANELA
        self._apply_style_to_window(self.popup)
        
        # 🔥 CORREÇÃO: Usar transient mas SEM grab_set
        self.popup.transient(self.root)
        
        # 🔥 CRIAR FRAME PRINCIPAL COM ESTILO
        main_frame = self._create_styled_frame(self.popup)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self._create_styled_label(main_frame, text="📷 Gravando Evidências", 
                                style_type="title").pack(pady=10)
        
        # Status
        status_frame = self._create_styled_frame(main_frame)
        status_frame.pack(fill=tk.X, pady=10)
        
        self._create_styled_label(status_frame, text="Status:", 
                                style_type="glass").pack(side=tk.LEFT)
        
        self.status_label = self._create_styled_label(status_frame, text="▶ Gravando", 
                                                    style_type="glass")
        self.status_label.pack(side=tk.RIGHT)
        
        # Contador
        count_frame = self._create_styled_frame(main_frame)
        count_frame.pack(fill=tk.X, pady=10)
        
        self._create_styled_label(count_frame, text="Evidências capturadas:", 
                                style_type="glass").pack(side=tk.LEFT)
        
        self.count_label = self._create_styled_label(count_frame, text="0", 
                                                   style_type="glass")
        self.count_label.pack(side=tk.RIGHT)
        
        # Imagem atual
        img_frame = self._create_styled_frame(main_frame)
        img_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self._create_styled_label(img_frame, text="Última captura:", 
                                style_type="glass").pack(anchor="w")
        
        # Container para a imagem (com tamanho fixo)
        img_container = self._create_styled_frame(img_frame)
        img_container.pack(fill=tk.BOTH, expand=True, pady=5)
        img_container.configure(height=150)
        
        self.current_img_label = tk.Label(img_container, bg="white", relief="solid", bd=1)
        self.current_img_label.pack(fill=tk.BOTH, expand=True)
        
        # Comentário
        comment_frame = self._create_styled_frame(main_frame)
        comment_frame.pack(fill=tk.X, pady=10)
        
        self._create_styled_label(comment_frame, text="Comentário (opcional):", 
                                style_type="glass").pack(anchor="w")
        
        self.comment_entry = self._create_styled_entry(comment_frame)
        self.comment_entry.pack(fill=tk.X, pady=5)
        self.comment_entry.bind("<Return>", lambda e: self.adicionar_comentario())
        
        self._create_styled_button(comment_frame, text="Adicionar Comentário", 
                                 command=self.adicionar_comentario, style_type="glass").pack(pady=5)
        
        # Botão finalizar
        btn_frame = self._create_styled_frame(main_frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        self._create_styled_button(btn_frame, text="⏹ Finalizar Gravação", 
                                 command=self.finalizar, style_type="accent").pack()
        
        self.atualizar_popup()

    def atualizar_popup(self):
        """Atualiza o popup de feedback"""
        if not self.popup or not self.popup.winfo_exists():
            return
            
        try:
            # Atualizar status
            status = "⏸ Pausada" if self.pausado else "▶ Gravando"
            self.status_label.config(text=status)
            
            # Atualizar contador
            self.count_label.config(text=str(self.evidencia_count))
            
            # Atualizar imagem se houver capturas
            if self.prints and self.current_index < len(self.prints):
                img_path = self.prints[self.current_index]
                
                # Carregar e redimensionar imagem
                img = Image.open(img_path)
                img.thumbnail((300, 150), Image.Resampling.LANCZOS)
                
                self.current_img_tk = ImageTk.PhotoImage(img)
                self.current_img_label.config(image=self.current_img_tk)
                
        except Exception as e:
            print(f"Erro ao atualizar popup: {e}")

    def adicionar_comentario(self):
        """Adiciona comentário à evidência atual"""
        if not self.prints or self.current_index >= len(self.prints):
            messagebox.showwarning("Aviso", "Nenhuma evidência selecionada.")
            return
            
        comentario = self.comment_entry.get().strip()
        if not comentario:
            messagebox.showwarning("Aviso", "Digite um comentário.")
            return
            
        try:
            # Atualizar metadados
            evidencia_id = self.metadata["evidencias"][self.current_index]["id"]
            for evidencia in self.metadata["evidencias"]:
                if evidencia["id"] == evidencia_id:
                    evidencia["comentario"] = comentario
                    break
                    
            self._salvar_metadata()
            
            # Limpar campo
            self.comment_entry.delete(0, tk.END)
            
            messagebox.showinfo("Sucesso", "Comentário adicionado com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao adicionar comentário: {e}")

    def gerar_documento(self):
        """Gera o documento DOCX com as evidências"""
        try:
            # Carregar template ou criar novo documento
            if self.template_path and os.path.exists(self.template_path):
                self.doc = Document(self.template_path)
                self.using_template = True
            else:
                self.doc = Document()
                self.using_template = False
            
            # Adicionar título se não estiver usando template
            if not self.using_template:
                titulo = self.doc.add_heading('Evidências Capturadas', 0)
                titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Adicionar data e hora
            if not self.using_template:
                data_hora = self.doc.add_paragraph()
                data_hora.add_run(f"Data e hora da geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}").italic = True
                data_hora.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Adicionar evidências
            for i, print_path in enumerate(self.prints, 1):
                # Adicionar título da evidência
                self.doc.add_paragraph().add_run(f"Evidência {i}").bold = True
                
                # Adicionar comentário se existir
                comentario = self.obter_comentario(print_path)
                if comentario:
                    comentario_para = self.doc.add_paragraph()
                    comentario_para.add_run(f"Comentário: {comentario}").italic = True
                
                # Adicionar imagem
                try:
                    paragraph = self.doc.add_paragraph()
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    run = paragraph.add_run()
                    run.add_picture(print_path, width=Inches(6.0))
                except Exception as e:
                    print(f"Erro ao adicionar imagem {print_path}: {e}")
                    self.doc.add_paragraph(f"[Erro ao carregar imagem: {print_path}]")
                
                # Adicionar separador
                self.doc.add_paragraph("―" * 50).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Salvar documento
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            doc_filename = f"evidencias_{timestamp}.docx"
            doc_path = os.path.join(self.output_dir, doc_filename)
            self.doc.save(doc_path)
            
            # 🔥 EXCLUSÃO CONDICIONAL DAS EVIDÊNCIAS
            if not self.manter_evidencias:
                print("🗑️ Excluindo arquivos de evidência conforme solicitado...")
                for print_path in self.prints:
                    try:
                        if os.path.exists(print_path):
                            os.remove(print_path)
                            print(f"🗑️ Excluído: {print_path}")
                    except Exception as e:
                        print(f"⚠️ Erro ao excluir {print_path}: {e}")
                
                # Também excluir o arquivo de metadados
                try:
                    if self.metadata_path and os.path.exists(self.metadata_path):
                        os.remove(self.metadata_path)
                        print(f"🗑️ Excluído: {self.metadata_path}")
                except Exception as e:
                    print(f"⚠️ Erro ao excluir metadados: {e}")
            
            print(f"✅ Documento gerado: {doc_path}")
            
        except Exception as e:
            print(f"❌ Erro ao gerar documento: {e}")
            raise

    def obter_comentario(self, print_path):
        """Obtém o comentário associado a uma evidência"""
        try:
            filename = os.path.basename(print_path)
            for evidencia in self.metadata["evidencias"]:
                if evidencia["arquivo"] == filename:
                    return evidencia.get("comentario", "")
            return ""
        except:
            return ""

    def close(self):
        """Fecha o módulo de forma segura"""
        try:
            # Parar gravação se estiver ativa
            if self.gravando:
                self.finalizar()
            
            # Parar listeners
            if self.listener_mouse:
                try:
                    self.listener_mouse.stop()
                except:
                    pass
                self.listener_mouse = None
                
            if hasattr(self, 'listener_keyboard') and self.listener_keyboard:
                try:
                    self.listener_keyboard.stop()
                except:
                    pass
                self.listener_keyboard = None
            
            # Fechar janelas
            if self.popup and self.popup.winfo_exists():
                try:
                    self.popup.destroy()
                except:
                    pass
                self.popup = None
                
            if self.root and self.root.winfo_exists():
                try:
                    self.root.destroy()
                except:
                    pass
                self.root = None
                
        except Exception as e:
            print(f"Erro ao fechar módulo de captura: {e}")

# Função de compatibilidade para manter a interface existente
def main():
    """Função principal para execução standalone"""
    root = tk.Tk()
    root.withdraw()  # Esconder a janela principal
    
    app = CaptureModule(parent=root)
    app.show()
    
    root.mainloop()

if __name__ == "__main__":
    main()