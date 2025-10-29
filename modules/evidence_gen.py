import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, colorchooser, ttk
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import pyautogui
from pynput import mouse, keyboard
from PIL import Image, ImageTk, ImageDraw, ImageFont
from datetime import datetime
import math
import re
import glob
import json
import uuid
import shutil
import subprocess

# Importar sistema de estilos
try:
    from modules.styles import LiquidGlassStyle
    STYLES_AVAILABLE = True
except ImportError:
    try:
        from styles import LiquidGlassStyle
        STYLES_AVAILABLE = True
    except ImportError:
        STYLES_AVAILABLE = False
        print("⚠️ Estilos Liquid Glass não disponíveis, usando fallback")

class EvidenceGeneratorModule:
    """Módulo completo de geração de documentos de evidências a partir de arquivos existentes"""
    
    def __init__(self, parent=None, settings=None):
        self.parent = parent  # Referência à janela principal
        self.root = None      # Janela do módulo
        self.settings = settings or {}
        self.output_dir = os.getcwd()
        self.prints = []            # lista de caminhos das imagens salvas
        self.doc = None
        self.using_template = False
        self.template_path = None
        self.current_index = 0  # controlar o índice atual
        self.evidence_dir = None  # Diretório das evidências
        self.metadata_path = None
        self.metadata = {"evidencias": [], "proximo_id": 1}
        
        # Variáveis necessárias para funcionamento
        self.gravando = False
        self.listener_mouse = None
        self.listener_keyboard = None
        self.popup = None
        self.processamento_cancelado = False
        self.saved_file_path = None
        
        # 🔥 NOVOS ATRIBUTOS PARA COMPATIBILIDADE COM CAPTURE.PY
        self.pausado = False
        self.evidencia_count = 0
        self.pasta_automatica = False
        self.pasta_automatica_path = None
        self.modo_captura = "manter"  # Sempre "manter" para o gerador
        
        # Configuração de estilos
        self.using_liquid_glass = STYLES_AVAILABLE and self.settings.get('theme', 'liquid_glass') == 'liquid_glass'
        self.style_manager = LiquidGlassStyle if STYLES_AVAILABLE else None

        # Atributos para navegação
        self.current_img_label = None
        self.current_img_tk = None
        self.comment_entry = None
        self.pos_label = None

    def _apply_styles(self, window):
        """Aplica estilos à janela"""
        if self.using_liquid_glass and self.style_manager:
            try:
                self.style_manager.apply_window_style(window)
                return True
            except Exception as e:
                print(f"⚠️ Erro ao aplicar estilos: {e}")
                return False
        else:
            # Fallback para estilo padrão
            window.configure(bg='#f5f5f5')
            return True

    def _create_styled_frame(self, parent, **kwargs):
        """Cria frame com estilos aplicados"""
        if self.using_liquid_glass and self.style_manager:
            try:
                return self.style_manager.create_glass_frame(parent, **kwargs)
            except:
                # Fallback se houver erro
                return ttk.Frame(parent, **kwargs)
        else:
            return ttk.Frame(parent, **kwargs)

    def _create_styled_button(self, parent, text, command, style_type="glass", **kwargs):
        """Cria botão com estilos aplicados"""
        if self.using_liquid_glass and self.style_manager:
            try:
                if style_type == "accent":
                    return self.style_manager.create_accent_button(parent, text, command, **kwargs)
                else:
                    return self.style_manager.create_glass_button(parent, text, command, **kwargs)
            except:
                # Fallback se houver erro
                btn = ttk.Button(parent, text=text, command=command, **kwargs)
                return btn
        else:
            # Fallback para botões padrão
            btn = tk.Button(parent, text=text, command=command, 
                          bg='#3498db' if style_type == "accent" else '#ecf0f1',
                          fg='white' if style_type == "accent" else '#2c3e50',
                          font=("Arial", 10, "bold" if style_type == "accent" else "normal"),
                          relief="flat",
                          cursor="hand2",
                          **kwargs)
            
            # Efeitos hover para fallback
            if style_type == "accent":
                btn.bind("<Enter>", lambda e: btn.config(bg='#2980b9'))
                btn.bind("<Leave>", lambda e: btn.config(bg='#3498db'))
            else:
                btn.bind("<Enter>", lambda e: btn.config(bg='#d5dbdb'))
                btn.bind("<Leave>", lambda e: btn.config(bg='#ecf0f1'))
            
            return btn

    def _create_styled_label(self, parent, text, style_type="glass", **kwargs):
        """Cria label com estilos aplicados"""
        if self.using_liquid_glass and self.style_manager:
            try:
                if style_type == "title":
                    return self.style_manager.create_title_label(parent, text, **kwargs)
                else:
                    return ttk.Label(parent, text=text, style="Glass.TLabel", **kwargs)
            except:
                # Fallback se houver erro
                return ttk.Label(parent, text=text, **kwargs)
        else:
            # Fallback para labels padrão
            bg_color = '#f5f5f5'
            font_config = ("Arial", 14, "bold") if style_type == "title" else ("Arial", 10)
            return tk.Label(parent, text=text, bg=bg_color, fg='#2c3e50', 
                          font=font_config, **kwargs)

    def _create_styled_entry(self, parent, **kwargs):
        """Cria entry com estilos aplicados"""
        if self.using_liquid_glass and self.style_manager:
            try:
                return self.style_manager.create_glass_entry(parent, **kwargs)
            except:
                # Fallback se houver erro
                return ttk.Entry(parent, **kwargs)
        else:
            return tk.Entry(parent, bg='white', fg='#2c3e50', 
                          relief="solid", bd=1, **kwargs)

    def _create_styled_listbox(self, parent, **kwargs):
        """Cria Listbox com estilos aplicados para Liquid Glass"""
        if self.using_liquid_glass and self.style_manager:
            # Para Liquid Glass, usar cores escuras
            listbox = tk.Listbox(parent, 
                               bg=self.style_manager.BG_SECONDARY,
                               fg=self.style_manager.TEXT_PRIMARY,
                               selectbackground=self.style_manager.ACCENT_PRIMARY,
                               selectforeground=self.style_manager.TEXT_PRIMARY,
                               insertbackground=self.style_manager.TEXT_PRIMARY,
                               relief="flat",
                               **kwargs)
        else:
            # Fallback para estilo padrão
            listbox = tk.Listbox(parent, 
                               bg='white', 
                               fg='#2c3e50',
                               selectbackground='#3498db',
                               selectforeground='white',
                               relief="solid",
                               bd=1,
                               **kwargs)
        return listbox

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
                print("✅ Metadata carregado do arquivo existente")
            except Exception as e:
                print(f"⚠️ Erro ao carregar metadata: {e}")
                self.metadata = {"evidencias": [], "proximo_id": 1}
        else:
            self.metadata = {"evidencias": [], "proximo_id": 1}
            print("ℹ️ Nenhum arquivo de metadata encontrado")
        
        # Carrega evidências ativas (não excluídas)
        evidencias_ativas = []
        
        # Primeiro: usar metadados se disponíveis
        if self.metadata and "evidencias" in self.metadata:
            for evidencia in self.metadata["evidencias"]:
                if not evidencia.get("excluida", False):
                    caminho = os.path.join(dir_path, evidencia["arquivo"])
                    if os.path.exists(caminho):
                        evidencias_ativas.append(caminho)
        
        # Segundo: buscar todos os PNGs na pasta se não encontrou via metadados
        if not evidencias_ativas:
            print("🔍 Buscando arquivos PNG na pasta...")
            for root_dir, _, files in os.walk(dir_path):
                for file in files:
                    if file.lower().endswith('.png'):
                        full_path = os.path.join(root_dir, file)
                        evidencias_ativas.append(full_path)
            
            # Ordenar por data de modificação
            evidencias_ativas.sort(key=os.path.getmtime)
            
            print(f"✅ Encontrados {len(evidencias_ativas)} arquivos PNG")
        
        return evidencias_ativas

    def recarregar_evidencias(self):
        """Recarrega a lista de evidências"""
        if self.evidence_dir:
            self.prints = self.carregar_evidencias(self.evidence_dir)
            return True
        return False

    def obter_comentario(self, nome_arquivo):
        """Obtém o comentário salvo nos metadados"""
        for evidencia in self.metadata["evidencias"]:
            if evidencia["arquivo"] == nome_arquivo:
                return evidencia.get("comentario", "")
        return ""
    
    def show(self):
        """Mostra a interface do módulo - CORREÇÃO MELHORADA"""
        try:
            if not self.root or not self.root.winfo_exists():
                self._create_interface()
            else:
                self.root.deiconify()
                self.root.lift()
                self.root.focus_set()
            
            # 🔥 CORREÇÃO: Garantir que a janela fique visível SEMPRE
            self._bring_to_front()
            
        except Exception as e:
            print(f"❌ Erro ao mostrar módulo evidence: {e}")
            # Fallback: criar nova interface
            try:
                self._create_interface()
                self._bring_to_front()
            except Exception as e2:
                messagebox.showerror("Erro", f"Falha ao abrir Gerador de Documentos: {e2}")

    def _bring_to_front(self):
        """Trazer janela para frente - CORREÇÃO MELHORADA"""
        if self.root and self.root.winfo_exists():
            try:
                # 🔥 CORREÇÃO: Forçar a janela para frente de forma mais agressiva
                self.root.lift()
                self.root.focus_force()
                self.root.attributes('-topmost', True)
                
                # Atualizar a janela para garantir que está visível
                self.root.update()
                
                # Remover o topmost após um breve período
                self.root.after(500, lambda: self.root.attributes('-topmost', False) if self.root else None)
                
            except Exception as e:
                print(f"⚠️ Erro ao trazer janela para frente: {e}")

    def _create_interface(self):
        """Cria a interface do módulo - IGUAL AO CAPTURE.PY"""
        self.root = tk.Toplevel(self.parent)
        self.root.title("PrintF - Gerador de Documentos de Evidências")
        self.root.geometry("500x400")
        
        # 🔥 APLICAR ESTILO À JANELA
        self._apply_styles(self.root)
        
        # 🔥 CORREÇÃO: Usar protocolo correto para fechar
        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.resizable(False, False)

        # 🔥 CRIAR FRAME PRINCIPAL COM ESTILO
        main_frame = self._create_styled_frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Interface com estilos aplicados
        self._create_styled_label(main_frame, text="📄 PrintF - Gerador de Evidências", 
                                style_type="title").pack(pady=20)
        
        # 🔥 BOTÃO PRINCIPAL: INICIAR GERADOR
        self._create_styled_button(main_frame, text="🚀 Iniciar Gerador", 
                                 command=self.iniciar, style_type="accent").pack(pady=8, fill=tk.X)
        
        # 🔥 ADICIONADO: Descrição do módulo
        desc_label = tk.Label(main_frame, 
                            text="Este módulo permite gerar documentos DOCX a partir de evidências PNG já capturadas.\n\n"
                                 "Selecione o template DOCX e a pasta contendo as evidências para iniciar.",
                            wraplength=400, justify=tk.CENTER, bg='#f5f5f5', fg='#2c3e50')
        desc_label.pack(pady=20)
        
        # 🔥 BOTÃO VOLTAR COM ESTILO
        if self.using_liquid_glass and self.style_manager:
            try:
                voltar_btn = ttk.Button(main_frame, text="⬅ Voltar ao Menu", 
                                      command=self.hide,
                                      style="Back.TButton")
                voltar_btn.pack(pady=15, fill=tk.X)
            except:
                # Fallback se o estilo Error não estiver disponível
                self._create_styled_button(main_frame, text="⬅ Voltar ao Menu", 
                                         command=self.hide, style_type="glass").pack(pady=15, fill=tk.X)
        else:
            voltar_btn = tk.Button(main_frame, text="⬅ Voltar ao Menu", 
                                 command=self.hide,
                                 bg="#e7b13c", fg="white", font=("Arial", 11), relief="flat")
            voltar_btn.pack(pady=15, fill=tk.X)

    def hide(self):
        """Esconde a interface do módulo de forma segura - CORREÇÃO MELHORADA"""
        if self.root:
            try:
                # Parar qualquer gravação em andamento
                if hasattr(self, 'gravando') and self.gravando:
                    self.finalizar()
                
                # Parar listeners
                if hasattr(self, 'listener_keyboard') and self.listener_keyboard:
                    try:
                        self.listener_keyboard.stop()
                    except:
                        pass
                    self.listener_keyboard = None
                
                if hasattr(self, 'listener_mouse') and self.listener_mouse:
                    try:
                        self.listener_mouse.stop()
                    except:
                        pass
                    self.listener_mouse = None
                
                # Fechar janelas secundárias
                if hasattr(self, 'popup') and self.popup and self.popup.winfo_exists():
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
                
                # 🔥 CORREÇÃO CRÍTICA: Restaurar a janela principal de forma mais confiável
                if self.parent and self.parent.winfo_exists():
                    try:
                        self.parent.deiconify()
                        self.parent.lift()
                        self.parent.focus_force()
                        
                        # 🔥 NOVO: Forçar a main para frente
                        self.parent.attributes('-topmost', True)
                        self.parent.after(100, lambda: self.parent.attributes('-topmost', False))
                        
                    except Exception as e:
                        print(f"⚠️ Erro ao restaurar janela principal: {e}")
                    
            except Exception as e:
                print(f"Erro ao esconder módulo: {e}")
                # Fallback: destruir completamente se houver problemas
                try:
                    self.root.destroy()
                    self.root = None
                except:
                    pass

    def finalizar(self):
        """Método para compatibilidade com capture.py"""
        self.gravando = False

    # 🔥 NOVO: MÉTODO INICIAR PARA COMPATIBILIDADE COM CAPTURE.PY
    def iniciar(self):
        """Inicia o processo de configuração do gerador - IGUAL AO CAPTURE.PY"""
        # 🔥 CORREÇÃO: Resetar estado ANTES de iniciar
        self.gravando = False
        self.pausado = False
        self.prints = []
        self.evidencia_count = 0
        self.pasta_automatica = False
        self.pasta_automatica_path = None
        
        # Mostrar janela de configuração
        if self.mostrar_janela_configuracao():
            print("✅ Configuração concluída, iniciando navegação...")
        else:
            print("❌ Configuração cancelada pelo usuário")

    def pausar(self):
        """Método para compatibilidade com capture.py"""
        messagebox.showinfo("Info", "Este módulo não suporta pausa.")

    def retomar(self):
        """Método para compatibilidade com capture.py"""
        messagebox.showinfo("Info", "Este módulo não suporta retomada.")

    # ---------- Nova janela de configuração ----------
    def mostrar_janela_configuracao(self):
        config_window = tk.Toplevel(self.root)
        config_window.title("Configuração do Gerador de Evidências")
        config_window.geometry("600x500")
        config_window.resizable(False, False)
        
        # Aplicar estilos
        self._apply_styles(config_window)
        
        config_window.transient(self.root)
        
        main_frame = self._create_styled_frame(config_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self._create_styled_label(main_frame, text="PrintF - Gerador de Evidências", 
                                 style_type="title").pack(pady=10)

        # 🔥 MODIFICADO: Descrição mais clara
        desc_label = tk.Label(main_frame, 
                            text="Selecione o template DOCX e a pasta contendo as evidências PNG capturadas.",
                            wraplength=500, justify=tk.LEFT, bg='#f5f5f5', fg='#2c3e50')
        desc_label.pack(pady=10)
        
        # Seleção de template
        tk.Label(main_frame, text="Selecione o template DOCX:").pack(anchor="w", pady=(10, 5))
        
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
        
        self._create_styled_button(template_frame, text="Procurar", 
                                  command=selecionar_template).pack(side=tk.RIGHT)
        
        # Seleção de diretório de evidências
        tk.Label(main_frame, text="Selecione a pasta com as evidências PNG:").pack(anchor="w", pady=(20, 5))
        
        dir_frame = self._create_styled_frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        self.dir_var = tk.StringVar()
        dir_entry = self._create_styled_entry(dir_frame, textvariable=self.dir_var, width=40)
        dir_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        def selecionar_diretorio():
            dir_path = filedialog.askdirectory(title="Selecione a pasta com as evidências PNG")
            if dir_path:
                self.dir_var.set(dir_path)
                # 🔥 NOVO: Atualizar lista automaticamente ao selecionar pasta
                atualizar_lista_arquivos(dir_path)
        
        self._create_styled_button(dir_frame, text="Procurar", 
                                  command=selecionar_diretorio).pack(side=tk.RIGHT)

        # 🔥 MODIFICADO: Frame para lista de arquivos com melhor visualização
        list_frame = self._create_styled_frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 5))
        
        tk.Label(list_frame, text="Evidências encontradas:").pack(anchor="w")
        
        # Frame para lista com scrollbar
        list_container = self._create_styled_frame(list_frame)
        list_container.pack(fill=tk.BOTH, expand=True, pady=5)
        
        scrollbar = tk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(list_container, yscrollcommand=scrollbar.set, height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 🔥 NOVO: Label para contagem de arquivos
        self.file_count_label = tk.Label(list_frame, text="Nenhuma evidência encontrada", bg='#f5f5f5', fg='#2c3e50')
        self.file_count_label.pack(anchor="w")
        
        def atualizar_lista_arquivos(dir_path):
            """Atualiza a lista de arquivos PNG encontrados"""
            self.file_listbox.delete(0, tk.END)
            
            if not dir_path or not os.path.exists(dir_path):
                self.file_count_label.config(text="Pasta não encontrada")
                return
                
            # 🔥 MODIFICADO: Buscar por arquivos PNG recursivamente
            png_files = []
            for root_dir, _, files in os.walk(dir_path):
                for file in files:
                    if file.lower().endswith('.png'):
                        png_files.append(os.path.join(root_dir, file))
            
            # Ordenar por data de modificação (mais recentes primeiro)
            png_files.sort(key=os.path.getmtime, reverse=True)
            
            for file_path in png_files:
                filename = os.path.basename(file_path)
                # Mostrar nome do arquivo e data
                timestamp = datetime.fromtimestamp(os.path.getmtime(file_path))
                display_text = f"{filename} - {timestamp.strftime('%d/%m/%Y %H:%M')}"
                self.file_listbox.insert(tk.END, display_text)
            
            if png_files:
                self.file_count_label.config(text=f"{len(png_files)} evidência(s) encontrada(s)")
            else:
                self.file_count_label.config(text="Nenhuma evidência PNG encontrada")
        
        # 🔥 NOVO: Botão para atualizar lista manualmente
        refresh_btn = self._create_styled_button(list_frame, text="🔄 Atualizar Lista", 
                                               command=lambda: atualizar_lista_arquivos(self.dir_var.get()))
        refresh_btn.pack(pady=5)
        
        # Botões de ação
        btn_frame = self._create_styled_frame(main_frame)
        btn_frame.pack(pady=20)
        
        def iniciar_navegacao():
            """Inicia a navegação pelas evidências"""
            if not self.template_var.get():
                messagebox.showerror("Erro", "Por favor, selecione o template DOCX.")
                return
            
            if not self.dir_var.get():
                messagebox.showerror("Erro", "Por favor, selecione a pasta com as evidências.")
                return
            
            template_path = self.template_var.get()
            evidence_dir = self.dir_var.get()
            
            if not os.path.exists(template_path):
                messagebox.showerror("Erro", "O arquivo de template selecionado não existe.")
                return
            
            if not os.path.exists(evidence_dir):
                messagebox.showerror("Erro", "A pasta de evidências selecionada não existe.")
                return
            
            # 🔥 CARREGAR EVIDÊNCIAS COM METADADOS SE EXISTIR
            self.template_path = template_path
            self.output_dir = evidence_dir
            self.evidence_dir = evidence_dir
            
            # Carregar evidências (com suporte a metadata.json se existir)
            self.prints = self.carregar_evidencias(evidence_dir)
            
            if not self.prints:
                messagebox.showerror("Erro", "Nenhuma evidência PNG encontrada na pasta selecionada.")
                return
            
            print(f"✅ {len(self.prints)} evidência(s) carregada(s)")
            
            # 🔥 NOVO: Verificar se metadata foi carregado
            metadata_path = os.path.join(evidence_dir, "evidencias_metadata.json")
            if os.path.exists(metadata_path):
                print("✅ Metadata carregado do arquivo existente")
            else:
                print("ℹ️ Nenhum arquivo de metadata encontrado, operando sem metadados")
            
            config_window.destroy()
            self.mostrar_janela_navegacao()
        
        def cancelar():
            config_window.destroy()
        
        self._create_styled_button(btn_frame, text="🚀 Iniciar Navegação", 
                                  command=iniciar_navegacao, style_type="accent").pack(side=tk.LEFT, padx=10)
        
        self._create_styled_button(btn_frame, text="❌ Cancelar", 
                                  command=cancelar, style_type="glass").pack(side=tk.LEFT, padx=10)
        
        # 🔥 NOVO: Atualizar lista se já houver um diretório selecionado
        if self.dir_var.get():
            atualizar_lista_arquivos(self.dir_var.get())
        
        self.root.wait_window(config_window)
        return len(self.prints) > 0

    def mostrar_janela_navegacao(self):
        """Janela principal de navegação pelas evidências - IGUAL AO CAPTURE.PY"""
        if self.popup and self.popup.winfo_exists():
            self.popup.destroy()

        self.popup = tk.Toplevel(self.root)
        self.popup.title("Navegação de Evidências - Gerador")
        self.popup.geometry("1200x800")
        self.popup.resizable(True, True)
        
        # Aplicar estilo à janela
        self._apply_styles(self.popup)
        
        self.popup.transient(self.root)
        
        # Configurar grid para melhor organização
        self.popup.grid_columnconfigure(0, weight=1)
        self.popup.grid_rowconfigure(0, weight=1)  # A área da imagem expande
        
        # Frame da imagem (maior para melhor visualização)
        img_frame = self._create_styled_frame(self.popup)
        img_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        img_frame.grid_rowconfigure(0, weight=1)
        img_frame.grid_columnconfigure(0, weight=1)
        
        self.current_img_label = tk.Label(img_frame, bg="white")
        self.current_img_label.grid(row=0, column=0, sticky="nsew")
        
        # Frame do comentário (abaixo da imagem)
        comment_frame = self._create_styled_frame(self.popup)
        comment_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
                
        self._create_styled_label(comment_frame, text="Comentário:").pack(anchor="w")
        
        # Criar um frame para o campo de entrada
        comment_entry_frame = self._create_styled_frame(comment_frame)
        comment_entry_frame.pack(fill=tk.X, pady=2)
        
        # Campo de comentário
        self.comment_entry = tk.Entry(comment_entry_frame, font=("Arial", 10))
        self.comment_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.comment_entry.bind("<FocusOut>", lambda e: self.salvar_comentario())
        
        # Frame principal para os botões de navegação e ação
        buttons_main_frame = self._create_styled_frame(self.popup)
        buttons_main_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        
        # Frame para centralizar os botões de navegação
        nav_frame = self._create_styled_frame(buttons_main_frame)
        nav_frame.pack(expand=True, pady=2)
        
        # Botões de navegação (centralizados)
        self._create_styled_button(nav_frame, text="⏮️ Primeira", command=self.primeira_evidencia, 
                                 style_type="glass").pack(side=tk.LEFT, padx=2)
        self._create_styled_button(nav_frame, text="◀️ Anterior", command=self.anterior_evidencia,
                                 style_type="glass").pack(side=tk.LEFT, padx=2)
        
        # Indicador de posição
        self.pos_label = tk.Label(nav_frame, text="", font=("Arial", 12, "bold"))
        self.pos_label.pack(side=tk.LEFT, padx=15)
        
        self._create_styled_button(nav_frame, text="▶️ Próxima", command=self.proxima_evidencia,
                                 style_type="glass").pack(side=tk.LEFT, padx=2)
        self._create_styled_button(nav_frame, text="⏭️ Última", command=self.ultima_evidencia,
                                 style_type="glass").pack(side=tk.LEFT, padx=2)
        
        # Pular para específica
        self._create_styled_button(nav_frame, text="🔢 Ir para...", command=self.ir_para_especifica,
                                 style_type="glass").pack(side=tk.LEFT, padx=2)
        
        # Botões de ação no mesmo nível (Editar e Excluir Print)
        action_frame = self._create_styled_frame(buttons_main_frame)
        action_frame.pack(expand=True, pady=2)
        
        self._create_styled_button(action_frame, text="✏️ Editar Print", command=self.editar_evidencia_atual,
                                 style_type="glass").pack(side=tk.LEFT, padx=5)
        self._create_styled_button(action_frame, text="🗑️ Excluir Print", command=self.excluir_evidencia_atual,
                                 style_type="glass").pack(side=tk.LEFT, padx=5)
        
        # Frame de controle (parte inferior)
        control_frame = self._create_styled_frame(self.popup)
        control_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        
        # Frame para centralizar os botões de controle
        control_buttons_frame = self._create_styled_frame(control_frame)
        control_buttons_frame.pack(expand=True)
        
        # Botões na ordem solicitada: Cancelar primeiro, depois Gerar Evidência
        self._create_styled_button(control_buttons_frame, text="❌ Cancelar", command=self.cancelar_processamento,
                                 style_type="error").pack(side=tk.LEFT, padx=5)
        
        # 🔥 ADICIONADO: Botão Voltar ao Menu Principal na navegação
        self._create_styled_button(control_buttons_frame, text="🏠 Voltar ao Menu Principal", command=self.hide,
                                 style_type="glass").pack(side=tk.LEFT, padx=5)
        
        self._create_styled_button(control_buttons_frame, text="✅ Gerar Documento", command=self.finalizar_processamento,
                                 style_type="accent").pack(side=tk.LEFT, padx=5)
        
        # Carregar primeira evidência
        self.current_index = 0
        self.atualizar_exibicao()
        
        self.popup.protocol("WM_DELETE_WINDOW", self.cancelar_processamento)

    def atualizar_exibicao(self):
        """Atualiza a exibição da evidência atual"""
        if not self.prints or self.current_index >= len(self.prints):
            return
            
        caminho_print = self.prints[self.current_index]
        
        try:
            # Carrega e exibe a imagem com tamanho maior
            img = Image.open(caminho_print)
            
            # Obter o tamanho da área disponível para a imagem
            self.popup.update()
            available_width = self.popup.winfo_width() - 40  # Margens
            available_height = self.popup.winfo_height() - 250  # Espaço para controles
            
            # Ajustar a imagem para caber na área disponível
            img.thumbnail((available_width, available_height))
            self.current_img_tk = ImageTk.PhotoImage(img)
            self.current_img_label.config(image=self.current_img_tk)
            
            # Atualiza indicador de posição
            self.pos_label.config(text=f"Evidência {self.current_index + 1} de {len(self.prints)}")
            
            # Carrega comentário salvo
            nome_arquivo = os.path.basename(caminho_print)
            comentario = self.obter_comentario(nome_arquivo)
            self.comment_entry.delete(0, tk.END)
            self.comment_entry.insert(0, comentario)
            
        except Exception as e:
            print(f"Erro ao carregar imagem: {e}")

    def salvar_comentario(self):
        """Salva o comentário da evidência atual"""
        if not self.prints or self.current_index >= len(self.prints):
            return
            
        caminho_print = self.prints[self.current_index]
        nome_arquivo = os.path.basename(caminho_print)
        comentario = self.comment_entry.get()
        
        # Atualiza metadados
        for evidencia in self.metadata["evidencias"]:
            if evidencia["arquivo"] == nome_arquivo:
                evidencia["comentario"] = comentario
                break
                
        self._salvar_metadata()        

    # Métodos de navegação
    def primeira_evidencia(self):
        self.salvar_comentario()
        self.current_index = 0
        self.atualizar_exibicao()

    def anterior_evidencia(self):
        self.salvar_comentario()
        if self.current_index > 0:
            self.current_index -= 1
            self.atualizar_exibicao()

    def proxima_evidencia(self):
        self.salvar_comentario()
        if self.current_index < len(self.prints) - 1:
            self.current_index += 1
            self.atualizar_exibicao()

    def ultima_evidencia(self):
        self.salvar_comentario()
        self.current_index = len(self.prints) - 1
        self.atualizar_exibicao()

    def ir_para_especifica(self):
        self.salvar_comentario()
        if not self.prints:
            return
            
        numero = simpledialog.askinteger("Navegar", 
                                       f"Digite o número da evidência (1-{len(self.prints)}):",
                                       minvalue=1, maxvalue=len(self.prints))
        if numero:
            self.current_index = numero - 1
            self.atualizar_exibicao()

    def editar_evidencia_atual(self):
        self.salvar_comentario()
        if not self.prints or self.current_index >= len(self.prints):
            return
            
        caminho_print = self.prints[self.current_index]
        self.abrir_editor(caminho_print, self.popup)
        self.atualizar_exibicao()

    def excluir_evidencia_atual(self):
        self.salvar_comentario()
        if not self.prints or self.current_index >= len(self.prints):
            return
            
        caminho_print = self.prints[self.current_index]
        nome_arquivo = os.path.basename(caminho_print)
        
        if messagebox.askyesno("Confirmar Exclusão", 
                             "Tem certeza que deseja excluir este print?"):
            try:
                os.remove(caminho_print)
                
                for evidencia in self.metadata["evidencias"]:
                    if evidencia["arquivo"] == nome_arquivo:
                        evidencia["excluida"] = True
                        break
                
                self._salvar_metadata()
                self.recarregar_evidencias()
                
                if not self.prints:
                    messagebox.showinfo("Info", "Todas as evidências foram processadas.")
                    self.finalizar_processamento()
                    return
                
                if self.current_index >= len(self.prints):
                    self.current_index = len(self.prints) - 1
                
                self.atualizar_exibicao()
                messagebox.showinfo("Sucesso", "Evidência excluída!")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao excluir: {str(e)}")

    def finalizar_processamento(self):
        """Processa todas as evidências e gera o DOCX"""
        self.salvar_comentario()
        
        try:
            doc_path = self.gerar_documento()
            
            pasta_para_abrir = os.path.dirname(doc_path)
            
            resposta = messagebox.askyesno(
                "Sucesso", 
                f"Documento gerado com sucesso em:\n{doc_path}\n\nDeseja abrir a pasta onde o documento foi salvo?",
                parent=self.popup
            )
            
            if resposta:
                if not self._abrir_pasta(pasta_para_abrir):
                    messagebox.showinfo(
                        "Abrir Pasta", 
                        f"Pasta do documento:\n{pasta_para_abrir}",
                        parent=self.popup
                    )
                    
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar documento: {e}", parent=self.popup)
        
        if self.popup and self.popup.winfo_exists():
            self.popup.destroy()
            self.popup = None

    def _abrir_pasta(self, caminho_pasta):
        """Abre a pasta no explorador de arquivos do sistema"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(caminho_pasta)
            elif os.name == 'posix':  # Linux ou macOS
                if sys.platform == 'darwin':  # macOS
                    subprocess.run(['open', caminho_pasta])
                else:  # Linux
                    subprocess.run(['xdg-open', caminho_pasta])
            return True
        except Exception as e:
            print(f"Erro ao abrir pasta: {e}")
            return False

    def cancelar_processamento(self):
        self.salvar_comentario()
        if messagebox.askyesno("Confirmar", "Deseja cancelar o processamento?"):
            if self.popup:
                self.popup.destroy()
                self.popup = None

    def gerar_documento(self):
        """Gera o documento DOCX com as evidências e retorna o caminho do documento"""
        doc_path = None
        try:
            print("🔄 Iniciando geração do documento DOCX...")
            
            if self.template_path and os.path.exists(self.template_path):
                self.doc = Document(self.template_path)
                self.using_template = True
                print(f"✅ Template carregado: {self.template_path}")
            else:
                self.doc = Document()
                self.using_template = False
                print("ℹ️ Criando documento vazio (sem template)")
            
            if not self.using_template:
                titulo = self.doc.add_heading('Evidências Capturadas', 0)
                titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            if not self.using_template:
                data_hora = self.doc.add_paragraph()
                data_hora.add_run(f"Data e hora da geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}").italic = True
                data_hora.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            for i, print_path in enumerate(self.prints, 1):
                print(f"📷 Adicionando evidência {i}: {print_path}")
                
                self.doc.add_paragraph().add_run(f"Evidência {i}").bold = True
                
                nome_arquivo = os.path.basename(print_path)
                comentario = self.obter_comentario(nome_arquivo)
                if comentario:
                    comentario_para = self.doc.add_paragraph()
                    comentario_para.add_run(f"Comentário: {comentario}").italic = True
                
                try:
                    paragraph = self.doc.add_paragraph()
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    run = paragraph.add_run()
                    
                    if os.path.exists(print_path):
                        run.add_picture(print_path, width=Inches(6.0))
                        print(f"✅ Imagem {i} adicionada com sucesso")
                    else:
                        print(f"⚠️ Arquivo não encontrado: {print_path}")
                        self.doc.add_paragraph(f"[Arquivo de imagem não encontrado: {print_path}]")
                        
                except Exception as e:
                    print(f"❌ Erro ao adicionar imagem {print_path}: {e}")
                    self.doc.add_paragraph(f"[Erro ao carregar imagem: {print_path}]")
                
                self.doc.add_paragraph("―" * 50).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            template_filename = os.path.basename(self.template_path)
            template_name = os.path.splitext(template_filename)[0]
            
            template_name = self._limpar_nome_arquivo(template_name)
            
            doc_filename = f"{template_name}_{timestamp}.docx"
            doc_path = os.path.join(self.output_dir, doc_filename)
            
            os.makedirs(os.path.dirname(doc_path), exist_ok=True)
            
            if len(doc_path) > 255:
                short_name = f"Evidencias_{timestamp}.docx"
                doc_path = os.path.join(self.output_dir, short_name)
                print(f"⚠️ Caminho muito longo, usando nome reduzido: {short_name}")
            
            self.doc.save(doc_path)
            print(f"✅ Documento salvo em: {doc_path}")
            
            return doc_path
            
        except Exception as e:
            print(f"❌ Erro ao gerar documento: {e}")
            import traceback
            traceback.print_exc()
            raise

    def _limpar_nome_arquivo(self, nome):
        """Remove caracteres inválidos para nomes de arquivo no Windows, mantendo caracteres PT-BR"""
        caracteres_invalidos = r'[\\/*?:"<>|]'
        nome_limpo = re.sub(caracteres_invalidos, '_', nome)
        
        nome_limpo = re.sub(r'[^\w\s\-\.\(\)áàâãéèêíïóôõöúçñÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]', '', nome_limpo)
        
        if len(nome_limpo) > 100:
            nome_limpo = nome_limpo[:100]
            
        return nome_limpo.strip()

    # ---------- Editor de prints ----------
    def abrir_editor(self, caminho_print, parent):
        editor = tk.Toplevel(parent)
        editor.title("Editor de Evidência")
        editor.geometry("1200x800")
        
        self._apply_styles(editor)
        
        main_frame = self._create_styled_frame(editor)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        tools_frame = self._create_styled_frame(main_frame)
        tools_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        
        canvas_frame = self._create_styled_frame(main_frame)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.original_img = Image.open(caminho_print).convert("RGBA")
        img_w, img_h = self.original_img.size
        
        max_w, max_h = 1000, 700
        scale = min(max_w / img_w, max_h / img_h)
        self.scale_factor = scale
        disp_w, disp_h = int(img_w * scale), int(img_h * scale)
        
        self.editing_img = self.original_img.copy()
        self.display_img = self.editing_img.resize((disp_w, disp_h), Image.LANCZOS)

        self.current_tk_img = ImageTk.PhotoImage(self.display_img)
        self.elements = []
        self.undo_stack = []
        self.temp_element = None
        
        canvas_bg = 'gray'
        self.canvas = tk.Canvas(canvas_frame, width=disp_w, height=disp_h, cursor="cross", bg=canvas_bg)
        self.canvas.pack(padx=5, pady=5)
        self.canvas_img = self.canvas.create_image(0, 0, anchor="nw", image=self.current_tk_img)
        
        tool_var = tk.StringVar(value="rectangle")
        color_var = tk.StringVar(value="#FF0000")
        width_var = tk.IntVar(value=3)
        
        if self.using_liquid_glass:
            ttk.Label(tools_frame, text="Ferramenta:", 
                     style="Glass.TLabel").pack(side=tk.LEFT, padx=5)
        else:
            tk.Label(tools_frame, text="Ferramenta:", 
                    bg='#f5f5f5', fg='#2c3e50').pack(side=tk.LEFT, padx=5)
        
        tools_buttons_frame = self._create_styled_frame(tools_frame)
        tools_buttons_frame.pack(side=tk.LEFT, padx=5)
        
        tools = [
            ("rectangle", "⬜", "Retângulo"),
            ("circle", "🔴", "Círculo"),
            ("arrow", "👉", "Seta"),
            ("text", "🆎", "Texto")
        ]
        
        for tool_value, icon, tooltip in tools:
            if self.using_liquid_glass:
                btn = ttk.Radiobutton(tools_buttons_frame, text=icon, variable=tool_var, 
                                    value=tool_value, style="Glass.TRadiobutton")
            else:
                btn = tk.Radiobutton(tools_buttons_frame, text=icon, variable=tool_var,
                                   value=tool_value, bg='white', indicatoron=0,
                                   width=3, height=2, relief=tk.RAISED)
            btn.pack(side=tk.LEFT, padx=2)
        
        if self.using_liquid_glass:
            ttk.Label(tools_frame, text="Cor:", style="Glass.TLabel").pack(side=tk.LEFT, padx=20)
        else:
            tk.Label(tools_frame, text="Cor:", bg='#f5f5f5', fg='#2c3e50').pack(side=tk.LEFT, padx=20)
        
        colors_frame = self._create_styled_frame(tools_frame)
        colors_frame.pack(side=tk.LEFT, padx=5)
        
        colors = [("#FF0000", "Vermelho"), ("#0000FF", "Azul"), ("#00FF00", "Verde"), 
                 ("#FFFF00", "Amarelo"), ("#000000", "Preto"), ("#FFFFFF", "Branco")]
        
        for color_value, color_name in colors:
            if self.using_liquid_glass:
                btn = ttk.Radiobutton(colors_frame, text="⬤", variable=color_var, 
                                    value=color_value, style="Glass.TRadiobutton")
            else:
                btn = tk.Radiobutton(colors_frame, text="⬤", variable=color_var,
                                   value=color_value, bg='white', indicatoron=0,
                                   width=2, height=2, relief=tk.RAISED,
                                   fg=color_value)
            btn.pack(side=tk.LEFT, padx=2)
        
        if self.using_liquid_glass:
            ttk.Label(tools_frame, text="Espessura:", style="Glass.TLabel").pack(side=tk.LEFT, padx=20)
        else:
            tk.Label(tools_frame, text="Espessura:", bg='#f5f5f5', fg='#2c3e50').pack(side=tk.LEFT, padx=20)
        
        width_scale = tk.Scale(tools_frame, from_=1, to=10, variable=width_var, 
                              orient=tk.HORIZONTAL, length=100, showvalue=True)
        width_scale.pack(side=tk.LEFT, padx=5)
        
        action_frame = self._create_styled_frame(tools_frame)
        action_frame.pack(side=tk.RIGHT, padx=10)
        
        self._create_styled_button(action_frame, text="↶ Desfazer", 
                                  command=self.desfazer_acao, style_type="glass").pack(side=tk.LEFT, padx=2)
        self._create_styled_button(action_frame, text="Salvar", 
                                  command=lambda: self.salvar_edicao(caminho_print, editor), 
                                  style_type="accent").pack(side=tk.LEFT, padx=2)
        self._create_styled_button(action_frame, text="Cancelar", 
                                  command=editor.destroy, style_type="glass").pack(side=tk.LEFT, padx=2)
        
        self.start_x = None
        self.start_y = None
        self.current_element = None
        
        self.canvas.bind("<Button-1>", lambda e: self.iniciar_desenho(e, tool_var.get()))
        self.canvas.bind("<B1-Motion>", lambda e: self.desenhar(e, tool_var.get()))
        self.canvas.bind("<ButtonRelease-1>", lambda e: self.finalizar_desenho(e, tool_var.get(), color_var.get(), width_var.get()))
        
        editor.transient(parent)
        editor.grab_set()

    def iniciar_desenho(self, event, tool):
        self.start_x = event.x
        self.start_y = event.y
        
        if tool == "text":
            texto = simpledialog.askstring("Texto", "Digite o texto:", parent=self.root)
            if texto:
                orig_x = int(event.x / self.scale_factor)
                orig_y = int(event.y / self.scale_factor)
                
                element_data = {
                    "type": "text",
                    "text": texto,
                    "x": orig_x,
                    "y": orig_y,
                    "color": "#FF0000",
                    "size": 20
                }
                self.elements.append(element_data)
                self.aplicar_elemento_na_imagem(element_data)
                self.atualizar_canvas()
        else:
            if tool == "rectangle":
                self.current_element = self.canvas.create_rectangle(
                    self.start_x, self.start_y, self.start_x, self.start_y,
                    outline="#FF0000", width=3
                )
            elif tool == "circle":
                self.current_element = self.canvas.create_oval(
                    self.start_x, self.start_y, self.start_x, self.start_y,
                    outline="#FF0000", width=3
                )
            elif tool == "arrow":
                self.current_element = self.canvas.create_line(
                    self.start_x, self.start_y, self.start_x, self.start_y,
                    arrow=tk.LAST, fill="#FF0000", width=3
                )

    def desenhar(self, event, tool):
        if self.current_element and tool != "text":
            if tool in ["rectangle", "circle"]:
                self.canvas.coords(self.current_element, self.start_x, self.start_y, event.x, event.y)
            elif tool == "arrow":
                self.canvas.coords(self.current_element, self.start_x, self.start_y, event.x, event.y)

    def finalizar_desenho(self, event, tool, color, width):
        if self.current_element and tool != "text":
            coords = self.canvas.coords(self.current_element)
            orig_coords = [int(coord / self.scale_factor) for coord in coords]
            
            element_data = {
                "type": tool,
                "coords": orig_coords,
                "color": color,
                "width": width
            }
            self.elements.append(element_data)
            self.undo_stack.append(element_data.copy())
            
            self.aplicar_elemento_na_imagem(element_data)
            self.atualizar_canvas()
            
            self.current_element = None

    def aplicar_elemento_na_imagem(self, element):
        draw = ImageDraw.Draw(self.editing_img)
        
        if element["type"] == "rectangle":
            x1, y1, x2, y2 = element["coords"]
            draw.rectangle([x1, y1, x2, y2], outline=element["color"], width=element["width"])
        
        elif element["type"] == "circle":
            x1, y1, x2, y2 = element["coords"]
            draw.ellipse([x1, y1, x2, y2], outline=element["color"], width=element["width"])
        
        elif element["type"] == "arrow":
            x1, y1, x2, y2 = element["coords"]
            draw.line([x1, y1, x2, y2], fill=element["color"], width=element["width"])
            
            arrow_size = element["width"] * 3
            angle = math.atan2(y2 - y1, x2 - x1)
            
            ax1 = x2 - arrow_size * math.cos(angle - math.pi/6)
            ay1 = y2 - arrow_size * math.sin(angle - math.pi/6)
            ax2 = x2 - arrow_size * math.cos(angle + math.pi/6)
            ay2 = y2 - arrow_size * math.sin(angle + math.pi/6)
            
            draw.line([x2, y2, ax1, ay1], fill=element["color"], width=element["width"])
            draw.line([x2, y2, ax2, ay2], fill=element["color"], width=element["width"])
        
        elif element["type"] == "text":
            try:
                font = ImageFont.truetype("arial.ttf", element["size"])
            except:
                font = ImageFont.load_default()
            draw.text((element["x"], element["y"]), element["text"], fill=element["color"], font=font)

    def atualizar_canvas(self):
        self.display_img = self.editing_img.resize(
            (int(self.editing_img.width * self.scale_factor), 
             int(self.editing_img.height * self.scale_factor)), 
            Image.LANCZOS
        )
        self.current_tk_img = ImageTk.PhotoImage(self.display_img)
        self.canvas.itemconfig(self.canvas_img, image=self.current_tk_img)

    def desfazer_acao(self):
        if self.undo_stack:
            ultimo_elemento = self.undo_stack.pop()
            if ultimo_elemento in self.elements:
                self.elements.remove(ultimo_elemento)
            
            self.editing_img = self.original_img.copy()
            for element in self.elements:
                self.aplicar_elemento_na_imagem(element)
            
            self.atualizar_canvas()

    def salvar_edicao(self, caminho_print, editor):
        try:
            if self.editing_img.mode == 'RGBA':
                save_img = self.editing_img.convert('RGB')
            else:
                save_img = self.editing_img
            
            save_img.save(caminho_print, "PNG")
            messagebox.showinfo("Sucesso", "Evidência editada salva com sucesso!")
            editor.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar evidência editada: {str(e)}")


# Modo de execução independente (para teste)
if __name__ == "__main__":
    root = tk.Tk()
    root.title("PrintF - Gerador de Evidências")
    root.geometry("400x200")
    root.resizable(False, False)
    
    try:
        root.configure(bg='#f5f5f5')
    except:
        pass
    
    root.eval('tk::PlaceWindow . center')
    
    main_frame = tk.Frame(root, bg='#f5f5f5', padx=30, pady=30)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    title_label = tk.Label(main_frame, text="PrintF - Gerador de Evidências", 
                         font=("Arial", 16, "bold"), bg='#f5f5f5', fg='#2c3e50')
    title_label.pack(pady=20)
    
    def iniciar_gerador():
        gerador = EvidenceGeneratorModule(root)
        gerador.show()
    
    start_btn = tk.Button(main_frame, text="Iniciar Gerador de Evidências", 
                         command=iniciar_gerador, width=25,
                         bg='#3498db', fg='white', font=("Arial", 12, "bold"),
                         relief="flat", cursor="hand2")
    start_btn.pack(pady=10)
    
    start_btn.bind("<Enter>", lambda e: start_btn.config(bg='#2980b9'))
    start_btn.bind("<Leave>", lambda e: start_btn.config(bg='#3498db'))
    
    exit_btn = tk.Button(main_frame, text="Sair", command=root.quit, width=15,
                        bg='#e74c3c', fg='white', font=("Arial", 10),
                        relief="flat", cursor="hand2")
    exit_btn.pack(pady=10)
    
    exit_btn.bind("<Enter>", lambda e: exit_btn.config(bg='#c0392b'))
    exit_btn.bind("<Leave>", lambda e: exit_btn.config(bg='#e74c3c'))
    
    root.mainloop()