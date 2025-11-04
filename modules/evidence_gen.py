import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, colorchooser, ttk
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import pyautogui
from pynput import mouse, keyboard
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageFilter  # 🔥 ADICIONADO ImageFilter
from datetime import datetime
import math
import re
import glob
import json
import uuid

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
    """Módulo completo de geração de documentos de evidências"""
    
    def __init__(self, parent=None, settings=None):
        self.parent = parent
        self.root = None
        self.settings = settings or {}
        self.output_dir = os.getcwd()
        self.prints = []
        self.doc = None
        self.using_template = False
        self.template_path = None
        self.current_index = 0
        self.evidence_dir = None
        self.metadata_path = None
        self.metadata = {"evidencias": [], "proximo_id": 1}
        
        self.gravando = False
        self.listener_mouse = None
        self.listener_keyboard = None
        self.popup = None
        self.processamento_cancelado = False
        self.saved_file_path = None
        
        self.using_liquid_glass = STYLES_AVAILABLE and self.settings.get('theme', 'liquid_glass') == 'liquid_glass'
        self.style_manager = LiquidGlassStyle if STYLES_AVAILABLE and self.using_liquid_glass else None
        
        print(f"🎨 EvidenceGenerator - Liquid Glass: {self.using_liquid_glass}")

    def _apply_styles(self, window):
        """Aplica estilos à janela"""
        if self.using_liquid_glass and self.style_manager:
            try:
                self.style_manager.apply_window_style(window)
                return True
            except Exception as e:
                print(f"⚠️ Erro ao aplicar estilos: {e}")
                window.configure(bg='#f5f5f5')
                return False
        else:
            window.configure(bg='#f5f5f5')
            return True

    def _create_styled_frame(self, parent, **kwargs):
        """Cria frame com estilos aplicados"""
        if self.using_liquid_glass and self.style_manager:
            try:
                return self.style_manager.create_glass_frame(parent, **kwargs)
            except Exception as e:
                print(f"⚠️ Erro ao criar frame: {e}")
                return tk.Frame(parent, bg='#f5f5f5', **kwargs)
        else:
            return tk.Frame(parent, bg='#f5f5f5', **kwargs)

    def _create_styled_button(self, parent, text, command, style_type="glass", **kwargs):
        """Cria botão com estilos aplicados"""
        if self.using_liquid_glass and self.style_manager:
            try:
                if style_type == "accent":
                    return self.style_manager.create_accent_button(parent, text, command, **kwargs)
                else:
                    return self.style_manager.create_glass_button(parent, text, command, **kwargs)
            except Exception as e:
                print(f"⚠️ Erro ao criar botão: {e}")
                return self._create_fallback_button(parent, text, command, style_type, **kwargs)
        else:
            return self._create_fallback_button(parent, text, command, style_type, **kwargs)

    def _create_fallback_button(self, parent, text, command, style_type="glass", **kwargs):
        """Cria botão fallback"""
        btn = tk.Button(parent, text=text, command=command, 
                      bg='#3498db' if style_type == "accent" else '#ecf0f1',
                      fg='white' if style_type == "accent" else '#2c3e50',
                      font=("Arial", 10, "bold" if style_type == "accent" else "normal"),
                      relief="flat",
                      cursor="hand2",
                      **kwargs)
        
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
                    label = ttk.Label(parent, text=text, style="Glass.TLabel", **kwargs)
                    return label
            except Exception as e:
                print(f"⚠️ Erro ao criar label: {e}")
                return self._create_fallback_label(parent, text, style_type, **kwargs)
        else:
            return self._create_fallback_label(parent, text, style_type, **kwargs)

    def _create_fallback_label(self, parent, text, style_type="glass", **kwargs):
        """Cria label fallback"""
        bg_color = '#f5f5f5'
        font_config = ("Arial", 14, "bold") if style_type == "title" else ("Arial", 10)
        return tk.Label(parent, text=text, bg=bg_color, fg='#2c3e50', 
                      font=font_config, **kwargs)

    def _create_styled_entry(self, parent, **kwargs):
        """Cria entry com estilos aplicados"""
        if self.using_liquid_glass and self.style_manager:
            try:
                return self.style_manager.create_glass_entry(parent, **kwargs)
            except Exception as e:
                print(f"⚠️ Erro ao criar entry: {e}")
                return tk.Entry(parent, bg='white', fg='#2c3e50', 
                              relief="solid", bd=1, **kwargs)
        else:
            return tk.Entry(parent, bg='white', fg='#2c3e50', 
                          relief="solid", bd=1, **kwargs)

    def _salvar_metadata(self):
        """Salva os metadados no arquivo JSON"""
        if self.metadata_path:
            with open(self.metadata_path, 'w', encoding='utf-8') as f:
                json.dump(self.metadata, f, indent=2, ensure_ascii=False)

    def carregar_evidencias(self, dir_path):
        """Carrega as evidências baseadas nos metadados - SUPORTA MÚLTIPLOS FORMATOS"""
        self.metadata_path = os.path.join(dir_path, "evidencias_metadata.json")
        
        FORMATOS_SUPORTADOS = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.tif']
        
        if os.path.exists(self.metadata_path):
            try:
                with open(self.metadata_path, 'r', encoding='utf-8') as f:
                    self.metadata = json.load(f)
            except:
                self.metadata = {"evidencias": [], "proximo_id": 1}
        else:
            self.metadata = {"evidencias": [], "proximo_id": 1}
            
            for arquivo in os.listdir(dir_path):
                _, ext = os.path.splitext(arquivo)
                if ext.lower() in FORMATOS_SUPORTADOS:
                    caminho_completo = os.path.join(dir_path, arquivo)
                    timestamp = datetime.fromtimestamp(os.path.getmtime(caminho_completo))
                    
                    self.metadata["evidencias"].append({
                        "id": self.metadata["proximo_id"],
                        "arquivo": arquivo,
                        "comentario": "",
                        "timestamp": timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                        "excluida": False
                    })
                    self.metadata["proximo_id"] += 1
            
            self._salvar_metadata()
        
        evidencias_ativas = []
        for evidencia in self.metadata["evidencias"]:
            if not evidencia.get("excluida", False):
                caminho = os.path.join(dir_path, evidencia["arquivo"])
                if os.path.exists(caminho):
                    _, ext = os.path.splitext(evidencia["arquivo"])
                    if ext.lower() in FORMATOS_SUPORTADOS:
                        evidencias_ativas.append(caminho)
        
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
        """Mostra a interface do módulo"""
        if not self.root:
            self._create_interface()
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_set()

    def _create_interface(self):
        """Cria a interface do módulo"""
        self.root = tk.Toplevel(self.parent)
        self.root.title("PrintF - Gerador de Documentos de Evidências")
        self.root.geometry("800x800")
        self.root.resizable(True, False)
        
        self._apply_styles(self.root)
        
        if self.parent:
            self.root.transient(self.parent)
            self.root.grab_set()
        
        main_frame = self._create_styled_frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        self._create_styled_label(main_frame, text="Gerador de Documentos de Evidências", 
                                 style_type="title").pack(pady=20)
        
        desc_text = "Este módulo permite gerar documentos DOCX\na partir de evidências capturadas."
        desc_label = self._create_styled_label(main_frame, text=desc_text)
        desc_label.pack(pady=10)
        
        def iniciar():
            if self.mostrar_janela_configuracao():
                pass
        
        btn_frame = self._create_styled_frame(main_frame)
        btn_frame.pack(pady=20)
        
        self._create_styled_button(btn_frame, text="Iniciar Gerador", 
                                  command=iniciar, style_type="accent", width=20).pack(pady=10)
        
        self._create_styled_button(btn_frame, text="Voltar ao Menu Principal", 
                                  command=self.hide, style_type="glass", width=20).pack(pady=5)

    def hide(self):
        """Esconde a interface do módulo de forma segura"""
        if self.root:
            try:
                if hasattr(self, 'gravando') and self.gravando:
                    self.finalizar()
                
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
                
                if hasattr(self, 'popup') and self.popup and self.popup.winfo_exists():
                    try:
                        self.popup.destroy()
                    except:
                        pass
                    self.popup = None
                
                try:
                    self.root.grab_release()
                except:
                    pass
                    
                self.root.withdraw()
                
                if self.parent and self.parent.winfo_exists():
                    try:
                        self.parent.deiconify()
                        self.parent.lift()
                        self.parent.focus_force()
                    except:
                        pass
                    
            except Exception as e:
                print(f"Erro ao esconder módulo: {e}")
                try:
                    self.root.destroy()
                    self.root = None
                except:
                    pass

    def finalizar(self):
        """Método para finalizar gravação (para compatibilidade)"""
        self.gravando = False

    def mostrar_janela_configuracao(self):
        """Janela de configuração - SIMPLIFICADA SEM LISTA"""
        config_window = tk.Toplevel(self.root)
        config_window.title("Configuração de Arquivo")
        config_window.geometry("1200x900")
        config_window.resizable(True, True)
        config_window.minsize(800,800)
        
        self._apply_styles(config_window)
        
        if self.parent:
            config_window.transient(self.root)
            config_window.grab_set()
        
        main_frame = self._create_styled_frame(config_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=50)
        
        self._create_styled_label(main_frame, text="PrintF - Configuração de Arquivo", 
                                 style_type="title").pack(pady=(0, 10))
        
        instrucoes_text = "Configure as opções abaixo e clique em 'GERAR DOCUMENTO' para iniciar"
        instrucoes_label = self._create_styled_label(main_frame, text=instrucoes_text)
        instrucoes_label.pack(pady=(0, 20))
        if not self.using_liquid_glass:
            instrucoes_label.config(font=("Arial", 9), fg='#7f8c8d')
        
        template_label = self._create_styled_label(main_frame, text="Selecione o template DOCX: *")
        template_label.pack(anchor="w", pady=(10, 5))
        if not self.using_liquid_glass:
            template_label.config(font=("Arial", 10, "bold"))
        
        template_frame = self._create_styled_frame(main_frame)
        template_frame.pack(fill=tk.X, pady=5)
        
        self.template_var = tk.StringVar()
        template_entry = self._create_styled_entry(template_frame, textvariable=self.template_var, width=50)
        template_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        def selecionar_template():
            template_path = filedialog.askopenfilename(
                title="Selecione o template DOCX",
                filetypes=[("Documentos Word", "*.docx")]
            )
            if template_path:
                self.template_var.set(template_path)
        
        btn_template = self._create_styled_button(template_frame, text="Procurar", 
                                                  command=selecionar_template, style_type="glass")
        btn_template.pack(side=tk.RIGHT)
        
        dir_label = self._create_styled_label(main_frame, text="Selecione o diretório onde estão as evidências: *")
        dir_label.pack(anchor="w", pady=(10, 5))
        if not self.using_liquid_glass:
            dir_label.config(font=("Arial", 10, "bold"))
        
        formatos_info = self._create_styled_label(main_frame, text="📷 Formatos suportados: PNG, JPG, JPEG, BMP, GIF, TIFF")
        formatos_info.pack(anchor="w", pady=(0, 5))
        if not self.using_liquid_glass:
            formatos_info.config(font=("Arial", 8), fg='#7f8c8d')
        
        dir_frame = self._create_styled_frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        self.dir_var = tk.StringVar()
        dir_entry = self._create_styled_entry(dir_frame, textvariable=self.dir_var, width=50)
        dir_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        def selecionar_diretorio():
            dir_path = filedialog.askdirectory(title="Selecione o diretório onde estão as evidências")
            if dir_path:
                self.dir_var.set(dir_path)
                try:
                    arquivos = self.carregar_evidencias(dir_path)
                    if arquivos:
                        messagebox.showinfo("Arquivos Encontrados", 
                                          f"✅ {len(arquivos)} arquivo(s) de imagem encontrado(s)\n\n"
                                          f"Clique em 'GERAR DOCUMENTO' para continuar.")
                    else:
                        messagebox.showwarning("Nenhum Arquivo", 
                                             "⚠️ Nenhum arquivo de imagem encontrado neste diretório.\n\n"
                                             "Formatos suportados: PNG, JPG, JPEG, BMP, GIF, TIFF")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao verificar diretório:\n{e}")
        
        btn_dir = self._create_styled_button(dir_frame, text="Procurar", 
                                            command=selecionar_diretorio, style_type="glass")
        btn_dir.pack(side=tk.RIGHT)
        
        info_frame = self._create_styled_frame(main_frame)
        info_frame.pack(fill=tk.X, pady=20)
        
        info_text = "💡 Dica: Os arquivos serão processados em ordem cronológica"
        info_label = self._create_styled_label(info_frame, text=info_text)
        info_label.pack()
        if not self.using_liquid_glass:
            info_label.config(font=("Arial", 9), fg='#7f8c8d', wraplength=600)
        
        separator_frame = self._create_styled_frame(main_frame)
        separator_frame.pack(fill=tk.X, pady=15)
        
        if self.using_liquid_glass and self.style_manager:
            separator = ttk.Separator(separator_frame, orient='horizontal', style="Glass.TSeparator")
        else:
            separator = ttk.Separator(separator_frame, orient='horizontal')
        separator.pack(fill=tk.X)
        
        btn_container = self._create_styled_frame(main_frame)
        btn_container.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        btn_frame = self._create_styled_frame(btn_container)
        btn_frame.pack(anchor="center")
        
        def iniciar_geracao():
            if not self.template_var.get() or not self.dir_var.get():
                messagebox.showerror("Erro", "Por favor, selecione o template e o diretório de evidências.")
                return
            
            if not os.path.exists(self.template_var.get()):
                messagebox.showerror("Erro", "O arquivo de template selecionado não existe.")
                return
            
            if not os.path.exists(self.dir_var.get()):
                messagebox.showerror("Erro", "O diretório de evidências selecionado não existe.")
                return
            
            png_files = self.carregar_evidencias(self.dir_var.get())
            if not png_files:
                messagebox.showerror("Erro", "Nenhuma evidência de imagem encontrada no diretório selecionado.\n\nFormatos suportados: PNG, JPG, JPEG, BMP, GIF, TIFF")
                return
            
            self.template_path = self.template_var.get()
            self.output_dir = self.dir_var.get()
            self.evidence_dir = self.dir_var.get()
            self.prints = png_files
            self.current_index = 0
            
            config_window.destroy()            
            self.iniciar_processamento()
        
        btn_gerar = self._create_styled_button(btn_frame, text="✅ GERAR DOCUMENTO", 
                                              command=iniciar_geracao, style_type="accent", width=25)
        btn_gerar.pack(side=tk.LEFT, padx=10, pady=5)
        
        btn_cancelar = self._create_styled_button(btn_frame, text="❌ CANCELAR", 
                                                  command=config_window.destroy, style_type="glass", width=25)
        btn_cancelar.pack(side=tk.LEFT, padx=10, pady=5)
        
        config_window.update_idletasks()
        
        self.root.wait_window(config_window)
        return self.template_path is not None and self.output_dir is not None and self.prints

    def iniciar_processamento(self):
        """Inicia o processamento das evidências"""
        os.makedirs(self.output_dir, exist_ok=True)

        try:
            if os.path.exists(self.template_path):
                self.doc = Document(self.template_path)
                self.using_template = True
            else:
                self.doc = Document()
                self.using_template = False
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar template: {str(e)}")
            self.doc = Document()
            self.using_template = False
        
        self.gerar_docx()

    def gerar_docx(self):
        """Gera o documento DOCX"""
        documento_salvo = False
        
        while self.current_index < len(self.prints):
            caminho_print = self.prints[self.current_index]
            
            if not os.path.exists(caminho_print):
                if not self.recarregar_evidencias():
                    break
                if self.current_index >= len(self.prints):
                    break
                caminho_print = self.prints[self.current_index]
            
            resultado = self.mostrar_imagem(caminho_print)
            
            if resultado is False:
                break
            elif resultado is None:
                self.recarregar_evidencias()
                continue
            elif resultado == "ja_salvou":
                documento_salvo = True
                break
            else:
                self.current_index += 1
        
        if not documento_salvo:
            self.salvar_docx()

    def mostrar_imagem(self, caminho_print):
        """Mostra popup para adicionar comentário à evidência"""
        popup = tk.Toplevel(self.root)
        popup.title("Adicionar Comentário à Evidência")
        popup.geometry("950x750")
        popup.resizable(False, False)

        self._apply_styles(popup)

        self.processamento_cancelado = False
        resultado = None

        if not os.path.exists(caminho_print):
            messagebox.showerror("Erro", f"Arquivo não encontrado: {os.path.basename(caminho_print)}")
            popup.destroy()
            return None

        img = Image.open(caminho_print)
        img.thumbnail((850, 550))
        img_tk = ImageTk.PhotoImage(img)
        
        img_frame = self._create_styled_frame(popup)
        img_frame.pack(pady=10)
        
        label_img = tk.Label(img_frame, image=img_tk, bg='white')
        label_img.image = img_tk
        label_img.pack()

        comment_frame = self._create_styled_frame(popup)
        comment_frame.pack(pady=5)
        
        comment_label = self._create_styled_label(comment_frame, text="Comentário (opcional):")
        comment_label.pack()
        
        entry = self._create_styled_entry(comment_frame, width=80)
        entry.pack(pady=5)

        info_frame = self._create_styled_frame(popup)
        info_frame.pack(pady=5)
        
        file_info = f"Arquivo: {os.path.basename(caminho_print)}"
        timestamp = datetime.fromtimestamp(os.path.getmtime(caminho_print))
        file_info += f" - {timestamp.strftime('%H:%M:%S')}"
        
        info_label = self._create_styled_label(info_frame, text=file_info)
        info_label.pack()

        def editar_print():
            self.abrir_editor(caminho_print, popup)

        def adicionar():
            nonlocal resultado
            comentario = entry.get()
            
            nome_arquivo = os.path.basename(caminho_print)
            for evidencia in self.metadata["evidencias"]:
                if evidencia["arquivo"] == nome_arquivo:
                    evidencia["comentario"] = comentario
                    break
            self._salvar_metadata()
            
            self.doc.add_picture(caminho_print, width=Inches(5))
            if comentario.strip():
                self.doc.add_paragraph(comentario)
            resultado = True
            popup.destroy()

        def cancelar_processamento():
            if messagebox.askyesno("Confirmar Cancelamento", 
                                  "Tem certeza que deseja cancelar o processamento?"):
                self.processamento_cancelado = True
                popup.destroy()

        def incluir_todos():
            if messagebox.askyesno("Confirmar Inclusão", 
                                  "Deseja incluir todas as evidências restantes sem editar?\nAs evidências serão adicionadas sem comentários."):
                comentario = entry.get()
                
                nome_arquivo = os.path.basename(caminho_print)
                for evidencia in self.metadata["evidencias"]:
                    if evidencia["arquivo"] == nome_arquivo:
                        evidencia["comentario"] = comentario
                        break
                self._salvar_metadata()
                
                self.doc.add_picture(caminho_print, width=Inches(5))
                if comentario.strip():
                    self.doc.add_paragraph(comentario)
                
                for i in range(self.current_index + 1, len(self.prints)):
                    print_path = self.prints[i]
                    if os.path.exists(print_path):
                        nome_arquivo_restante = os.path.basename(print_path)
                        comentario_restante = self.obter_comentario(nome_arquivo_restante)
                        
                        self.doc.add_picture(print_path, width=Inches(5))
                        if comentario_restante.strip():
                            self.doc.add_paragraph(comentario_restante)
                        else:
                            self.doc.add_paragraph("")
                
                self.current_index = len(self.prints)
                self.salvar_docx()
                popup.destroy()
                resultado = "ja_salvou"
            else:
                resultado = False

        def excluir_print():
            nonlocal resultado
            if messagebox.askyesno("Confirmar Exclusão", "Tem certeza que deseja excluir esta evidência?"):
                try:
                    nome_arquivo = os.path.basename(caminho_print)
                    
                    for evidencia in self.metadata["evidencias"]:
                        if evidencia["arquivo"] == nome_arquivo:
                            evidencia["excluida"] = True
                            break
                    
                    self._salvar_metadata()
                    os.remove(caminho_print)
                    print(f"Arquivo excluído: {caminho_print}")
                    
                    resultado = None
                    popup.destroy()
                    
                except Exception as e:
                    print(f"Erro ao excluir arquivo: {e}")
                    messagebox.showerror("Erro", f"Não foi possível excluir o arquivo: {e}")
                    resultado = False
                    popup.destroy()
                    return
            else:
                resultado = False

        acoes_frame = self._create_styled_frame(popup)
        acoes_frame.pack(pady=10)

        self._create_styled_button(acoes_frame, text="✏ Editar Print", 
                                  command=editar_print, style_type="glass", width=15).pack(side=tk.LEFT, padx=5)
        self._create_styled_button(acoes_frame, text="Adicionar e Próximo", 
                                  command=adicionar, style_type="accent", width=15).pack(side=tk.LEFT, padx=5)
        self._create_styled_button(acoes_frame, text="🗑️ Excluir Print", 
                                  command=excluir_print, style_type="glass", width=15).pack(side=tk.LEFT, padx=5)

        controle_frame = self._create_styled_frame(popup)
        controle_frame.pack(pady=10)

        self._create_styled_button(controle_frame, text="❌ Cancelar", 
                                  command=cancelar_processamento, style_type="glass", width=15).pack(side=tk.LEFT, padx=5)
        self._create_styled_button(controle_frame, text="✅ Incluir Todos", 
                                  command=incluir_todos, style_type="accent", width=15).pack(side=tk.LEFT, padx=5)

        def on_closing():
            cancelar_processamento()

        popup.protocol("WM_DELETE_WINDOW", on_closing)
        popup.grab_set()
        self.root.wait_window(popup)
        
        if self.processamento_cancelado:
            return False
        
        return resultado

    def salvar_docx(self):
        """Salva o documento DOCX gerado"""
        if self.template_path:
            nome_base = os.path.basename(self.template_path)
            if nome_base.lower().endswith('.docx'):
                nome_base = nome_base[:-5]
            nome_arquivo = f"{nome_base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        else:
            nome_arquivo = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        diretorio_inicial = self.output_dir if self.output_dir else os.path.expanduser("~")
        caminho_sugerido = os.path.join(diretorio_inicial, nome_arquivo)
        
        caminho_save = filedialog.asksaveasfilename(
            title="Salvar Documento de Evidências",
            initialdir=diretorio_inicial,
            initialfile=nome_arquivo,
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx"), ("Todos os arquivos", "*.*")]
        )
        
        if not caminho_save:
            messagebox.showwarning("Cancelado", "Salvamento cancelado pelo usuário.")
            return
        
        try:
            diretorio_destino = os.path.dirname(caminho_save)
            if not os.path.exists(diretorio_destino):
                os.makedirs(diretorio_destino, exist_ok=True)
            
            self.doc.save(caminho_save)
            self.saved_file_path = caminho_save
            
            def abrir_pasta_apos_mensagem():
                pasta_destino = os.path.dirname(caminho_save)
                try:
                    if os.name == 'nt':
                        os.startfile(pasta_destino)
                    elif os.name == 'posix':
                        import subprocess
                        if sys.platform == 'darwin':
                            subprocess.Popen(['open', pasta_destino])
                        else:
                            subprocess.Popen(['xdg-open', pasta_destino])
                except Exception as e:
                    print(f"Não foi possível abrir a pasta: {e}")
            
            messagebox.showinfo("Concluído", 
                              f"Documento gerado com sucesso!\n\nSalvo em:\n{caminho_save}\n\nA pasta será aberta automaticamente.")
            
            self.root.after(100, abrir_pasta_apos_mensagem)
                
        except PermissionError:
            messagebox.showerror("Erro de Permissão", 
                               f"Não foi possível salvar o arquivo.\n\n"
                               f"O sistema negou permissão para escrever em:\n{caminho_save}\n\n"
                               f"Tente salvar em outro local (como Documentos ou Área de Trabalho).")
        except Exception as e:
            messagebox.showerror("Erro", 
                               f"Erro ao salvar documento:\n\n{str(e)}\n\n"
                               f"Tente salvar em outro local.")

    # 🔥 ADICIONADO: MÉTODOS DE EDIÇÃO COM FERRAMENTA DE MOSAICO (CENSURA)
    def abrir_editor(self, caminho_print, parent):
        """Abre editor de imagens para a evidência"""
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
        
        tools_label = self._create_styled_label(tools_frame, text="Ferramenta:")
        tools_label.pack(side=tk.LEFT, padx=5)
        
        tools_buttons_frame = self._create_styled_frame(tools_frame)
        tools_buttons_frame.pack(side=tk.LEFT, padx=5)
        
        # 🔥 ADICIONADO: Ferramenta de mosaico
        tools = [
            ("rectangle", "⬜", "Retângulo"),
            ("circle", "🔴", "Círculo"),
            ("arrow", "👉", "Seta"),
            ("text", "🆎", "Texto"),
            ("blur", "🌀", "Mosaico")  # NOVA FERRAMENTA
        ]
        
        for tool_value, icon, tooltip in tools:
            if self.using_liquid_glass and self.style_manager:
                btn = ttk.Radiobutton(tools_buttons_frame, text=icon, variable=tool_var, 
                                    value=tool_value, style="Glass.TRadiobutton")
            else:
                btn = tk.Radiobutton(tools_buttons_frame, text=icon, variable=tool_var,
                                   value=tool_value, bg='white', indicatoron=0,
                                   width=3, height=2, relief=tk.RAISED)
            btn.pack(side=tk.LEFT, padx=2)
        
        colors_label = self._create_styled_label(tools_frame, text="Cor:")
        colors_label.pack(side=tk.LEFT, padx=20)
        
        colors_frame = self._create_styled_frame(tools_frame)
        colors_frame.pack(side=tk.LEFT, padx=5)
        
        colors = [("#FF0000", "Vermelho"), ("#0000FF", "Azul"), ("#00FF00", "Verde"), 
                 ("#FFFF00", "Amarelo"), ("#000000", "Preto"), ("#FFFFFF", "Branco")]
        
        for color_value, color_name in colors:
            if self.using_liquid_glass and self.style_manager:
                btn = ttk.Radiobutton(colors_frame, text="⬤", variable=color_var, 
                                    value=color_value, style="Glass.TRadiobutton")
            else:
                btn = tk.Radiobutton(colors_frame, text="⬤", variable=color_var,
                                   value=color_value, bg='white', indicatoron=0,
                                   width=2, height=2, relief=tk.RAISED,
                                   fg=color_value)
            btn.pack(side=tk.LEFT, padx=2)
        
        width_label = self._create_styled_label(tools_frame, text="Espessura:")
        width_label.pack(side=tk.LEFT, padx=20)
        
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
        """Inicia o desenho no canvas"""
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
            # 🔥 ADICIONADO: Caso para mosaico
            elif tool == "blur":
                self.current_element = self.canvas.create_rectangle(
                    self.start_x, self.start_y, self.start_x, self.start_y,
                    outline="#FF00FF", width=2, dash=(5,5)
                )

    def desenhar(self, event, tool):
        """Atualiza o desenho enquanto arrasta"""
        if self.current_element and tool != "text":
            if tool in ["rectangle", "circle", "blur"]:  # 🔥 ADICIONADO: blur
                self.canvas.coords(self.current_element, self.start_x, self.start_y, event.x, event.y)
            elif tool == "arrow":
                self.canvas.coords(self.current_element, self.start_x, self.start_y, event.x, event.y)

    def finalizar_desenho(self, event, tool, color, width):
        """Finaliza o desenho e salva o elemento"""
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
        """Aplica um elemento desenhado na imagem"""
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
        
        # 🔥 ADICIONADO: Aplicar efeito de mosaico
        elif element["type"] == "blur":
            x1, y1, x2, y2 = element["coords"]
            # Ajustar coordenadas para garantir ordem correta
            x1_norm = min(x1, x2)
            y1_norm = min(y1, y2)
            x2_norm = max(x1, x2)
            y2_norm = max(y1, y2)
            
            # Aplicar efeito de blur na área selecionada
            region = self.editing_img.crop((x1_norm, y1_norm, x2_norm, y2_norm))
            # Aumentar o valor do radius para um blur mais forte
            blurred_region = region.filter(ImageFilter.GaussianBlur(15))
            self.editing_img.paste(blurred_region, (x1_norm, y1_norm, x2_norm, y2_norm))

    def atualizar_canvas(self):
        """Atualiza o canvas com a imagem editada"""
        self.display_img = self.editing_img.resize(
            (int(self.editing_img.width * self.scale_factor), 
             int(self.editing_img.height * self.scale_factor)), 
            Image.LANCZOS
        )
        self.current_tk_img = ImageTk.PhotoImage(self.display_img)
        self.canvas.itemconfig(self.canvas_img, image=self.current_tk_img)

    def desfazer_acao(self):
        """Desfaz a última ação de desenho"""
        if self.undo_stack:
            ultimo_elemento = self.undo_stack.pop()
            if ultimo_elemento in self.elements:
                self.elements.remove(ultimo_elemento)
            
            self.editing_img = self.original_img.copy()
            for element in self.elements:
                self.aplicar_elemento_na_imagem(element)
            
            self.atualizar_canvas()

    def salvar_edicao(self, caminho_print, editor):
        """Salva a edição da imagem - PRESERVA FORMATO ORIGINAL"""
        try:
            _, ext = os.path.splitext(caminho_print)
            formato_original = ext.upper().replace('.', '')
            
            formato_map = {
                'JPG': 'JPEG',
                'JPEG': 'JPEG',
                'PNG': 'PNG',
                'BMP': 'BMP',
                'GIF': 'GIF',
                'TIFF': 'TIFF',
                'TIF': 'TIFF'
            }
            
            formato_save = formato_map.get(formato_original, 'PNG')
            
            if formato_save == 'JPEG':
                if self.editing_img.mode in ('RGBA', 'LA', 'P'):
                    background = Image.new('RGB', self.editing_img.size, (255, 255, 255))
                    if self.editing_img.mode == 'P':
                        self.editing_img = self.editing_img.convert('RGBA')
                    background.paste(self.editing_img, mask=self.editing_img.split()[-1])
                    save_img = background
                else:
                    save_img = self.editing_img.convert('RGB')
            elif formato_save == 'PNG':
                save_img = self.editing_img
            elif formato_save in ['BMP', 'GIF', 'TIFF']:
                if self.editing_img.mode == 'RGBA':
                    save_img = self.editing_img.convert('RGB')
                else:
                    save_img = self.editing_img
            else:
                save_img = self.editing_img
            
            save_img.save(caminho_print, formato_save)
            messagebox.showinfo("Sucesso", f"Evidência editada salva com sucesso!\nFormato: {formato_save}")
            editor.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar evidência editada: {str(e)}")


# Modo de execução independente
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