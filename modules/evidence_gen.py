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
        
        # CORREÇÃO: Extensões suportadas
        self.supported_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp')
        
        # Configuração de estilos - CORREÇÃO: Verificar tema nas settings
        self.using_liquid_glass = STYLES_AVAILABLE and self.settings.get('theme', 'liquid_glass') == 'liquid_glass'
        self.style_manager = LiquidGlassStyle if STYLES_AVAILABLE else None

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
        """Carrega as evidências baseadas nos metadados - SUPORTA MÚLTIPLOS FORMATOS"""
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
        
        # Se não houver metadados, busca arquivos de imagem no diretório
        if not evidencias_ativas:
            for arquivo in os.listdir(dir_path):
                if arquivo.lower().endswith(self.supported_extensions):
                    caminho = os.path.join(dir_path, arquivo)
                    evidencias_ativas.append(caminho)
                    # Adiciona aos metadados
                    self.metadata["evidencias"].append({
                        "id": self.metadata["proximo_id"],
                        "arquivo": arquivo,
                        "comentario": "",
                        "excluida": False
                    })
                    self.metadata["proximo_id"] += 1
            
            # Salva metadados se foram criados
            if evidencias_ativas:
                self._salvar_metadata()
        
        # Ordena por timestamp
        evidencias_ativas.sort(key=lambda x: os.path.getmtime(x))
        
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
        """Mostra a interface do módulo - CORREÇÃO PARA EVITAR MINIMIZAÇÃO"""
        if not self.root:
            self._create_interface()
        
        # 🔥 CORREÇÃO: Garantir que a janela fique visível corretamente
        self.root.deiconify()  # Garante que a janela não está minimizada
        self.root.lift()       # Traz para frente
        self.root.focus_set()  # Define o foco
        
        # 🔥 CORREÇÃO ADICIONAL: Remover qualquer atributo que force minimização
        try:
            self.root.attributes('-zoomed', False)  # Remove estado maximizado se existir
            self.root.state('normal')               # Garante estado normal
        except:
            pass

    def _center_on_parent(self):
        """Centraliza a janela do módulo em relação à janela principal"""
        if self.parent:
            try:
                self.root.update_idletasks()
                parent_x = self.parent.winfo_x()
                parent_y = self.parent.winfo_y()
                parent_width = self.parent.winfo_width()
                parent_height = self.parent.winfo_height()
                
                width = 500
                height = 300
                
                x = parent_x + (parent_width - width) // 2
                y = parent_y + (parent_height - height) // 2
                
                self.root.geometry(f"{width}x{height}+{x}+{y}")
            except Exception as e:
                print(f"⚠️ Erro ao centralizar janela: {e}")
                # Fallback: centralizar na tela
                self.root.eval('tk::PlaceWindow . center')

    def _create_interface(self):
        """Cria a interface do módulo - CORREÇÃO PARA EVITAR PROBLEMAS DE FOCUS"""
        self.root = tk.Toplevel(self.parent)
        self.root.title("PrintF - Gerador de Documentos de Evidências")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        
        # 🔥 CORREÇÃO: Configurar para não minimizar automaticamente
        self.root.transient(self.parent)  # Define como janela filha
        self.root.grab_set()              # Mantém o foco
        
        # Aplicar estilos
        self._apply_styles(self.root)
        
        # Centralizar na tela principal
        self._center_on_parent()
        
        main_frame = self._create_styled_frame(self.root, padding=30)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self._create_styled_label(main_frame, text="Gerador de Documentos de Evidências", 
                                 style_type="title").pack(pady=20)
        
        # Label descritivo
        if self.using_liquid_glass:
            desc_label = ttk.Label(main_frame, text="Este módulo permite gerar documentos DOCX a partir de evidências capturadas.",
                                 style="Glass.TLabel")
        else:
            desc_label = tk.Label(main_frame, text="Este módulo permite gerar documentos DOCX a partir de evidências capturadas.",
                                bg='#f5f5f5', fg='#2c3e50', font=("Arial", 10))
        desc_label.pack(pady=10)
        
        def iniciar():
            if self.mostrar_janela_configuracao():
                # O processamento continua automaticamente
                pass
        
        self._create_styled_button(main_frame, text="Iniciar Gerador", 
                                  command=iniciar, style_type="accent", width=20).pack(pady=15)
        
        self._create_styled_button(main_frame, text="Voltar ao Menu Principal", 
                                  command=self.hide, style_type="glass", width=20).pack(pady=5)

    def hide(self):
        """Esconde a interface do módulo de forma segura"""
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
                
                # Voltar o foco para a janela principal
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

    def finalizar(self):
        """Método para finalizar gravação (para compatibilidade)"""
        self.gravando = False

    # ---------- Nova janela de configuração ----------
    def mostrar_janela_configuracao(self):
        """CORREÇÃO: Janela de configuração agora retorna True corretamente"""
        config_window = tk.Toplevel(self.root)
        config_window.title("Configuração de Arquivo")
        config_window.geometry("600x500")
        config_window.resizable(False, False)
        
        # Aplicar estilos
        self._apply_styles(config_window)
        
        config_window.transient(self.root)
        config_window.grab_set()
        
        main_frame = self._create_styled_frame(config_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self._create_styled_label(main_frame, text="PrintF - Configuração de Arquivo", 
                                 style_type="title").pack(pady=10)
        
        # Seleção de template
        if self.using_liquid_glass:
            ttk.Label(main_frame, text="Selecione o template DOCX:", 
                     style="Glass.TLabel").pack(anchor="w", pady=(10, 5))
        else:
            tk.Label(main_frame, text="Selecione o template DOCX:", 
                    bg='#f5f5f5', fg='#2c3e50', font=("Arial", 10)).pack(anchor="w", pady=(10, 5))
        
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
                                  command=selecionar_template, style_type="glass").pack(side=tk.RIGHT)
        
        # Seleção de diretório de evidências
        if self.using_liquid_glass:
            ttk.Label(main_frame, text="Selecione o diretório onde estão as evidências:", 
                     style="Glass.TLabel").pack(anchor="w", pady=(10, 5))
        else:
            tk.Label(main_frame, text="Selecione o diretório onde estão as evidências:", 
                    bg='#f5f5f5', fg='#2c3e50', font=("Arial", 10)).pack(anchor="w", pady=(10, 5))
        
        dir_frame = self._create_styled_frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        self.dir_var = tk.StringVar()
        dir_entry = self._create_styled_entry(dir_frame, textvariable=self.dir_var, width=40)
        dir_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        def selecionar_diretorio():
            dir_path = filedialog.askdirectory(title="Selecione o diretório onde estão as evidências")
            if dir_path:
                self.dir_var.set(dir_path)
                atualizar_lista_arquivos(dir_path)
        
        self._create_styled_button(dir_frame, text="Procurar", 
                                  command=selecionar_diretorio, style_type="glass").pack(side=tk.RIGHT)
        
        # Frame para exibir a lista de arquivos
        file_list_frame = self._create_styled_frame(main_frame)
        file_list_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        
        if self.using_liquid_glass:
            file_list_scrollbar = ttk.Scrollbar(file_list_frame, style="Glass.Vertical.TScrollbar")
        else:
            file_list_scrollbar = tk.Scrollbar(file_list_frame)
        file_list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox para mostrar arquivos encontrados
        self.file_listbox = self._create_styled_listbox(file_list_frame, 
                                                       yscrollcommand=file_list_scrollbar.set, 
                                                       height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        file_list_scrollbar.config(command=self.file_listbox.yview)
        
        if self.using_liquid_glass:
            self.file_count_label = ttk.Label(main_frame, text="Nenhum arquivo de imagem encontrado",
                                             style="Glass.TLabel")
        else:
            self.file_count_label = tk.Label(main_frame, text="Nenhum arquivo de imagem encontrado",
                                           bg='#f5f5f5', fg='#2c3e50', font=("Arial", 9))
        self.file_count_label.pack(anchor="w", pady=(0, 10))
        
        def atualizar_lista_arquivos(dir_path):
            """CORREÇÃO: Atualiza lista com múltiplos formatos de imagem"""
            self.file_listbox.delete(0, tk.END)
            image_files = self.carregar_evidencias(dir_path)
            
            for file_path in image_files:
                filename = os.path.basename(file_path)
                # Mostra também o timestamp para referência
                timestamp = datetime.fromtimestamp(os.path.getmtime(file_path))
                ext = os.path.splitext(filename)[1].upper()
                self.file_listbox.insert(tk.END, f"{filename} ({timestamp.strftime('%H:%M:%S')}) [{ext}]")
            
            if image_files:
                self.file_count_label.config(text=f"{len(image_files)} arquivo(s) de imagem encontrado(s)")
            else:
                self.file_count_label.config(text="Nenhum arquivo de imagem encontrado")
        
        # CORREÇÃO: Variável para controlar se deve processar
        processar = [False]  # Usa lista para permitir modificação dentro de função aninhada
        
        # Botões
        btn_frame = self._create_styled_frame(main_frame)
        btn_frame.pack(pady=20)
        
        def iniciar_geracao():
            """CORREÇÃO: Valida e inicia a geração"""
            if not self.template_var.get() or not self.dir_var.get():
                messagebox.showerror("Erro", "Por favor, selecione o template e o diretório de evidências.")
                return
            
            if not os.path.exists(self.template_var.get()):
                messagebox.showerror("Erro", "O arquivo de template selecionado não existe.")
                return
            
            if not os.path.exists(self.dir_var.get()):
                messagebox.showerror("Erro", "O diretório de evidências selecionado não existe.")
                return
            
            image_files = self.carregar_evidencias(self.dir_var.get())
            if not image_files:
                messagebox.showerror("Erro", "Nenhuma evidência de imagem encontrada no diretório selecionado.\n\n" +
                                   f"Formatos suportados: {', '.join(self.supported_extensions)}")
                return
            
            self.template_path = self.template_var.get()
            self.output_dir = self.dir_var.get()
            self.evidence_dir = self.dir_var.get()
            self.prints = image_files
            self.current_index = 0
            
            processar[0] = True
            config_window.destroy()
        
        self._create_styled_button(btn_frame, text="Gerar Documento", 
                                  command=iniciar_geracao, style_type="accent").pack(side=tk.LEFT, padx=5)
        self._create_styled_button(btn_frame, text="Cancelar", 
                                  command=config_window.destroy, style_type="glass").pack(side=tk.LEFT, padx=5)
        
        self.root.wait_window(config_window)
        
        # CORREÇÃO: Inicia processamento se configuração foi completada
        if processar[0]:
            self.iniciar_processamento()
            return True
        
        return False

    def iniciar_processamento(self):
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
        # Processa as evidências usando índice em vez de loop for
        documento_salvo = False  # Flag para controlar se o documento já foi salvo
        
        while self.current_index < len(self.prints):
            caminho_print = self.prints[self.current_index]
            
            # Verifica se o arquivo ainda existe
            if not os.path.exists(caminho_print):
                # Recarrega a lista se o arquivo não existir mais
                if not self.recarregar_evidencias():
                    break
                if self.current_index >= len(self.prints):
                    break
                caminho_print = self.prints[self.current_index]
            
            resultado = self.mostrar_imagem(caminho_print)
            
            if resultado is False:  # Processamento cancelado
                break
            elif resultado is None:  # Exclusão ocorreu, não incrementa índice
                # Recarrega a lista após exclusão
                self.recarregar_evidencias()
                continue
            elif resultado == "ja_salvou":  # Já salvou via "Incluir Todos"
                documento_salvo = True
                break  # Sai do loop completamente
            else:  # Adicionou com sucesso, vai para próxima
                self.current_index += 1
        
        # Só salva se não salvou anteriormente (no incluir_todos)
        if not documento_salvo:
            self.salvar_docx()

    def mostrar_imagem(self, caminho_print):
        popup = tk.Toplevel(self.root)
        popup.title("Adicionar Comentário à Evidência")
        popup.geometry("950x750")
        popup.resizable(False, False)

        # Aplicar estilos
        self._apply_styles(popup)

        self.processamento_cancelado = False
        resultado = None

        # Verifica se o arquivo ainda existe
        if not os.path.exists(caminho_print):
            messagebox.showerror("Erro", f"Arquivo não encontrado: {os.path.basename(caminho_print)}")
            popup.destroy()
            return None

        # CORREÇÃO: Abre imagem com tratamento de erro
        try:
            img = Image.open(caminho_print)
            img.thumbnail((850, 550))
            img_tk = ImageTk.PhotoImage(img)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar imagem: {str(e)}")
            popup.destroy()
            return None
        
        # Frame para imagem
        img_frame = self._create_styled_frame(popup)
        img_frame.pack(pady=10)
        
        # Para a imagem, manter fundo branco para melhor contraste
        label_img = tk.Label(img_frame, image=img_tk, bg='white')
        label_img.image = img_tk
        label_img.pack()

        # Frame para comentário
        comment_frame = self._create_styled_frame(popup)
        comment_frame.pack(pady=5)
        
        if self.using_liquid_glass:
            ttk.Label(comment_frame, text="Comentário (opcional):", 
                     style="Glass.TLabel").pack()
        else:
            tk.Label(comment_frame, text="Comentário (opcional):", 
                    bg='#f5f5f5', fg='#2c3e50', font=("Arial", 10)).pack()
        
        entry = self._create_styled_entry(comment_frame, width=80)
        entry.pack(pady=5)

        # Mostra informações do arquivo
        info_frame = self._create_styled_frame(popup)
        info_frame.pack(pady=5)
        
        file_info = f"Arquivo: {os.path.basename(caminho_print)}"
        timestamp = datetime.fromtimestamp(os.path.getmtime(caminho_print))
        file_info += f" - {timestamp.strftime('%H:%M:%S')}"
        
        if self.using_liquid_glass:
            ttk.Label(info_frame, text=file_info, font=("Arial", 10),
                     style="Glass.TLabel").pack()
        else:
            tk.Label(info_frame, text=file_info, font=("Arial", 10),
                    bg='#f5f5f5', fg='#2c3e50').pack()

        def editar_print():
            self.abrir_editor(caminho_print, popup)

        def adicionar():
            nonlocal resultado
            comentario = entry.get()
            
            # Atualiza comentário nos metadados
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
                # Adicionar a evidência atual primeiro
                comentario = entry.get()
                
                # Atualiza comentário nos metadados
                nome_arquivo = os.path.basename(caminho_print)
                for evidencia in self.metadata["evidencias"]:
                    if evidencia["arquivo"] == nome_arquivo:
                        evidencia["comentario"] = comentario
                        break
                self._salvar_metadata()
                
                self.doc.add_picture(caminho_print, width=Inches(5))
                if comentario.strip():
                    self.doc.add_paragraph(comentario)
                
                # Adicionar todas as evidências restantes
                for i in range(self.current_index + 1, len(self.prints)):
                    print_path = self.prints[i]
                    if os.path.exists(print_path):  # Verifica se o arquivo existe
                        # Usa comentário salvo nos metadados
                        nome_arquivo_restante = os.path.basename(print_path)
                        comentario_restante = self.obter_comentario(nome_arquivo_restante)
                        
                        self.doc.add_picture(print_path, width=Inches(5))
                        if comentario_restante.strip():
                            self.doc.add_paragraph(comentario_restante)
                        else:
                            self.doc.add_paragraph("")
                
                # Atualiza o índice para o final
                self.current_index = len(self.prints)
             
                # Salva o documento imediatamente
                self.salvar_docx()
                
                # Fecha o popup após salvar
                popup.destroy()
                
                resultado = "ja_salvou"  # Retorna um valor especial para indicar que já salvou
            else:
                resultado = False

        def excluir_print():
            nonlocal resultado
            if messagebox.askyesno("Confirmar Exclusão", "Tem certeza que deseja excluir esta evidência?"):
                # Remove o arquivo
                try:
                    nome_arquivo = os.path.basename(caminho_print)
                    
                    # Marca como excluída nos metadados em vez de remover fisicamente
                    for evidencia in self.metadata["evidencias"]:
                        if evidencia["arquivo"] == nome_arquivo:
                            evidencia["excluida"] = True
                            break
                    
                    self._salvar_metadata()
                    
                    # Remove fisicamente o arquivo
                    os.remove(caminho_print)
                    print(f"Arquivo excluído: {caminho_print}")
                    
                    resultado = None  # Indica que houve exclusão
                    popup.destroy()
                    
                except Exception as e:
                    print(f"Erro ao excluir arquivo: {e}")
                    messagebox.showerror("Erro", f"Não foi possível excluir o arquivo: {e}")
                    resultado = False
                    popup.destroy()
                    return
            else:
                resultado = False  # Usuário cancelou a exclusão

        # Frame para botões de ação
        acoes_frame = self._create_styled_frame(popup)
        acoes_frame.pack(pady=10)

        self._create_styled_button(acoes_frame, text="✏ Editar Print", 
                                  command=editar_print, style_type="glass", width=15).pack(side=tk.LEFT, padx=5)
        self._create_styled_button(acoes_frame, text="Adicionar e Próximo", 
                                  command=adicionar, style_type="accent", width=15).pack(side=tk.LEFT, padx=5)
        self._create_styled_button(acoes_frame, text="🗑️ Excluir Print", 
                                  command=excluir_print, style_type="glass", width=15).pack(side=tk.LEFT, padx=5)

        # Frame para botões de controle
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
        # CORREÇÃO: Verifica se template_path existe e não é None
        if self.template_path and os.path.exists(self.template_path):
            # Usa exatamente o mesmo nome do template
            nome_arquivo = os.path.basename(self.template_path)
        else:
            # Nome simples com timestamp se não houver template
            nome_arquivo = f"Evidencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        caminho_save = os.path.join(self.output_dir, nome_arquivo)
        
        # Se o arquivo já existir, adiciona um sufixo numérico
        if os.path.exists(caminho_save):
            nome_base = os.path.splitext(nome_arquivo)[0]
            extensao = os.path.splitext(nome_arquivo)[1]
            contador = 1
            while os.path.exists(caminho_save):
                nome_arquivo = f"{nome_base}_{contador}{extensao}"
                caminho_save = os.path.join(self.output_dir, nome_arquivo)
                contador += 1
        
        try:
            self.doc.save(caminho_save)
            self.saved_file_path = caminho_save
            
            # Função para abrir a pasta (será chamada após fechar o messagebox)
            def abrir_pasta_apos_mensagem():
                if os.name == 'nt':
                    os.startfile(self.output_dir)
                elif os.name == 'posix':
                    import subprocess
                    if sys.platform == 'darwin':
                        subprocess.Popen(['open', self.output_dir])
                    else:
                        subprocess.Popen(['xdg-open', self.output_dir])
            
            # Mostra a mensagem e agenda a abertura da pasta para depois
            messagebox.showinfo("Concluído", f"Documento gerado com sucesso!\nSalvo em:\n{caminho_save}")
            
            # Agenda a abertura da pasta para depois de fechar o messagebox
            self.root.after(100, abrir_pasta_apos_mensagem)
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar documento: {str(e)}")

    # ---------- Editor de prints ----------
    def abrir_editor(self, caminho_print, parent):
        editor = tk.Toplevel(parent)
        editor.title("Editor de Evidência")
        editor.geometry("1200x800")
        
        # Aplicar estilos
        self._apply_styles(editor)
        
        # Frame principal
        main_frame = self._create_styled_frame(editor)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Frame para ferramentas e opções
        tools_frame = self._create_styled_frame(main_frame)
        tools_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        
        # Frame para a área de desenho
        canvas_frame = self._create_styled_frame(main_frame)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Carrega a imagem original
        self.original_img = Image.open(caminho_print).convert("RGBA")
        img_w, img_h = self.original_img.size
        
        # Calcula o fator de escala para exibição
        max_w, max_h = 1000, 700
        scale = min(max_w / img_w, max_h / img_h)
        self.scale_factor = scale
        disp_w, disp_h = int(img_w * scale), int(img_h * scale)
        
        # Cria cópia da imagem para edição
        self.editing_img = self.original_img.copy()
        self.display_img = self.editing_img.resize((disp_w, disp_h), Image.LANCZOS)

        # Variáveis para controle
        self.current_tk_img = ImageTk.PhotoImage(self.display_img)
        self.elements = []  # Lista de elementos desenhados
        self.undo_stack = []  # PILHA PARA DESFAZER AÇÕES
        self.temp_element = None
        
        # Canvas para a imagem - manter fundo cinza para melhor contraste com imagens
        canvas_bg = 'gray'
        self.canvas = tk.Canvas(canvas_frame, width=disp_w, height=disp_h, cursor="cross", bg=canvas_bg)
        self.canvas.pack(padx=5, pady=5)
        self.canvas_img = self.canvas.create_image(0, 0, anchor="nw", image=self.current_tk_img)
        
        # Variáveis de controle - COR PADRÃO VERMELHA
        tool_var = tk.StringVar(value="rectangle")  # RETÂNGULO COMO PADRÃO
        color_var = tk.StringVar(value="#FF0000")   # VERMELHO COMO PADRÃO
        width_var = tk.IntVar(value=3)
        
        # Ferramentas
        if self.using_liquid_glass:
            ttk.Label(tools_frame, text="Ferramenta:", style="Glass.TLabel").pack(side=tk.LEFT, padx=5)
        else:
            tk.Label(tools_frame, text="Ferramenta:", bg='#f5f5f5', fg='#2c3e50').pack(side=tk.LEFT, padx=5)
        
        # Frame para os botões de ferramentas
        tools_buttons_frame = self._create_styled_frame(tools_frame)
        tools_buttons_frame.pack(side=tk.LEFT, padx=5)
        
        # Ferramentas disponíveis
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
        
        # Cores
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
        
        # Espessura
        if self.using_liquid_glass:
            ttk.Label(tools_frame, text="Espessura:", style="Glass.TLabel").pack(side=tk.LEFT, padx=20)
        else:
            tk.Label(tools_frame, text="Espessura:", bg='#f5f5f5', fg='#2c3e50').pack(side=tk.LEFT, padx=20)
        
        width_scale = tk.Scale(tools_frame, from_=1, to=10, variable=width_var, 
                              orient=tk.HORIZONTAL, length=100, showvalue=True)
        width_scale.pack(side=tk.LEFT, padx=5)
        
        # Botões de ação do editor
        action_frame = self._create_styled_frame(tools_frame)
        action_frame.pack(side=tk.RIGHT, padx=10)
        
        self._create_styled_button(action_frame, text="↶ Desfazer", 
                                  command=self.desfazer_acao, style_type="glass").pack(side=tk.LEFT, padx=2)
        self._create_styled_button(action_frame, text="Salvar", 
                                  command=lambda: self.salvar_edicao(caminho_print, editor), 
                                  style_type="accent").pack(side=tk.LEFT, padx=2)
        self._create_styled_button(action_frame, text="Cancelar", 
                                  command=editor.destroy, style_type="glass").pack(side=tk.LEFT, padx=2)
        
        # Variáveis para controle de desenho
        self.start_x = None
        self.start_y = None
        self.current_element = None
        
        # Bind eventos do canvas
        self.canvas.bind("<Button-1>", lambda e: self.iniciar_desenho(e, tool_var.get()))
        self.canvas.bind("<B1-Motion>", lambda e: self.desenhar(e, tool_var.get()))
        self.canvas.bind("<ButtonRelease-1>", lambda e: self.finalizar_desenho(e, tool_var.get(), color_var.get(), width_var.get()))
        
        # Centralizar
        editor.transient(parent)
        editor.grab_set()

    def iniciar_desenho(self, event, tool):
        self.start_x = event.x
        self.start_y = event.y
        
        if tool == "text":
            # Para texto, pede o texto via dialog
            texto = simpledialog.askstring("Texto", "Digite o texto:", parent=self.root)
            if texto:
                # Converte coordenadas para escala original
                orig_x = int(event.x / self.scale_factor)
                orig_y = int(event.y / self.scale_factor)
                
                element_data = {
                    "type": "text",
                    "text": texto,
                    "x": orig_x,
                    "y": orig_y,
                    "color": "#FF0000",  # Vermelho padrão
                    "size": 20
                }
                self.elements.append(element_data)
                self.aplicar_elemento_na_imagem(element_data)
                self.atualizar_canvas()
        else:
            # Para outras ferramentas, inicia desenho temporário
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
            # Salva o elemento
            coords = self.canvas.coords(self.current_element)
            # Converte coordenadas para escala original
            orig_coords = [int(coord / self.scale_factor) for coord in coords]
            
            element_data = {
                "type": tool,
                "coords": orig_coords,
                "color": color,
                "width": width
            }
            self.elements.append(element_data)
            self.undo_stack.append(element_data.copy())
            
            # Aplica na imagem
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
            
            # Adiciona ponta da seta
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
            
            # Recria a imagem
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
    
    # Aplicar estilos básicos mesmo no modo independente
    try:
        root.configure(bg='#f5f5f5')
    except:
        pass
    
    # Centraliza a janela
    root.eval('tk::PlaceWindow . center')
    
    # Frame principal
    main_frame = tk.Frame(root, bg='#f5f5f5', padx=30, pady=30)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Título
    title_label = tk.Label(main_frame, text="PrintF - Gerador de Evidências", 
                         font=("Arial", 16, "bold"), bg='#f5f5f5', fg='#2c3e50')
    title_label.pack(pady=20)
    
    # Botão para iniciar
    def iniciar_gerador():
        gerador = EvidenceGeneratorModule(root)
        gerador.show()
    
    start_btn = tk.Button(main_frame, text="Iniciar Gerador de Evidências", 
                         command=iniciar_gerador, width=25,
                         bg='#3498db', fg='white', font=("Arial", 12, "bold"),
                         relief="flat", cursor="hand2")
    start_btn.pack(pady=10)
    
    # Efeitos hover
    start_btn.bind("<Enter>", lambda e: start_btn.config(bg='#2980b9'))
    start_btn.bind("<Leave>", lambda e: start_btn.config(bg='#3498db'))
    
    # Botão para sair
    exit_btn = tk.Button(main_frame, text="Sair", command=root.quit, width=15,
                        bg='#e74c3c', fg='white', font=("Arial", 10),
                        relief="flat", cursor="hand2")
    exit_btn.pack(pady=10)
    
    exit_btn.bind("<Enter>", lambda e: exit_btn.config(bg='#c0392b'))
    exit_btn.bind("<Leave>", lambda e: exit_btn.config(bg='#e74c3c'))
    
    root.mainloop()