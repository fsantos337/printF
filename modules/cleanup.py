import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime

class CleanupModule:
    def __init__(self, parent, settings):
        self.parent = parent
        self.settings = settings
        self.root = None
        self.pasta_selecionada = ""
        self.arquivos_selecionados = set()
        
        # Detectar se estamos usando Liquid Glass
        self.using_liquid_glass = False
        self._detect_theme()
        
    def _detect_theme(self):
        """Detecta se o tema Liquid Glass está ativo"""
        try:
            # Verificar se o estilo está disponível e ativo
            from modules.styles import LiquidGlassStyle
            # Verificar configurações do usuário
            if self.settings.get('theme') == 'liquid_glass':
                self.using_liquid_glass = True
                self.style_manager = LiquidGlassStyle
        except ImportError:
            # Fallback para estilo padrão
            self.using_liquid_glass = False
            self.style_manager = None
        except Exception as e:
            self.using_liquid_glass = False
            self.style_manager = None
    
    def _apply_theme_to_widgets(self):
        """Aplica o tema Liquid Glass aos widgets se estiver ativo"""
        if not self.using_liquid_glass or not self.style_manager:
            return
            
        try:
            # Aplicar estilo à janela principal
            self.style_manager.apply_window_style(self.root)
            
            # Configurar cores de fundo para frames
            self._configure_widget_colors()
            
        except Exception as e:
            pass
    
    def _configure_widget_colors(self):
        """Configura cores dos widgets para o tema Liquid Glass"""
        if not self.using_liquid_glass:
            return
            
        try:
            # Configurar cores de fundo
            bg_color = self.style_manager.BG_PRIMARY
            card_bg = self.style_manager.BG_CARD
            text_color = self.style_manager.TEXT_PRIMARY
            secondary_text = self.style_manager.TEXT_SECONDARY
            
            # Aplicar às frames principais (com verificação de existência)
            widgets_to_configure = [
                'main_frame', 'frame_superior', 'frame_selecao', 
                'frame_controles', 'frame_botoes_selecao', 'frame_lista',
                'frame_acao', 'frame_info', 'frame_botoes_acao'
            ]
            
            for widget_name in widgets_to_configure:
                if hasattr(self, widget_name):
                    widget = getattr(self, widget_name)
                    if hasattr(widget, 'configure') and widget.winfo_exists():
                        widget.configure(bg=bg_color)
            
            # Configurar LabelFrame
            if hasattr(self, 'frame_lista') and self.frame_lista.winfo_exists():
                self.frame_lista.configure(bg=bg_color, fg=text_color)
            
            # Configurar labels
            labels = [
                'titulo', 'label_info', 'label_docx', 
                'label_imagens', 'label_outros', 'label_selecionados'
            ]
            
            for label_name in labels:
                if hasattr(self, label_name):
                    label = getattr(self, label_name)
                    if hasattr(label, 'configure') and label.winfo_exists():
                        label.configure(bg=bg_color, fg=text_color)
            
            # Configurar botões padrão
            buttons = [
                'btn_selecionar', 'btn_sel_todos', 'btn_desel_todos',
                'btn_sel_imagens', 'btn_sel_docx', 'btn_voltar'
            ]
            
            for button_name in buttons:
                if hasattr(self, button_name):
                    button = getattr(self, button_name)
                    if hasattr(button, 'configure') and button.winfo_exists():
                        button.configure(
                            bg=self.style_manager.BG_CARD,
                            fg=text_color,
                            relief="flat",
                            borderwidth=1
                        )
            
            # Configurar botão de excluir selecionados
            if hasattr(self, 'btn_excluir_selecionados') and self.btn_excluir_selecionados.winfo_exists():
                self.btn_excluir_selecionados.configure(
                    bg=self.style_manager.ACCENT_WARNING,
                    fg=text_color,
                    relief="flat"
                )
            
            # Configurar entry
            if hasattr(self, 'entry_pasta') and self.entry_pasta.winfo_exists():
                self.entry_pasta.configure(
                    bg=self.style_manager.BG_SECONDARY,
                    fg=text_color,
                    insertbackground=text_color,
                    relief="flat"
                )
        
        except Exception as e:
            pass
    
    def _create_styled_button(self, parent, text, command, style_type="glass", **kwargs):
        """Cria botão com estilo apropriado baseado no tema"""
        if self.using_liquid_glass and self.style_manager:
            if style_type == "accent":
                return self.style_manager.create_accent_button(parent, text, command, **kwargs)
            elif style_type == "error":
                btn = self.style_manager.create_glass_button(parent, text, command, **kwargs)
                btn.configure(style="Error.TButton")
                return btn
            else:
                return self.style_manager.create_glass_button(parent, text, command, **kwargs)
        else:
            # Fallback para estilo padrão
            bg_color = "#f0f0f0"
            fg_color = "#000000"
            
            if style_type == "accent":
                bg_color = "#3498db"
                fg_color = "white"
            elif style_type == "error":
                bg_color = "#e74c3c"
                fg_color = "white"
            
            return tk.Button(
                parent, 
                text=text, 
                command=command,
                bg=bg_color,
                fg=fg_color,
                relief="raised",
                **kwargs
            )
    
    def _create_styled_frame(self, parent, **kwargs):
        """Cria frame com estilo apropriado"""
        if self.using_liquid_glass and self.style_manager:
            return self.style_manager.create_glass_frame(parent, **kwargs)
        else:
            return tk.Frame(parent, **kwargs)
    
    def _create_styled_label(self, parent, text, style_type="glass", **kwargs):
        """Cria label com estilo apropriado"""
        if self.using_liquid_glass and self.style_manager:
            if style_type == "title":
                return self.style_manager.create_title_label(parent, text, **kwargs)
            else:
                return ttk.Label(parent, text=text, style="Glass.TLabel", **kwargs)
        else:
            return tk.Label(parent, text=text, **kwargs)
    
    def _create_styled_entry(self, parent, **kwargs):
        """Cria entry com estilo apropriado"""
        if self.using_liquid_glass and self.style_manager:
            return self.style_manager.create_glass_entry(parent, **kwargs)
        else:
            return tk.Entry(parent, **kwargs)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta para limpar arquivos")
        if pasta:
            self.pasta_selecionada = pasta
            self.entry_pasta.delete(0, tk.END)
            self.entry_pasta.insert(0, pasta)
            self.analisar_pasta()
            
    def analisar_pasta(self):
        if not self.pasta_selecionada:
            return
            
        try:
            arquivos = os.listdir(self.pasta_selecionada)
            
            # Limpar seleções anteriores
            self.arquivos_selecionados.clear()
            
            # Limpar a treeview anterior
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Contar arquivos por tipo
            total_arquivos = len(arquivos)
            self.arquivos_docx = []
            self.arquivos_imagens = []
            self.outros_arquivos = []
            
            # Adicionar arquivos à treeview
            for arquivo in arquivos:
                caminho_completo = os.path.join(self.pasta_selecionada, arquivo)
                if os.path.isfile(caminho_completo):
                    extensao = os.path.splitext(arquivo)[1].lower()
                    tamanho = os.path.getsize(caminho_completo)
                    
                    # Formatar tamanho
                    if tamanho < 1024:
                        tamanho_str = f"{tamanho} B"
                    elif tamanho < 1024 * 1024:
                        tamanho_str = f"{tamanho / 1024:.1f} KB"
                    else:
                        tamanho_str = f"{tamanho / (1024 * 1024):.1f} MB"
                    
                    # Determinar tipo
                    tipo = "Outro"
                    cor = "gray"
                    if extensao == '.docx':
                        tipo = "DOCX"
                        cor = "blue"
                        self.arquivos_docx.append(arquivo)
                    elif extensao in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff']:
                        tipo = "Imagem"
                        cor = "green"
                        self.arquivos_imagens.append(arquivo)
                    else:
                        tipo = "Outro"
                        cor = "gray"
                        self.outros_arquivos.append(arquivo)
                    
                    # Adicionar à lista
                    item = self.tree.insert("", "end", values=(arquivo, tipo, tamanho_str), tags=(cor,))
                    self.tree.tag_configure("blue", foreground="blue")
                    self.tree.tag_configure("green", foreground="green")
                    self.tree.tag_configure("gray", foreground="gray")
            
            # Atualizar informações na interface
            self.label_info.config(text=f"Total de arquivos: {total_arquivos}")
            self.label_docx.config(text=f"DOCX: {len(self.arquivos_docx)}")
            self.label_imagens.config(text=f"Imagens: {len(self.arquivos_imagens)}")
            self.label_outros.config(text=f"Outros: {len(self.outros_arquivos)}")
            self.label_selecionados.config(text=f"Selecionados: 0")
            
            # Atualizar botões
            self.atualizar_botoes()
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao analisar pasta: {str(e)}")
    
    def on_item_select(self, event):
        """Atualiza a contagem de arquivos selecionados"""
        selecionados = self.tree.selection()
        self.arquivos_selecionados = set(selecionados)
        self.label_selecionados.config(text=f"Selecionados: {len(self.arquivos_selecionados)}")
        self.atualizar_botoes()
    
    def selecionar_todos(self):
        """Seleciona todos os arquivos da lista"""
        todos_itens = self.tree.get_children()
        self.tree.selection_set(todos_itens)
        self.arquivos_selecionados = set(todos_itens)
        self.label_selecionados.config(text=f"Selecionados: {len(self.arquivos_selecionados)}")
        self.atualizar_botoes()
    
    def desselecionar_todos(self):
        """Desseleciona todos os arquivos"""
        self.tree.selection_remove(self.tree.selection())
        self.arquivos_selecionados.clear()
        self.label_selecionados.config(text=f"Selecionados: 0")
        self.atualizar_botoes()
    
    def selecionar_por_tipo(self, tipo):
        """Seleciona arquivos por tipo"""
        todos_itens = self.tree.get_children()
        itens_selecionar = []
        
        for item in todos_itens:
            valores = self.tree.item(item, "values")
            if valores[1] == tipo:  # valores[1] é o tipo
                itens_selecionar.append(item)
        
        self.tree.selection_set(itens_selecionar)
        self.arquivos_selecionados = set(itens_selecionar)
        self.label_selecionados.config(text=f"Selecionados: {len(self.arquivos_selecionados)}")
        self.atualizar_botoes()
    
    def atualizar_botoes(self):
        """Atualiza o estado dos botões baseado na seleção"""
        # Botão de excluir selecionados
        if len(self.arquivos_selecionados) > 0:
            self.btn_excluir_selecionados.config(state=tk.NORMAL)
        else:
            self.btn_excluir_selecionados.config(state=tk.DISABLED)
        
        # Botões de seleção por tipo
        self.btn_sel_imagens.config(state=tk.NORMAL if len(self.arquivos_imagens) > 0 else tk.DISABLED)
        self.btn_sel_docx.config(state=tk.NORMAL if len(self.arquivos_docx) > 0 else tk.DISABLED)
        
        # Botões de seleção geral
        total_itens = len(self.tree.get_children())
        self.btn_sel_todos.config(state=tk.NORMAL if total_itens > 0 else tk.DISABLED)
        self.btn_desel_todos.config(state=tk.NORMAL if len(self.arquivos_selecionados) > 0 else tk.DISABLED)
    
    def excluir_selecionados(self):
        """Exclui apenas os arquivos selecionados"""
        if not self.arquivos_selecionados:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para exclusão.")
            return
        
        # Confirmar ação
        confirmacao = messagebox.askyesno(
            "Confirmação", 
            f"⚠️ ATENÇÃO: Esta ação é irreversível!\n\n"
            f"Serão excluídos {len(self.arquivos_selecionados)} arquivo(s) selecionado(s).\n\n"
            "Deseja continuar?"
        )
        
        if not confirmacao:
            return
            
        try:
            arquivos_apagados = 0
            tamanho_liberado = 0
            
            for item_id in self.arquivos_selecionados:
                valores = self.tree.item(item_id, "values")
                nome_arquivo = valores[0]  # Nome do arquivo
                caminho_completo = os.path.join(self.pasta_selecionada, nome_arquivo)
                
                if os.path.isfile(caminho_completo):
                    try:
                        tamanho_arquivo = os.path.getsize(caminho_completo)
                        os.remove(caminho_completo)
                        arquivos_apagados += 1
                        tamanho_liberado += tamanho_arquivo
                    except Exception as e:
                        print(f"Erro ao apagar {nome_arquivo}: {str(e)}")
            
            # Converter tamanho para formato legível
            if tamanho_liberado < 1024:
                tamanho_texto = f"{tamanho_liberado} bytes"
            elif tamanho_liberado < 1024 * 1024:
                tamanho_texto = f"{tamanho_liberado / 1024:.2f} KB"
            else:
                tamanho_texto = f"{tamanho_liberado / (1024 * 1024):.2f} MB"
            
            messagebox.showinfo("Exclusão Concluída", 
                               f"Foram removidos {arquivos_apagados} arquivo(s).\n"
                               f"Espaço liberado: {tamanho_texto}")
            
            # Atualizar análise
            self.analisar_pasta()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante a exclusão: {str(e)}")
    
    def excluir_por_tipo(self):
        """Exclui arquivos por tipo (modo antigo)"""
        excluir_png = self.var_png.get()
        excluir_docx = self.var_docx.get()
        
        if not excluir_png and not excluir_docx:
            messagebox.showwarning("Aviso", "Selecione pelo menos um tipo de arquivo para excluir.")
            return
            
        # Confirmar ação
        confirmacao = messagebox.askyesno(
            "Confirmação", 
            "⚠️ ATENÇÃO: Esta ação é irreversível!\n\n" +
            f"Serão excluídos permanentemente:\n" +
            f"• Arquivos PNG/JPG: {excluir_png and 'SIM' or 'NÃO'}\n" +
            f"• Arquivos DOCX: {excluir_docx and 'SIM' or 'NÃO'}\n\n" +
            "Deseja continuar?"
        )
        
        if not confirmacao:
            return
            
        try:
            arquivos = os.listdir(self.pasta_selecionada)
            arquivos_apagados = 0
            tamanho_liberado = 0
            tipos_apagados = []
            
            for arquivo in arquivos:
                caminho_completo = os.path.join(self.pasta_selecionada, arquivo)
                if os.path.isfile(caminho_completo):
                    extensao = os.path.splitext(arquivo)[1].lower()
                    
                    # Verificar se deve excluir baseado nas seleções
                    deve_excluir = False
                    
                    if excluir_png and extensao in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff']:
                        deve_excluir = True
                        if 'Imagens' not in tipos_apagados:
                            tipos_apagados.append('Imagens')
                            
                    if excluir_docx and extensao == '.docx':
                        deve_excluir = True
                        if 'Documentos DOCX' not in tipos_apagados:
                            tipos_apagados.append('Documentos DOCX')
                    
                    if deve_excluir:
                        try:
                            tamanho_arquivo = os.path.getsize(caminho_completo)
                            os.remove(caminho_completo)
                            arquivos_apagados += 1
                            tamanho_liberado += tamanho_arquivo
                        except Exception as e:
                            print(f"Erro ao apagar {arquivo}: {str(e)}")
            
            # Mensagem de sucesso
            if arquivos_apagados > 0:
                messagebox.showinfo("Limpeza Concluída", 
                                   f"Foram removidos {arquivos_apagados} arquivos.\n"
                                   f"Tipos removidos: {', '.join(tipos_apagados)}")
            else:
                messagebox.showinfo("Limpeza Concluída", 
                                   "Nenhum arquivo foi removido.")
            
            # Atualizar análise
            self.analisar_pasta()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante a limpeza: {str(e)}")

    def show(self):
        """Mostra a interface do módulo"""
        if not self.root:
            self._create_interface()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_set()
        
        # Configurar protocolo de fechamento para restaurar janela principal
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_window)

    def _on_close_window(self):
        """Manipula o fechamento da janela do módulo"""
        self.hide()

    def _create_interface(self):
        """Cria a interface do módulo"""
        self.root = tk.Toplevel(self.parent)
        self.root.title("PrintF - Limpeza de Arquivos")
        self.root.geometry("900x800")
        
        # Centralizar na tela principal
        self.root.transient(self.parent)
        self.root.grab_set()
        
        # Variáveis para checkboxes
        self.var_png = tk.BooleanVar(value=True)
        self.var_docx = tk.BooleanVar(value=False)
        
        # Frame principal
        self.main_frame = self._create_styled_frame(self.root, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        self.titulo = self._create_styled_label(self.main_frame, 
                                              text="🗑️ PrintF - Limpar Arquivos", 
                                              style_type="title")
        self.titulo.pack(pady=10)
        
        # Seleção de pasta
        self.frame_superior = self._create_styled_frame(self.main_frame)
        self.frame_superior.pack(fill=tk.X, pady=5)
        
        tk.Label(self.frame_superior, text="Pasta:").pack(anchor="w")
        
        self.frame_selecao = self._create_styled_frame(self.frame_superior)
        self.frame_selecao.pack(fill=tk.X, pady=5)
        
        # Campo de entrada maior e mais próximo do botão
        self.entry_pasta = self._create_styled_entry(self.frame_selecao, width=70)
        self.entry_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.btn_selecionar = self._create_styled_button(self.frame_selecao, 
                                                       text="Procurar", 
                                                       command=self.selecionar_pasta)
        self.btn_selecionar.pack(side=tk.RIGHT)
        
        # Controles de seleção
        self.frame_controles = self._create_styled_frame(self.main_frame)
        self.frame_controles.pack(fill=tk.X, pady=10)
        
        # Botões de seleção
        tk.Label(self.frame_controles, text="Seleção:", font=("Arial", 10)).pack(anchor="w")
        
        self.frame_botoes_selecao = self._create_styled_frame(self.frame_controles)
        self.frame_botoes_selecao.pack(fill=tk.X, pady=5)
        
        self.btn_sel_todos = self._create_styled_button(self.frame_botoes_selecao, 
                                                      text="✓ Selecionar Todos", 
                                                      command=self.selecionar_todos, 
                                                      state=tk.DISABLED)
        self.btn_sel_todos.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_desel_todos = self._create_styled_button(self.frame_botoes_selecao, 
                                                        text="✗ Desmarcar Todos", 
                                                        command=self.desselecionar_todos, 
                                                        state=tk.DISABLED)
        self.btn_desel_todos.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_sel_imagens = self._create_styled_button(self.frame_botoes_selecao, 
                                                        text="🖼️ Selecionar Imagens", 
                                                        command=lambda: self.selecionar_por_tipo("Imagem"), 
                                                        state=tk.DISABLED)
        self.btn_sel_imagens.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_sel_docx = self._create_styled_button(self.frame_botoes_selecao, 
                                                     text="📄 Selecionar DOCX", 
                                                     command=lambda: self.selecionar_por_tipo("DOCX"), 
                                                     state=tk.DISABLED)
        self.btn_sel_docx.pack(side=tk.LEFT)
        
        # Lista de arquivos
        self.frame_lista = tk.LabelFrame(self.main_frame, text="Arquivos na Pasta", padx=10, pady=10)
        self.frame_lista.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Treeview para listar arquivos
        columns = ("nome", "tipo", "tamanho")
        
        # Usar estilo ttk se Liquid Glass estiver ativo
        if self.using_liquid_glass:
            self.tree = ttk.Treeview(self.frame_lista, columns=columns, show="headings", height=12, selectmode="extended")
        else:
            self.tree = ttk.Treeview(self.frame_lista, columns=columns, show="headings", height=12, selectmode="extended")
        
        # Configurar colunas
        self.tree.heading("nome", text="Nome do Arquivo")
        self.tree.heading("tipo", text="Tipo")
        self.tree.heading("tamanho", text="Tamanho")
        
        self.tree.column("nome", width=500)
        self.tree.column("tipo", width=100)
        self.tree.column("tamanho", width=100)
        
        # Scrollbar
        if self.using_liquid_glass:
            scrollbar = self.style_manager.create_scrollbar(self.frame_lista, orient=tk.VERTICAL)
        else:
            scrollbar = ttk.Scrollbar(self.frame_lista, orient=tk.VERTICAL, command=self.tree.yview)
        
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)
        
        # Informações e botões de ação
        self.frame_acao = self._create_styled_frame(self.main_frame)
        self.frame_acao.pack(fill=tk.X, pady=10)
        
        # Informações
        self.frame_info = self._create_styled_frame(self.frame_acao)
        self.frame_info.pack(side=tk.LEFT, fill=tk.Y)
        
        self.label_info = self._create_styled_label(self.frame_info, text="Total: 0", font=("Arial", 9))
        self.label_info.pack(anchor="w")
        
        self.label_docx = self._create_styled_label(self.frame_info, text="DOCX: 0", font=("Arial", 9))
        self.label_docx.pack(anchor="w")
        
        self.label_imagens = self._create_styled_label(self.frame_info, text="Imagens: 0", font=("Arial", 9))
        self.label_imagens.pack(anchor="w")
        
        self.label_outros = self._create_styled_label(self.frame_info, text="Outros: 0", font=("Arial", 9))
        self.label_outros.pack(anchor="w")
        
        self.label_selecionados = self._create_styled_label(self.frame_info, text="Selecionados: 0", font=("Arial", 9, "bold"))
        self.label_selecionados.pack(anchor="w")
        
        # Botões de ação
        self.frame_botoes_acao = self._create_styled_frame(self.frame_acao)
        self.frame_botoes_acao.pack(side=tk.RIGHT)
        
        self.btn_excluir_selecionados = self._create_styled_button(self.frame_botoes_acao, 
                                                                text="🗑️ Excluir Selecionados", 
                                                                command=self.excluir_selecionados, 
                                                                state=tk.DISABLED,
                                                                style_type="error")
        self.btn_excluir_selecionados.pack(pady=5)
        
        # Botão voltar
        self.btn_voltar = self._create_styled_button(self.main_frame, 
                                                   text="Voltar ao Menu Principal", 
                                                   command=self.hide, 
                                                   width=20)
        self.btn_voltar.pack(pady=10)
        
        # Aplicar tema após criar todos os widgets
        self.root.after(100, self._apply_theme_to_widgets)

    def hide(self):
        """Esconde a interface do módulo"""
        if self.root:
            try:
                self.root.grab_release()
                self.root.withdraw()
            except:
                pass

# Função principal para teste independente
def main():
    root = tk.Tk()
    app = CleanupModule(root, {})
    app.show()
    root.mainloop()

if __name__ == "__main__":
    main()