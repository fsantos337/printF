import tkinter as tk
from tkinter import messagebox
import subprocess
import sys
import os
import threading
import time

# Controle de instância única
def check_single_instance():
    """Garante que apenas uma instância do aplicativo esteja rodando"""
    try:
        # Tenta criar um arquivo de lock
        lock_file = os.path.join(os.path.dirname(__file__), "printF.lock")
        
        if os.path.exists(lock_file):
            # Verifica se o processo ainda está ativo
            with open(lock_file, 'r') as f:
                pid = f.read().strip()
            
            # Para Windows, verifica si o processo existe
            try:
                import psutil
                if psutil.pid_exists(int(pid)):
                    return False
            except:
                # Se não tiver psutil, assumes que outra instância está rodando
                return False
        
        # Cria o arquivo de lock
        with open(lock_file, 'w') as f:
            f.write(str(os.getpid()))
        
        return True
        
    except Exception as e:
        print(f"Erro no controle de instância: {e}")
        return True

def cleanup_lock():
    """Remove o arquivo de lock ao fechar o aplicativo"""
    try:
        lock_file = os.path.join(os.path.dirname(__file__), "printF.lock")
        if os.path.exists(lock_file):
            os.remove(lock_file)
    except:
        pass

# Função para obter o caminho base correto
def get_base_path():
    """Retorna o caminho base correto dependendo se estamos em desenvolvimento ou no executável"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_exe_path(exe_name):
    """Retorna o caminho completo para um executável"""
    base_path = get_base_path()
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    possible_paths = [
        os.path.join(current_dir, exe_name),
        os.path.join(base_path, exe_name),
        os.path.join(current_dir, "dist", exe_name),
        os.path.join(os.getcwd(), exe_name),
        exe_name
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return os.path.abspath(path)
    
    return exe_name

# Dicionário para controlar processos ativos
processos_ativos = {}

def executar_com_timeout(exe_path, nome_processo, timeout=10):
    """Executa um processo com timeout"""
    try:
        if exe_path in processos_ativos and processos_ativos[exe_path].poll() is None:
            messagebox.showwarning("Aviso", f"{nome_processo} já está em execução!")
            return
        
        # Minimiza a janela principal antes de executar
        root.iconify()
        
        processo = subprocess.Popen([exe_path])
        processos_ativos[exe_path] = processo
        
        # Thread para verificar se o processo iniciou corretamente
        def verificar_processo():
            time.sleep(2)  # Aguarda 2 segundos
            if processo.poll() is not None:  # Processo já terminou
                messagebox.showerror("Erro", f"{nome_processo} fechou inesperadamente!")
                processos_ativos.pop(exe_path, None)
                # Restaura a janela principal se o processo fechar
                root.deiconify()
        
        threading.Thread(target=verificar_processo, daemon=True).start()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao executar {nome_processo}: {str(e)}")
        # Restaura a janela principal em caso de erro
        root.deiconify()

def gerar_templates():
    try:
        exe_path = get_exe_path("gerador_templates.exe")
        if os.path.exists(exe_path):
            executar_com_timeout(exe_path, "Gerador de Templates")
        else:
            messagebox.showerror("Erro", f"Executável não encontrado!\n{exe_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir Gerador de Templates: {str(e)}")

def capturar_evidencias():
    try:
        exe_path = get_exe_path("gravador_evidencias.exe")
        if os.path.exists(exe_path):
            executar_com_timeout(exe_path, "Gravador de Evidências")
        else:
            messagebox.showerror("Erro", f"Executável não encontrado!\n{exe_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir Gravador de Evidências: {str(e)}")

def limpar_arquivos():
    try:
        exe_path = get_exe_path("limpador_arquivos.exe")
        if os.path.exists(exe_path):
            executar_com_timeout(exe_path, "Limpador de Arquivos")
        else:
            messagebox.showerror("Erro", f"Executável não encontrado!\n{exe_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir Limpador de Arquivos: {str(e)}")

def fechar():
    # Para todos os processos ativos antes de fechar
    for processo in processos_ativos.values():
        try:
            if processo.poll() is None:  # Se ainda está rodando
                processo.terminate()
        except:
            pass
    
    cleanup_lock()
    root.destroy()
    os._exit(0)  # Força saída completa

# Verificar instância única
if not check_single_instance():
    messagebox.showerror("Erro", "PrintF já está em execução!\nFeche a outra instância antes de abrir novamente.")
    sys.exit(1)

# Configuração da janela principal
root = tk.Tk()
root.title("PrintF - Gerador de Evidências de Testes")
root.configure(bg="#f0f0f0")
root.protocol("WM_DELETE_WINDOW", fechar)  # Captura o evento de fechar janela

# Configurar para evitar múltiplas instâncias do Tkinter
root.wm_attributes("-topmost", 1)
root.resizable(False, False)

# Reduzir o tamanho da janela já que removemos um botão
largura_janela = 450  # Reduzido de 550 para 450
altura_janela = 60
root.geometry(f"{largura_janela}x{altura_janela}")

# Centralizar na tela
root.eval('tk::PlaceWindow . center')
root.geometry(f"+{root.winfo_screenwidth() - largura_janela - 20}+20")

# Frame para os botões
frame_botoes = tk.Frame(root, bg="#f0f0f0")
frame_botoes.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

def criar_botao_compacto(parent, texto, comando, cor="#ffffff"):
    return tk.Button(
        parent,
        text=texto,
        command=comando,
        font=("Arial", 9, "bold"),
        bg=cor,
        fg="#333",
        relief="flat",
        bd=1,
        activebackground="#d6e0f0",
        activeforeground="#000",
        padx=8,
        pady=4,
        width=12
    )

# Criar botões - REMOVIDO O BOTÃO "DOCX"
btn1 = criar_botao_compacto(frame_botoes, "📄 Templates", gerar_templates, "#4fc3f7")  # Azul
btn2 = criar_botao_compacto(frame_botoes, "📷 Gravar", capturar_evidencias, "#81c784")  # Verde
btn3 = criar_botao_compacto(frame_botoes, "🗑️ Limpar", limpar_arquivos, "#ffd54f")  # Amarelo
btn4 = criar_botao_compacto(frame_botoes, "❌ Fechar", fechar, "#f8d7da")  # Vermelho

# Posicionar botões - REMOVIDO O BOTÃO DOCX
btn1.pack(side=tk.LEFT, padx=2)
btn2.pack(side=tk.LEFT, padx=2)
btn3.pack(side=tk.LEFT, padx=2)
btn4.pack(side=tk.LEFT, padx=2)

# Implementação simples de tooltip
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)

    def enter(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        
        # Cria a tooltip
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", 
                        relief="solid", borderwidth=1, font=("Arial", 8))
        label.pack()

    def leave(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

# Criar tooltips simples - REMOVIDO TOOLTIP DO BOTÃO DOCX
ToolTip(btn1, "Gerar templates de documentação")
ToolTip(btn2, "Capturar evidências de testes")
ToolTip(btn3, "Limpar arquivos temporários")
ToolTip(btn4, "Fechar a barra de ferramentas")

def verificar_executaveis():
    """Verifica se os executáveis existem e atualiza os botões apenas se necessário"""
    executaveis = [
        ("gerador_templates.exe", btn1, "#4fc3f7"),
        ("gravador_evidencias.exe", btn2, "#81c784"),
        ("limpador_arquivos.exe", btn3, "#ffd54f")
    ]
    
    for exe_name, button, cor_original in executaveis:
        exe_path = get_exe_path(exe_name)
        if not os.path.exists(exe_path):
            # Apenas desabilita se o executável realmente não existir
            button.config(state=tk.DISABLED, bg="#cccccc")
        else:
            # Garante que o botão está habilitado e com a cor correta
            button.config(state=tk.NORMAL, bg=cor_original)

# Verificar executáveis imediatamente ao iniciar
verificar_executaveis()

# Função para restaurar a janela quando necessário
def restaurar_janela():
    """Restaura a janela principal se estiver minimizada"""
    try:
        root.deiconify()
    except:
        pass

# Garantir cleanup mesmo em falhas
import atexit
atexit.register(cleanup_lock)

# Iniciar aplicação
try:
    root.mainloop()
finally:
    cleanup_lock()