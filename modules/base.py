import tkinter as tk
from tkinter import ttk
import os
import sys

# Adiciona o diretório modules ao path
sys.path.append(os.path.join(os.path.dirname(__file__), 'modules'))

from modules.capture import CaptureModule
from modules.template_gen import TemplateGeneratorModule
from modules.evidence_gen import EvidenceGeneratorModule
from modules.cleanup import CleanupModule

class PrintFApp:
    """Aplicação principal unificada - APENAS 1 CLASSE"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PrintF")
        self.root.geometry("400x500")
        self.root.configure(bg='#f5f5f5')
        
        # Módulos - cada um é auto-contido
        self.modules = {
            'capture': CaptureModule(self.root),
            'templates': TemplateGeneratorModule(self.root),
            'evidence': EvidenceGeneratorModule(self.root),
            'cleanup': CleanupModule(self.root)
        }
        
        self.current_module = None
    
    def create_ui(self):
        """Cria interface principal"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title = tk.Label(main_frame, text="🖨️ PrintF Unificado", 
                        font=("Arial", 16, "bold"), bg='#f5f5f5')
        title.pack(pady=20)
        
        # Botões dos módulos
        modules_config = [
            ("📷 Capturar Evidências", "capture", "#27ae60"),
            ("📄 Gerar Templates", "templates", "#3498db"), 
            ("📋 Gerar Documentos", "evidence", "#f39c12"),
            ("🗑️ Limpar Arquivos", "cleanup", "#e74c3c"),
        ]
        
        for text, module_key, color in modules_config:
            btn = tk.Button(main_frame, text=text, 
                          command=lambda k=module_key: self.open_module(k),
                          bg=color, fg="white", font=("Arial", 12, "bold"),
                          width=25, height=2, relief="flat", cursor="hand2")
            btn.pack(pady=8)
    
    def open_module(self, module_key):
        """Abre um módulo específico"""
        # Fecha módulo atual se existir
        if self.current_module:
            self.current_module.hide()
        
        # Abre novo módulo
        self.current_module = self.modules[module_key]
        self.current_module.show()
    
    def run(self):
        """Executa a aplicação"""
        self.create_ui()
        self.root.mainloop()

if __name__ == "__main__":
    app = PrintFApp()
    app.run()