import csv
from docx import Document
import os
import pandas as pd
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import time

class GeradorTemplates:
    def __init__(self, root):
        self.root = root
        self.root.title("PrintF - Gerar Templates")
        self.root.geometry("800x700")
        self.root.configure(bg='#f0f0f0')
        
        self.criar_interface()
        
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar pesos para redimensionamento
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Título
        titulo = ttk.Label(main_frame, text="📄 PrintF - Gerar Templates", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Separador
        separator1 = ttk.Separator(main_frame, orient='horizontal')
        separator1.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        # Campos fixos
        ttk.Label(main_frame, text="Campos Fixos:", font=("Arial", 12, "bold")).grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        campos = [
            ("Demanda:", "-DEMANDA-"),
            ("Tipo Demanda:", "-TIPO_DEMANDA-"),
            ("STI:", "-STI-"),
            ("Chamado:", "-CHAMADO-"),
            ("Título:", "-TITULO-"),
            ("Caminho Funcionalidade:", "-CAMINHO-")
        ]

        for i, (label, key) in enumerate(campos):
            ttk.Label(main_frame, text=label).grid(row=3+i, column=0, sticky=tk.W, pady=2)
            entry = ttk.Entry(main_frame, width=40)
            entry.grid(row=3+i, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
            setattr(self, key, entry)

        # Separador
        separator2 = ttk.Separator(main_frame, orient='horizontal')
        separator2.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Arquivos
        ttk.Label(main_frame, text="Arquivos:", font=("Arial", 12, "bold")).grid(row=10, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        # CSV
        ttk.Label(main_frame, text="CSV:").grid(row=11, column=0, sticky=tk.W, pady=2)
        self.csv_entry = ttk.Entry(main_frame, width=40)
        self.csv_entry.grid(row=11, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        ttk.Button(main_frame, text="Procurar", command=self.selecionar_csv).grid(row=11, column=2, padx=(5, 0))

        # Template
        ttk.Label(main_frame, text="Template:").grid(row=12, column=0, sticky=tk.W, pady=2)
        self.template_entry = ttk.Entry(main_frame, width=40)
        self.template_entry.grid(row=12, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        ttk.Button(main_frame, text="Procurar", command=self.selecionar_template).grid(row=12, column=2, padx=(5, 0))

        # Pasta Saída
        ttk.Label(main_frame, text="Pasta Saída:").grid(row=13, column=0, sticky=tk.W, pady=2)
        self.pasta_entry = ttk.Entry(main_frame, width=40)
        self.pasta_entry.grid(row=13, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        ttk.Button(main_frame, text="Procurar", command=self.selecionar_pasta).grid(row=13, column=2, padx=(5, 0))

        # Separador
        separator3 = ttk.Separator(main_frame, orient='horizontal')
        separator3.grid(row=14, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=15, column=0, columnspan=3, pady=10)
        
        self.gerar_btn = ttk.Button(button_frame, text="▶️ Gerar Documentos", 
                                   command=self.iniciar_processamento, style='Accent.TButton')
        self.gerar_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🔄 Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="❌ Sair", command=self.root.quit).pack(side=tk.LEFT, padx=5)

        # Separador
        separator4 = ttk.Separator(main_frame, orient='horizontal')
        separator4.grid(row=16, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Progresso
        ttk.Label(main_frame, text="Progresso:").grid(row=17, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=18, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        # Log
        ttk.Label(main_frame, text="Log de Execução:").grid(row=19, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(main_frame, width=70, height=15, state='disabled')
        self.log_text.grid(row=20, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # Configurar redimensionamento
        main_frame.rowconfigure(20, weight=1)

        # Preencher template padrão se existir
        if os.path.exists('template_evidencias.docx'):
            self.template_entry.insert(0, 'template_evidencias.docx')

    def selecionar_csv(self):
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo CSV",
            filetypes=(("CSV Files", "*.csv"), ("Todos os arquivos", "*.*"))
        )
        if arquivo:
            self.csv_entry.delete(0, tk.END)
            self.csv_entry.insert(0, arquivo)

    def selecionar_template(self):
        arquivo = filedialog.askopenfilename(
            title="Selecionar template DOCX",
            filetypes=(("Word Documents", "*.docx"), ("Todos os arquivos", "*.*"))
        )
        if arquivo:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, arquivo)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta de saída")
        if pasta:
            self.pasta_entry.delete(0, tk.END)
            self.pasta_entry.insert(0, pasta)

    def limpar_campos(self):
        # Limpar todos os campos de entrada
        for entry in [self.csv_entry, self.template_entry, self.pasta_entry]:
            entry.delete(0, tk.END)
        
        for key in ["-DEMANDA-", "-TIPO_DEMANDA-", "-STI-", "-CHAMADO-", "-TITULO-", "-CAMINHO-"]:
            getattr(self, key).delete(0, tk.END)
        
        # Limpar log
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        # Resetar progresso
        self.progress['value'] = 0

    def log(self, mensagem):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, mensagem + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def limpar_nome_arquivo(self, nome):
        nome_limpo = re.sub(r'[<>:"/\\|?*]', '_', nome)
        if len(nome_limpo) > 100:
            nome_limpo = nome_limpo[:100]
        nome_limpo = nome_limpo.strip()
        return nome_limpo or "caso_teste"

    def ler_csv_complexo(self, arquivo_csv):
        try:
            encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'windows-1252']
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(arquivo_csv, encoding=encoding, engine='python', on_bad_lines='skip')
                    if 'Nome' in df.columns:
                        nomes = df['Nome'].dropna().str.strip()
                        nomes = nomes[nomes != ''].tolist()
                        return nomes
                except:
                    continue
            return self.ler_csv_manual(arquivo_csv)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o CSV: {str(e)}")
            return None

    def ler_csv_manual(self, arquivo_csv):
        nomes = []
        encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
        
        for encoding in encodings:
            try:
                with open(arquivo_csv, 'r', encoding=encoding) as file:
                    lines = file.readlines()
                    
                    if not lines:
                        continue
                    
                    header_line = None
                    for line in lines:
                        if 'Nome' in line:
                            header_line = line
                            break
                    
                    if not header_line:
                        continue
                    
                    headers = header_line.strip().split(',')
                    if 'Nome' not in headers:
                        continue
                    
                    nome_index = headers.index('Nome')
                    
                    for line in lines[1:]:
                        try:
                            if '"' in line:
                                reader = csv.reader([line])
                                parts = next(reader)
                            else:
                                parts = line.strip().split(',')
                            
                            if len(parts) > nome_index:
                                nome = parts[nome_index].strip().strip('"')
                                if nome and nome != '':
                                    nomes.append(nome)
                        except:
                            continue
                    
                    return nomes
            except:
                continue
        return None

    def preencher_template_estruturado(self, doc, dados):
        campos = {
            'Demanda': dados.get('Demanda', ''),
            'Tipo Demanda': dados.get('Tipo Demanda', ''),
            'STI': dados.get('STI', ''),
            'Chamado': dados.get('Chamado', ''),
            'Título': dados.get('Título', ''),
            'Caso de Teste': dados.get('Caso de Teste', ''),
            'Caminho da Funcionalidade': dados.get('Caminho da Funcionalidade', '')
        }
        
        for paragraph in doc.paragraphs:
            for campo, valor in campos.items():
                if paragraph.text.startswith(f"{campo}:"):
                    paragraph.text = f"{campo}: {valor}"
                    break
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for campo, valor in campos.items():
                            if paragraph.text.startswith(f"{campo}:"):
                                paragraph.text = f"{campo}: {valor}"
                                break

    def criar_template_exemplo(self):
        try:
            doc = Document()
            doc.add_heading('Evidências de Teste', level=1)
            doc.add_paragraph('Demanda:')
            doc.add_paragraph('Tipo Demanda:')
            doc.add_paragraph('STI:')
            doc.add_paragraph('Chamado:')
            doc.add_paragraph('Título:')
            doc.add_paragraph()
            
            table = doc.add_table(rows=3, cols=1)
            table.style = 'Table Grid'
            table.cell(0, 0).text = 'Caso de Teste:'
            table.cell(1, 0).text = 'Caminho da Funcionalidade:'
            table.cell(2, 0).text = 'Observações:'
            
            doc.save('template_evidencias.docx')
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar template: {str(e)}")
            return False

    def processar_documentos(self):
        try:
            # Obter valores dos campos
            dados_fixos = {
                'Demanda': getattr(self, '-DEMANDA-').get(),
                'Tipo Demanda': getattr(self, '-TIPO_DEMANDA-').get(),
                'STI': getattr(self, '-STI-').get(),
                'Chamado': getattr(self, '-CHAMADO-').get(),
                'Título': getattr(self, '-TITULO-').get(),
                'Caminho da Funcionalidade': getattr(self, '-CAMINHO-').get()
            }
            
            # Validar campos
            for campo, valor in dados_fixos.items():
                if not valor.strip():
                    messagebox.showerror("Erro", f"O campo '{campo}' é obrigatório!")
                    self.gerar_btn.config(state='normal')
                    return
            
            csv_path = self.csv_entry.get()
            template_path = self.template_entry.get()
            output_folder = self.pasta_entry.get() or 'evidencias_geradas'
            
            if not csv_path:
                messagebox.showerror("Erro", "Selecione um arquivo CSV!")
                self.gerar_btn.config(state='normal')
                return
            
            if not template_path:
                messagebox.showerror("Erro", "Selecione um template DOCX!")
                self.gerar_btn.config(state='normal')
                return
            
            if not os.path.exists(csv_path):
                messagebox.showerror("Erro", "Arquivo CSV não encontrado!")
                self.gerar_btn.config(state='normal')
                return
            
            if not os.path.exists(template_path):
                messagebox.showerror("Erro", "Template DOCX não encontrado!")
                self.gerar_btn.config(state='normal')
                return
            
            # Criar pasta de saída
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            self.log("Lendo arquivo CSV...")
            casos_teste = self.ler_csv_complexo(csv_path)
            
            if not casos_teste:
                messagebox.showerror("Erro", "Não foi possível ler os casos de teste do CSV")
                self.gerar_btn.config(state='normal')
                return
            
            self.log(f"Encontrados {len(casos_teste)} casos de teste\n")
            self.progress['maximum'] = len(casos_teste)
            
            sucessos = 0
            erros = []
            arquivos_gerados = set()
            
            for i, caso_teste in enumerate(casos_teste, 1):
                try:
                    self.progress['value'] = i
                    self.log(f"Processando: {caso_teste}")
                    
                    doc = Document(template_path)
                    dados_completos = dados_fixos.copy()
                    dados_completos['Caso de Teste'] = caso_teste
                    
                    self.preencher_template_estruturado(doc, dados_completos)
                    
                    nome_base = self.limpar_nome_arquivo(caso_teste)
                    nome_arquivo = f"Evidencia_{dados_fixos['Demanda']}_{nome_base}.docx"
                    
                    contador = 1
                    nome_original = nome_arquivo
                    while nome_arquivo in arquivos_gerados:
                        nome_arquivo = f"Evidencia_{dados_fixos['Demanda']}_{nome_base}_{contador}.docx"
                        contador += 1
                    
                    arquivos_gerados.add(nome_arquivo)
                    caminho_completo = os.path.join(output_folder, nome_arquivo)
                    
                    doc.save(caminho_completo)
                    
                    self.log(f"✅ Salvo: {nome_arquivo}\n")
                    sucessos += 1
                    
                except Exception as e:
                    self.log(f"❌ Erro: {str(e)}\n")
                    erros.append((caso_teste, str(e)))
                
                time.sleep(0.1)
            
            # Resultado final
            self.log("\n" + "=" * 50)
            self.log("🎉 PROCESSO CONCLUÍDO!")
            self.log("=" * 50)
            self.log(f"📊 Total processado: {len(casos_teste)}")
            self.log(f"✅ Sucessos: {sucessos}")
            self.log(f"❌ Erros: {len(erros)}")
            self.log(f"📁 Pasta: {os.path.abspath(output_folder)}")
            
            if sucessos > 0:
                self.log(f"\n📋 Arquivos gerados:")
                for arquivo in sorted(arquivos_gerados)[:10]:
                    self.log(f"• {arquivo}")
                if len(arquivos_gerados) > 10:
                    self.log(f"• ... e mais {len(arquivos_gerados) - 10} arquivos")
            
            self.gerar_btn.config(state='normal')
            
            if erros:
                messagebox.showinfo("Concluído", f"Processo concluído com {len(erros)} erros.\nVerifique o log para detalhes.")
            else:
                messagebox.showinfo("Sucesso", "Todos os documentos foram gerados com sucesso!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado: {str(e)}")
            self.gerar_btn.config(state='normal')

    def iniciar_processamento(self):
        self.gerar_btn.config(state='disabled')
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        self.progress['value'] = 0
        
        # Executar em thread separada
        thread = threading.Thread(target=self.processar_documentos, daemon=True)
        thread.start()

if __name__ == "__main__":
    # Verificar dependências básicas
    try:
        import pandas as pd
        from docx import Document
    except ImportError as e:
        print(f"❌ Dependências não instaladas: {e}")
        print("💡 Execute: pip install pandas python-docx")
        input("Pressione Enter para sair...")
        exit()
    
    root = tk.Tk()
    app = GeradorTemplates(root)
    root.mainloop()