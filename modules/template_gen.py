import csv
import json
import os
import re
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List, Optional, Tuple
from datetime import datetime

import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


class ConfigManager:
    """Gerencia o carregamento e salvamento da configuração de campos"""
    
    DEFAULT_CONFIG = [
        {"label": "Campo1:", "key": "campo1"},
        {"label": "Campo2:", "key": "campo2"},
        {"label": "Campo3:", "key": "campo3"},
        {"label": "Campo4:", "key": "campo4"},
        {"label": "Campo5:", "key": "campo5"},
        {"label": "Campo6", "key": "campo6"}
    ]

    def __init__(self, config_file: str = 'config_campos.json'):
        self.config_file = Path(config_file)

    def load_config(self) -> List[Dict]:
        """Carrega a configuração do arquivo JSON ou cria uma padrão"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    print(f"✅ Configuração carregada de '{self.config_file}'")
                    return config
            else:
                return self._create_default_config()
        except Exception as e:
            print(f"⚠️ Erro ao carregar configuração: {e}")
            return self.DEFAULT_CONFIG

    def _create_default_config(self) -> List[Dict]:
        """Cria arquivo de configuração padrão"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)
            print(f"ℹ️ Arquivo '{self.config_file}' criado com configuração padrão")
            return self.DEFAULT_CONFIG
        except Exception as e:
            print(f"❌ Erro ao criar configuração padrão: {e}")
            return self.DEFAULT_CONFIG


class CSVReader:
    """Responsável pela leitura de arquivos CSV com suporte a campos BDD - VERSÃO CORRIGIDA PARA ROVO"""
    
    ENCODINGS = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'windows-1252']

    @staticmethod
    def read_csv(file_path: str) -> Optional[List[Dict]]:
        """Lê um arquivo CSV e retorna lista de dicionários com dados completos"""
        try:
            # Primeiro tenta o método específico para CSV do Rovo
            result = CSVReader._read_rovo_csv(file_path)
            if result:
                return result
            
            # Fallback para métodos anteriores
            return CSVReader._read_with_pandas_advanced(file_path) or CSVReader._read_manual_advanced(file_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o CSV: {e}")
            return None

    @staticmethod
    def _read_rovo_csv(file_path: str) -> Optional[List[Dict]]:
        """Método específico para ler CSVs do Rovo com campos multilinha"""
        for encoding in CSVReader.ENCODINGS:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    content = file.read()
                
                # Processar linhas manualmente para lidar com campos multilinha
                lines = content.splitlines()
                if not lines:
                    return None
                
                # Obter cabeçalho
                headers = [h.strip() for h in lines[0].split(',')]
                
                resultados = []
                current_row = {}
                current_field = None
                buffer = []
                
                for line in lines[1:]:
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Verificar se é início de novo registro (começa com aspas)
                    if line.startswith('"') and not current_field:
                        # Finalizar registro anterior se existir
                        if current_row:
                            # Processar BDD do registro anterior
                            bdd_text = current_row.get('Script de Teste (BDD)', '')
                            cenario_bdd = CSVReader._parse_bdd_scenario(bdd_text)
                            current_row.update(cenario_bdd)
                            resultados.append(current_row)
                        
                        # Novo registro
                        current_row = {}
                        buffer = []
                        
                        # Parsear primeira linha do novo registro
                        if '"' in line:
                            reader = csv.reader([line])
                            parts = next(reader)
                        else:
                            parts = line.split(',')
                        
                        # Preencher campos iniciais
                        for i, header in enumerate(headers):
                            if i < len(parts):
                                current_row[header] = parts[i].strip().strip('"')
                            else:
                                current_row[header] = ""
                    
                    else:
                        # Continuidade do campo anterior (provavelmente BDD)
                        if 'Script de Teste (BDD)' in current_row and current_row['Script de Teste (BDD)']:
                            # Adicionar à linha atual do BDD
                            current_row['Script de Teste (BDD)'] += " " + line.strip().strip('"')
                        else:
                            # Tentar parsear como linha normal
                            if '"' in line:
                                reader = csv.reader([line])
                                parts = next(reader)
                            else:
                                parts = line.split(',')
                            
                            for i, header in enumerate(headers):
                                if i < len(parts) and header not in current_row:
                                    current_row[header] = parts[i].strip().strip('"')
                
                # Adicionar último registro
                if current_row:
                    bdd_text = current_row.get('Script de Teste (BDD)', '')
                    cenario_bdd = CSVReader._parse_bdd_scenario(bdd_text)
                    current_row.update(cenario_bdd)
                    resultados.append(current_row)
                
                return resultados if resultados else None
                
            except Exception as e:
                print(f"Encoding {encoding} falhou no método Rovo: {e}")
                continue
        return None

    @staticmethod
    def _read_with_pandas_advanced(file_path: str) -> Optional[List[Dict]]:
        """Tenta ler o CSV usando pandas com configuração específica para Rovo"""
        for encoding in CSVReader.ENCODINGS:
            try:
                # Configuração específica para CSV com campos multilinha
                df = pd.read_csv(
                    file_path, 
                    encoding=encoding, 
                    engine='python',
                    quotechar='"',
                    doublequote=True,
                    skipinitialspace=True,
                    on_bad_lines='skip'
                )
                
                if df.empty:
                    return None
                
                # Normalizar nomes de colunas
                df.columns = [col.strip() for col in df.columns]
                
                resultados = []
                for _, row in df.iterrows():
                    dados = {}
                    
                    # Coletar todos os campos disponíveis
                    for coluna in df.columns:
                        valor = str(row[coluna]).strip() if pd.notna(row[coluna]) else ""
                        dados[coluna] = valor
                    
                    # Processar cenário BDD se existir
                    bdd_text = ""
                    if 'Script de Teste (BDD)' in df.columns:
                        bdd_text = dados.get('Script de Teste (BDD)', '')
                    
                    # Extrair Given, When, Then do texto BDD
                    cenario_bdd = CSVReader._parse_bdd_scenario(bdd_text)
                    dados.update(cenario_bdd)
                    
                    resultados.append(dados)
                
                return resultados
                
            except Exception as e:
                print(f"Tentativa com encoding {encoding} falhou: {e}")
                continue
        return None

    @staticmethod
    def _read_manual_advanced(file_path: str) -> Optional[List[Dict]]:
        """Leitura manual do CSV como fallback - versão avançada"""
        for encoding in CSVReader.ENCODINGS:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    # Ler todo o conteúdo para lidar com quebras de linha
                    content = file.read()
                
                lines = content.splitlines()
                if not lines:
                    return None
                
                # Processar cabeçalho
                header_line = lines[0].strip()
                headers = []
                if '"' in header_line:
                    reader = csv.reader([header_line])
                    headers = next(reader)
                else:
                    headers = header_line.split(',')
                
                headers = [h.strip().strip('"') for h in headers]
                
                resultados = []
                i = 1
                while i < len(lines):
                    line = lines[i].strip()
                    if not line:
                        i += 1
                        continue
                    
                    dados = {}
                    
                    # Verificar se a linha começa com aspas (indicando novo registro)
                    if line.startswith('"'):
                        # Tentar parsear a linha atual
                        current_line = line
                        line_parts = []
                        
                        if '"' in current_line:
                            try:
                                reader = csv.reader([current_line])
                                line_parts = next(reader)
                            except:
                                line_parts = current_line.split(',')
                        else:
                            line_parts = current_line.split(',')
                        
                        # Preencher dados
                        for j, header in enumerate(headers):
                            if j < len(line_parts):
                                dados[header] = line_parts[j].strip().strip('"')
                            else:
                                dados[header] = ""
                        
                        # Verificar se o campo BDD continua nas próximas linhas
                        if 'Script de Teste (BDD)' in headers:
                            bdd_index = headers.index('Script de Teste (BDD)')
                            if bdd_index < len(line_parts):
                                bdd_content = line_parts[bdd_index]
                                # Procurar por continuação do BDD nas próximas linhas
                                k = i + 1
                                while k < len(lines) and not lines[k].strip().startswith('"'):
                                    bdd_content += " " + lines[k].strip().strip('"')
                                    k += 1
                                dados['Script de Teste (BDD)'] = bdd_content
                                i = k - 1  # Ajustar índice
                    
                    else:
                        # Linha normal
                        if '"' in line:
                            try:
                                reader = csv.reader([line])
                                parts = next(reader)
                            except:
                                parts = line.split(',')
                        else:
                            parts = line.split(',')
                        
                        for j, header in enumerate(headers):
                            if j < len(parts):
                                dados[header] = parts[j].strip().strip('"')
                            else:
                                dados[header] = ""
                    
                    # Processar cenário BDD se existir
                    bdd_text = dados.get('Script de Teste (BDD)', '')
                    cenario_bdd = CSVReader._parse_bdd_scenario(bdd_text)
                    dados.update(cenario_bdd)
                    
                    resultados.append(dados)
                    i += 1
                
                return resultados if resultados else None
                
            except Exception as e:
                print(f"Encoding {encoding} falhou: {e}")
                continue
        return None

    @staticmethod
    def _parse_bdd_scenario(bdd_text: str) -> Dict[str, str]:
        """Extrai Given, When, Then do texto BDD - VERSÃO MELHORADA"""
        cenario = {
            'Given': '',
            'When': '', 
            'Then': '',
            'And': ''
        }
        
        if not bdd_text or bdd_text == 'N/A':
            return cenario
        
        # Limpar e normalizar o texto
        bdd_text = ' '.join(bdd_text.split())  # Remove quebras de linha múltiplas
        bdd_text = bdd_text.replace('\n', ' ').replace('\r', ' ')
        
        # Padrões para encontrar as seções BDD (case insensitive)
        patterns = {
            'Given': r'Given\s+(.*?)(?=When|Then|And|$)',
            'When': r'When\s+(.*?)(?=Then|And|Given|$)',
            'Then': r'Then\s+(.*?)(?=And|When|Given|$)',
            'And': r'And\s+(.*?)(?=Then|When|Given|$)'
        }
        
        for key, pattern in patterns.items():
            matches = re.findall(pattern, bdd_text, re.IGNORECASE | re.DOTALL)
            if matches:
                # Juntar múltiplas ocorrências e limpar
                extracted = ' '.join([match.strip() for match in matches if match.strip()])
                # Remover pontuação extra no final
                cenario[key] = extracted.rstrip('.,;')
        
        # Log para debug
        if any(cenario.values()):
            print(f"📋 BDD extraído: {cenario}")
        
        return cenario


class DocumentProcessor:
    """Processa e gera documentos Word baseados em templates"""
    
    @staticmethod
    def clean_filename(filename: str, max_length: int = 100) -> str:
        """Limpa o nome do arquivo removendo caracteres inválidos"""
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', filename)
        cleaned = cleaned.strip()[:max_length]
        return cleaned or "caso_teste"

    @staticmethod
    def adjust_template_fields(doc: Document, field_mapping: Dict[str, str]) -> None:
        """Ajusta os campos do template conforme o mapeamento fornecido"""
        for paragraph in doc.paragraphs:
            DocumentProcessor._adjust_paragraph_fields(paragraph, field_mapping)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        DocumentProcessor._adjust_paragraph_fields(paragraph, field_mapping)

    @staticmethod
    def _adjust_paragraph_fields(paragraph, field_mapping: Dict[str, str]) -> None:
        """Ajusta os campos em um parágrafo específico"""
        texto_original = paragraph.text.strip()
        if ':' in texto_original:
            field_key = texto_original.split(':', 1)[0].strip()
            if field_key in field_mapping:
                paragraph.text = f"{field_mapping[field_key]}: "

    @staticmethod
    def fill_template(doc: Document, data: Dict[str, str], field_mapping: Dict[str, str]) -> bool:
        """Preenche o template com os dados fornecidos - INCLUINDO CAMPOS BDD
        Retorna True se algum campo BDD foi preenchido"""
        
        # Primeiro ajusta os campos do template de acordo com as labels do JSON
        DocumentProcessor.adjust_template_fields(doc, field_mapping)
        
        # Cria mapeamento label -> valor para campos configurados
        label_to_value = {}
        for original_key, label in field_mapping.items():
            label_to_value[label] = data.get(original_key, '')
        
        # Adicionar campos especiais
        label_to_value['Caso de Teste'] = data.get('Caso de Teste', '')
        
        # Adicionar campos BDD se existirem nos dados
        bdd_fields = ['Given', 'When', 'Then', 'And', 'Script de Teste (BDD)']
        for field in bdd_fields:
            if field in data and data[field]:
                # Usar o próprio nome do campo como label
                label_to_value[field] = data[field]
        
        # Adicionar outros campos comuns do CSV
        common_fields = ['Objetivo', 'Pré-condição', 'Precondição', 'Status', 'Priority']
        for field in common_fields:
            if field in data and data[field] and data[field] != 'N/A':
                label_to_value[field] = data[field]
        
        # Preenche parágrafos e verifica se algum campo BDD foi preenchido
        bdd_was_filled = False
        for paragraph in doc.paragraphs:
            if DocumentProcessor._fill_paragraph(paragraph, label_to_value):
                bdd_was_filled = True
        
        # Preenche tabelas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if DocumentProcessor._fill_paragraph(paragraph, label_to_value):
                            bdd_was_filled = True
        
        return bdd_was_filled

    @staticmethod
    def _fill_paragraph(paragraph, label_to_value: Dict[str, str]) -> bool:
        """Preenche um parágrafo específico com os dados. Retorna True se preencheu campo BDD."""
        texto = paragraph.text.strip()
        bdd_was_filled = False
        
        # Verificar se o parágrafo contém algum campo que precisa ser preenchido
        for field_name, value in label_to_value.items():
            # Padrão: "Field Name: " ou "Field Name:" 
            patterns = [
                f"{field_name}: ",
                f"{field_name}:",
                f"{field_name} : ",
                f"{field_name} :"
            ]
            
            for pattern in patterns:
                if pattern in texto:
                    # Substituir mantendo a formatação
                    paragraph.text = paragraph.text.replace(pattern, f"{field_name}: {value}")
                    # Verificar se era um campo BDD
                    if field_name in ['Given', 'When', 'Then', 'And', 'Script de Teste (BDD)']:
                        bdd_was_filled = True
                    return bdd_was_filled
        
        # Fallback: se o texto terminar com ":" e corresponder a um campo
        if ':' in texto:
            field_name = texto.split(':', 1)[0].strip()
            if field_name in label_to_value:
                paragraph.text = f"{field_name}: {label_to_value[field_name]}"
                if field_name in ['Given', 'When', 'Then', 'And', 'Script de Teste (BDD)']:
                    bdd_was_filled = True
        
        return bdd_was_filled

    @staticmethod
    def add_bdd_section_to_template(doc: Document, data: Dict[str, str]) -> None:
        """Adiciona uma seção BDD no final do documento template se não existir"""
        # Verificar se existem dados BDD
        has_bdd_data = any(data.get(key) for key in ['Given', 'When', 'Then', 'And', 'Script de Teste (BDD)'])
        
        if not has_bdd_data:
            return
        
        # Verificar se já existe uma seção BDD no documento
        has_existing_bdd = False
        for paragraph in doc.paragraphs:
            if 'BDD' in paragraph.text or 'Behavior Driven Development' in paragraph.text:
                has_existing_bdd = True
                break
        
        if has_existing_bdd:
            return
        
        # Adicionar quebra de página
        doc.add_page_break()
        
        # Adicionar título da seção BDD
        title = doc.add_heading('Cenário BDD (Behavior Driven Development)', level=1)
        
        # Coletar dados BDD
        bdd_data = []
        
        # Adicionar Given se existir
        given_text = data.get('Given', '')
        if given_text:
            bdd_data.append(('Given', given_text))
        
        # Adicionar When se existir
        when_text = data.get('When', '')
        if when_text:
            bdd_data.append(('When', when_text))
        
        # Adicionar Then se existir
        then_text = data.get('Then', '')
        if then_text:
            bdd_data.append(('Then', then_text))
        
        # Adicionar And se existir
        and_text = data.get('And', '')
        if and_text:
            bdd_data.append(('And', and_text))
        
        # Adicionar cenário completo se existir e for diferente dos dados extraídos
        cenario_completo = data.get('Script de Teste (BDD)', '')
        if cenario_completo and cenario_completo != 'N/A':
            # Verificar se o cenário completo é diferente dos dados extraídos
            partes_extraidas = [given_text, when_text, then_text, and_text]
            texto_extraido = ' '.join(filter(None, partes_extraidas))
            
            if texto_extraido.strip() != cenario_completo.strip():
                bdd_data.append(('Cenário Completo', cenario_completo))
        
        if not bdd_data:
            return
        
        # Tabela para cenário BDD
        bdd_table = doc.add_table(rows=len(bdd_data), cols=2)
        bdd_table.style = 'Light Grid Accent 1'
        
        # Configurar largura das colunas
        bdd_table.columns[0].width = Inches(1.5)
        bdd_table.columns[1].width = Inches(5.5)
        
        for i, (etapa, descricao) in enumerate(bdd_data):
            cells = bdd_table.rows[i].cells
            cells[0].text = etapa
            cells[1].text = descricao
            
            # Formatar célula da etapa em negrito
            for paragraph in cells[0].paragraphs:
                for run in paragraph.runs:
                    run.bold = True


class DefaultDocumentGenerator:
    """Gera documentos padrão quando nenhum template é fornecido"""
    
    @staticmethod
    def create_default_document(data: Dict[str, str], field_config: List[Dict]) -> Document:
        """Cria um documento padrão com estrutura organizada incluindo BDD"""
        doc = Document()
        
        # Título do documento
        title = doc.add_heading('Evidências de Teste - Documentação', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Adicionar data e hora
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        date_para = doc.add_paragraph(f"Gerado em: {current_time}")
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        
        # Seção de informações do teste
        doc.add_heading('Informações do Teste', level=1)
        
        # Tabela para dados organizados
        table = doc.add_table(rows=len(field_config) + 1, cols=2)
        table.style = 'Light Grid Accent 1'
        
        # Cabeçalho da tabela
        header_cells = table.rows[0].cells
        header_cells[0].text = "Campo"
        header_cells[1].text = "Valor"
        
        # Formatar cabeçalho
        for cell in header_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
        
        # Preencher dados da configuração
        for i, campo_info in enumerate(field_config, 1):
            key = campo_info['key']
            label = campo_info['label'].rstrip(':')
            
            row_cells = table.rows[i].cells
            row_cells[0].text = label
            row_cells[1].text = data.get(key, 'Não informado')
        
        doc.add_paragraph()
        
        # Seção do caso de teste
        doc.add_heading('Caso de Teste', level=1)
        caso_teste_para = doc.add_paragraph()
        caso_teste_para.add_run('Nome do Caso de Teste: ').bold = True
        caso_teste_para.add_run(data.get('Caso de Teste', 'Não informado'))
        
        # Seção de objetivo se disponível
        objetivo = data.get('Objetivo', '')
        if objetivo and objetivo != 'N/A':
            doc.add_heading('Objetivo do Teste', level=2)
            doc.add_paragraph(objetivo)
        
        # Seção de pré-condições se disponível
        pre_condicao = data.get('Pré-condição', data.get('Precondição', ''))
        if pre_condicao and pre_condicao != 'N/A':
            doc.add_heading('Pré-condições', level=2)
            doc.add_paragraph(pre_condicao)
        
        # SEÇÃO BDD - SÓ ADICIONA SE HOUVER DADOS BDD
        has_bdd_data = DefaultDocumentGenerator._add_bdd_section(doc, data)
        
        # Seção de descrição do teste
        doc.add_heading('Descrição do Teste', level=2)
        if has_bdd_data:
            doc.add_paragraph(
                "O cenário de teste foi definido utilizando a metodologia BDD (Behavior Driven Development) "
                "na seção acima. Esta seção deve conter a descrição detalhada da execução do teste, "
                "incluindo os passos executados e evidências coletadas."
            )
        else:
            doc.add_paragraph(
                "Esta seção deve conter a descrição detalhada do caso de teste executado, "
                "incluindo pré-condições, passos de execução e resultados esperados."
            )
        
        # Seção de evidências
        doc.add_heading('Evidências Coletadas', level=2)
        doc.add_paragraph("Registro das evidências coletadas durante a execução do teste:")
        
        # Tabela para evidências
        evidencias_table = doc.add_table(rows=5, cols=3)
        evidencias_table.style = 'Light Grid Accent 1'
        
        # Cabeçalho da tabela de evidências
        evidencias_header = evidencias_table.rows[0].cells
        headers = ['Etapa', 'Evidência', 'Resultado']
        for col, header in enumerate(headers):
            evidencias_header[col].text = header
            for paragraph in evidencias_header[col].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
        
        # Linhas para preenchimento
        etapas = [
            'Pré-condições',
            'Configuração Inicial', 
            'Execução do Teste',
            'Pós-condições',
            'Resultado Final'
        ]
        
        for row, etapa in enumerate(etapas, 1):
            row_cells = evidencias_table.rows[row].cells
            row_cells[0].text = etapa
            row_cells[1].text = "[Descreva a evidência coletada]"
            row_cells[2].text = "[Resultado obtido - OK/Erro]"
        
        doc.add_paragraph()
        
        # Seção de observações
        doc.add_heading('Observações e Comentários', level=2)
        doc.add_paragraph("Adicione observações relevantes sobre a execução do teste:")
        
        # Área para observações
        obs_para = doc.add_paragraph()
        obs_para.add_run("Observações Gerais:\n").bold = True
        obs_para.add_run("• [Insira observações sobre problemas encontrados]\n")
        obs_para.add_run("• [Comentários sobre o comportamento do sistema]\n")
        obs_para.add_run("• [Sugestões de melhorias]\n")
        obs_para.add_run("• [Outras informações relevantes]")
        
        # Rodapé informativo
        doc.add_paragraph()
        footer = doc.add_paragraph()
        footer.add_run("Documento gerado automaticamente pelo PrintF - Gerador de Templates").italic = True
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return doc

    @staticmethod
    def _add_bdd_section(doc: Document, data: Dict[str, str]) -> bool:
        """Adiciona seção BDD ao documento se existirem dados. Retorna True se adicionou dados BDD."""
        # Verificar se existem dados BDD
        has_bdd_data = any(data.get(key) for key in ['Given', 'When', 'Then', 'And', 'Script de Teste (BDD)'])
        
        if not has_bdd_data:
            return False
        
        doc.add_heading('Cenário BDD (Behavior Driven Development)', level=2)
        
        # Coletar dados BDD
        bdd_data = []
        
        # Adicionar Given se existir
        given_text = data.get('Given', '')
        if given_text:
            bdd_data.append(('Given', given_text))
        
        # Adicionar When se existir
        when_text = data.get('When', '')
        if when_text:
            bdd_data.append(('When', when_text))
        
        # Adicionar Then se existir
        then_text = data.get('Then', '')
        if then_text:
            bdd_data.append(('Then', then_text))
        
        # Adicionar And se existir
        and_text = data.get('And', '')
        if and_text:
            bdd_data.append(('And', and_text))
        
        # Adicionar cenário completo se existir e for diferente dos dados extraídos
        cenario_completo = data.get('Script de Teste (BDD)', '')
        if cenario_completo and cenario_completo != 'N/A':
            # Verificar se o cenário completo é diferente dos dados extraídos
            partes_extraidas = [given_text, when_text, then_text, and_text]
            texto_extraido = ' '.join(filter(None, partes_extraidas))
            
            if texto_extraido.strip() != cenario_completo.strip():
                bdd_data.append(('Cenário Completo', cenario_completo))
        
        if not bdd_data:
            return False
        
        # Tabela para cenário BDD
        bdd_table = doc.add_table(rows=len(bdd_data), cols=2)
        bdd_table.style = 'Light Grid Accent 1'
        
        # Configurar largura das colunas
        bdd_table.columns[0].width = Inches(1.5)
        bdd_table.columns[1].width = Inches(5.5)
        
        for i, (etapa, descricao) in enumerate(bdd_data):
            cells = bdd_table.rows[i].cells
            cells[0].text = etapa
            cells[1].text = descricao
            
            # Formatar célula da etapa em negrito
            for paragraph in cells[0].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
        
        doc.add_paragraph()
        return True


class TemplateGenerator:
    """Gerador de templates de exemplo"""
    
    @staticmethod
    def create_example_template(field_config: List[Dict]) -> bool:
        """Cria um template de exemplo com base na configuração"""
        try:
            doc = Document()
            doc.add_heading('Template de Evidências de Teste', level=1)
            
            # Adicionar instruções
            info_para = doc.add_paragraph()
            info_para.add_run("Instruções: ").bold = True
            info_para.add_run("Este é um template de exemplo. Os campos abaixo serão preenchidos automaticamente.")
            
            doc.add_paragraph()
            
            # Adicionar campos da configuração
            for campo_info in field_config:
                doc.add_paragraph(f"{campo_info['label']} [VALOR]")
            
            doc.add_paragraph()
            
            # Seção para caso de teste
            doc.add_heading('Detalhes do Caso de Teste', level=2)
            doc.add_paragraph("Caso de Teste: [NOME_DO_CASO]")
            
            # Seção BDD
            doc.add_heading('Cenário BDD', level=3)
            doc.add_paragraph("Given: [PRÉ-CONDIÇÕES]")
            doc.add_paragraph("When: [AÇÃO]")
            doc.add_paragraph("Then: [RESULTADO ESPERADO]")
            doc.add_paragraph("And: [CONDIÇÕES ADICIONAIS]")
            
            # Tabela para evidências
            table = doc.add_table(rows=4, cols=2)
            table.style = 'Table Grid'
            table.cell(0, 0).text = 'Caminho da Funcionalidade:'
            table.cell(0, 1).text = '[CAMINHO]'
            table.cell(1, 0).text = 'Resultado Esperado:'
            table.cell(1, 1).text = '[RESULTADO_ESPERADO]'
            table.cell(2, 0).text = 'Resultado Obtido:'
            table.cell(2, 1).text = '[RESULTADO_OBTIDO]'
            table.cell(3, 0).text = 'Observações:'
            table.cell(3, 1).text = '[OBSERVACOES]'
            
            doc.save('template_evidencias.docx')
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar template: {e}")
            return False


class TemplateGeneratorModule:
    """Interface principal da aplicação"""
    
    def __init__(self, parent, settings=None):
        self.parent = parent
        self.settings = settings or {}
        self.window = None
        
        self.config_manager = ConfigManager()
        self.csv_reader = CSVReader()
        self.doc_processor = DocumentProcessor()
        self.default_doc_generator = DefaultDocumentGenerator()
        
        self.campos_config = self.config_manager.load_config()
        self.campos_entries: Dict[str, tk.Entry] = {}
        
        # Variável para controle de diretório automático
        self.auto_directory_var = tk.BooleanVar(value=True)
    
    def show(self):
        """Mostra interface completa"""
        if self.window and self.window.winfo_exists():
            self.window.lift()
            return
        
        self.window = tk.Toplevel(self.parent)
        self.window.title("PrintF - Gerador de Templates")
        self.window.geometry("1000x800")
        self.window.minsize(900, 700)
        
        # Configurar cor de fundo
        try:
            from modules.styles import LiquidGlassStyle
            self.window.configure(bg=LiquidGlassStyle.BG_PRIMARY)
        except ImportError:
            self.window.configure(bg='#f0f0f0')
        
        self._create_complete_ui()

    def hide(self):
        """Esconde o módulo"""
        if self.window and self.window.winfo_exists():
            self.window.destroy()
        self.window = None

    def _create_complete_ui(self):
        """Cria a interface completa do módulo"""
        # Frame principal
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self._configure_grid_weights()
        self._create_title_section(main_frame)
        self._create_dynamic_fields_section(main_frame)
        self._create_file_section(main_frame)
        self._create_control_buttons(main_frame)
        self._create_progress_section(main_frame)
        self._create_log_section(main_frame)
        
        self._set_default_template()

    def _configure_grid_weights(self) -> None:
        """Configura os pesos do grid para redimensionamento"""
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

    def _create_title_section(self, parent) -> None:
        """Cria a seção do título"""
        titulo = ttk.Label(parent, text="📄 PrintF - Gerar Templates", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        ttk.Separator(parent, orient='horizontal').grid(
            row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

    def _create_dynamic_fields_section(self, parent) -> None:
        """Cria os campos dinâmicos baseados na configuração"""
        ttk.Label(parent, text="Dados dos Testes:", 
                 font=("Arial", 12, "bold")).grid(
                     row=2, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        for i, campo_info in enumerate(self.campos_config):
            self._create_field_row(parent, campo_info, i)

    def _create_field_row(self, parent, campo_info: Dict, row_index: int) -> None:
        """Cria uma linha de campo na interface"""
        label_text = campo_info['label']
        campo_key = campo_info['key']
        
        ttk.Label(parent, text=label_text).grid(
            row=3 + row_index, column=0, sticky=tk.W, pady=2)
        
        entry = ttk.Entry(parent, width=40)
        entry.grid(row=3 + row_index, column=1, columnspan=2, 
                  sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        self.campos_entries[campo_key] = entry

    def _create_file_section(self, parent) -> None:
        """Cria a seção de seleção de arquivos"""
        next_row = 3 + len(self.campos_config)
        
        ttk.Separator(parent, orient='horizontal').grid(
            row=next_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(parent, text="Arquivos:", 
                 font=("Arial", 12, "bold")).grid(
                     row=next_row + 1, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        # Campos de arquivo
        self.csv_entry = self._create_file_field(parent, "CSV:*", next_row + 2, self.selecionar_csv)
        self.template_entry = self._create_file_field(parent, "Template (Opcional):", next_row + 3, self.selecionar_template)
        
        # Nova opção para diretório automático
        directory_frame = ttk.Frame(parent)
        directory_frame.grid(row=next_row + 4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Label(directory_frame, text="Pasta Saída:").pack(side=tk.LEFT)
        
        self.pasta_entry = ttk.Entry(directory_frame, width=40)
        self.pasta_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        ttk.Button(directory_frame, text="Procurar", command=self.selecionar_pasta).pack(side=tk.LEFT)
        
        # Checkbox para diretório automático
        self.auto_dir_check = ttk.Checkbutton(
            directory_frame, 
            text="Criar automaticamente com nome do template", 
            variable=self.auto_directory_var,
            command=self._toggle_auto_directory
        )
        self.auto_dir_check.pack(side=tk.LEFT, padx=(10, 0))

        # Info sobre campos obrigatórios
        info_label = ttk.Label(parent, text="* Campos obrigatórios", font=("Arial", 9), foreground="gray")
        info_label.grid(row=next_row + 5, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

    def _toggle_auto_directory(self):
        """Ativa/desativa campo de pasta quando usar diretório automático"""
        if self.auto_directory_var.get():
            self.pasta_entry.config(state='disabled')
            # Preencher automaticamente se tiver template
            template_path = self.template_entry.get().strip()
            if template_path and Path(template_path).exists():
                self._update_auto_directory()
        else:
            self.pasta_entry.config(state='normal')

    def _update_auto_directory(self):
        """Atualiza o diretório automático baseado no template"""
        if not self.auto_directory_var.get():
            return
            
        template_path = self.template_entry.get().strip()
        if template_path and Path(template_path).exists():
            template_name = Path(template_path).stem
            auto_dir = f"evidencias_{template_name}"
            self.pasta_entry.delete(0, tk.END)
            self.pasta_entry.insert(0, auto_dir)

    def _create_file_field(self, parent, label: str, row: int, command) -> ttk.Entry:
        """Cria um campo de seleção de arquivo"""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=2)
        entry = ttk.Entry(parent, width=40)
        entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        ttk.Button(parent, text="Procurar", command=command).grid(
            row=row, column=2, padx=(5, 0))
        return entry

    def _create_control_buttons(self, parent) -> None:
        """Cria os botões de controle"""
        next_row = 3 + len(self.campos_config) + 6
        
        ttk.Separator(parent, orient='horizontal').grid(
            row=next_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=next_row + 1, column=0, columnspan=3, pady=10)
        
        self.gerar_btn = ttk.Button(button_frame, text="▶️ Gerar Documentos", 
                                   command=self.iniciar_processamento)
        self.gerar_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🔄 Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="❌ Sair", command=self.hide).pack(side=tk.LEFT, padx=5)

    def _create_progress_section(self, parent) -> None:
        """Cria a seção de progresso"""
        next_row = 3 + len(self.campos_config) + 8
        
        ttk.Separator(parent, orient='horizontal').grid(
            row=next_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(parent, text="Progresso:").grid(
            row=next_row + 1, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        self.progress = ttk.Progressbar(parent, mode='determinate')
        self.progress.grid(row=next_row + 2, column=0, columnspan=3, 
                          sticky=(tk.W, tk.E), pady=(0, 10))

    def _create_log_section(self, parent) -> None:
        """Cria a seção de log"""
        next_row = 3 + len(self.campos_config) + 10
        
        ttk.Label(parent, text="Log de Execução:").grid(
            row=next_row, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(parent, width=70, height=15, state='disabled')
        self.log_text.grid(row=next_row + 1, column=0, columnspan=3, 
                          sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        parent.rowconfigure(next_row + 1, weight=1)

    def _set_default_template(self) -> None:
        """Preenche template padrão se existir - AGORA CRIA AUTOMATICAMENTE SE NÃO EXISTIR"""
        template_path = 'template_evidencias.docx'
        if not Path(template_path).exists():
            # Cria template automaticamente se não existir
            self.log("📝 Criando template padrão automaticamente...")
            if self._criar_template_exemplo_automatico():
                self.template_entry.insert(0, template_path)
                self.log("✅ Template padrão criado com sucesso!")
                # Atualizar diretório automático se estiver ativo
                if self.auto_directory_var.get():
                    self._update_auto_directory()
        else:
            self.template_entry.insert(0, template_path)
            if self.auto_directory_var.get():
                self._update_auto_directory()

    def _criar_template_exemplo_automatico(self) -> bool:
        """Cria template de exemplo automaticamente (sem interação do usuário)"""
        try:
            return TemplateGenerator.create_example_template(self.campos_config)
        except Exception as e:
            self.log(f"⚠️ Aviso: Não foi possível criar template automático: {e}")
            return False

    # Métodos de seleção de arquivos
    def selecionar_csv(self) -> None:
        self._select_file(self.csv_entry, "Selecionar arquivo CSV", 
                         [("CSV Files", "*.csv"), ("Todos os arquivos", "*.*")])

    def selecionar_template(self) -> None:
        arquivo = filedialog.askopenfilename(
            title="Selecionar template DOCX", 
            filetypes=[("Word Documents", "*.docx"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, arquivo)
            # Atualizar diretório automático se estiver ativo
            if self.auto_directory_var.get():
                self._update_auto_directory()

    def selecionar_pasta(self) -> None:
        pasta = filedialog.askdirectory(title="Selecionar pasta de saída")
        if pasta:
            self.pasta_entry.delete(0, tk.END)
            self.pasta_entry.insert(0, pasta)
            # Desmarcar auto diretório se usuário selecionar manualmente
            self.auto_directory_var.set(False)
            self.pasta_entry.config(state='normal')

    def _select_file(self, entry_widget: ttk.Entry, title: str, filetypes: List[Tuple]) -> None:
        """Seleciona um arquivo e atualiza o campo de entrada"""
        arquivo = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if arquivo:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, arquivo)

    def limpar_campos(self) -> None:
        """Limpa todos os campos da interface"""
        for entry in [self.csv_entry, self.template_entry, self.pasta_entry]:
            entry.delete(0, tk.END)
        
        for entry in self.campos_entries.values():
            entry.delete(0, tk.END)
        
        self._clear_log()
        self.progress['value'] = 0
        # Restaurar auto directory
        self.auto_directory_var.set(True)
        self.pasta_entry.config(state='disabled')
        self._set_default_template()

    def _clear_log(self) -> None:
        """Limpa o log de execução"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def log(self, mensagem: str) -> None:
        """Adiciona mensagem ao log"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, mensagem + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.window.update()

    def _get_fixed_data(self) -> Dict[str, str]:
        """Obtém os dados dos campos fixos"""
        dados = {}
        for campo_key, entry in self.campos_entries.items():
            valor = entry.get().strip()
            if not valor:
                raise ValueError(f"O campo '{campo_key}' é obrigatório!")
            dados[campo_key] = valor
        return dados

    def _validate_inputs(self, csv_path: str) -> bool:
        """Valida os inputs necessários (apenas CSV é obrigatório)"""
        if not csv_path:
            messagebox.showerror("Erro", "Selecione um arquivo CSV!")
            return False
        
        if not Path(csv_path).exists():
            messagebox.showerror("Erro", "Arquivo CSV não encontrado!")
            return False
        
        return True

    def _get_output_directory(self, template_path: str) -> str:
        """Determina o diretório de saída baseado nas configurações"""
        if self.auto_directory_var.get() and template_path:
            # Usar nome do template para criar diretório automático
            template_name = Path(template_path).stem
            return f"evidencias_{template_name}"
        else:
            # Usar diretório especificado pelo usuário ou padrão
            output_folder = self.pasta_entry.get().strip()
            return output_folder or 'evidencias_geradas'

    def processar_documentos(self) -> None:
        """Processa os documentos em lote - versão robusta que nunca falha"""
        try:
            # Obter dados fixos
            try:
                dados_fixos = self._get_fixed_data()
            except ValueError as e:
                messagebox.showerror("Erro", str(e))
                self.gerar_btn.config(state='normal')
                return
            
            csv_path = self.csv_entry.get()
            template_path = self.template_entry.get().strip()
            
            if not self._validate_inputs(csv_path):
                self.gerar_btn.config(state='normal')
                return

            # Determinar pasta de saída
            output_folder = self._get_output_directory(template_path)
            
            # Garantir que temos um template válido
            template_path = self._garantir_template_valido(template_path)
            
            # Criar pasta de saída
            try:
                Path(output_folder).mkdir(exist_ok=True)
                self.log(f"📁 Pasta de saída: {output_folder}")
            except Exception as e:
                self.log(f"⚠️ Aviso: Não foi possível criar a pasta '{output_folder}': {e}")
                self.log("📁 Usando pasta atual para salvar os documentos...")
                output_folder = '.'
            
            self.log("📖 Lendo arquivo CSV...")
            dados_csv = self.csv_reader.read_csv(csv_path)
            
            if not dados_csv:
                messagebox.showerror("Erro", "Não foi possível ler os dados do CSV")
                self.gerar_btn.config(state='normal')
                return
            
            # Log detalhado sobre o que foi lido
            self.log(f"📊 Total de registros lidos: {len(dados_csv)}")
            
            # Verificar se há dados BDD
            casos_com_bdd = 0
            for item in dados_csv:
                if any(item.get(key) for key in ['Given', 'When', 'Then', 'And', 'Script de Teste (BDD)']):
                    casos_com_bdd += 1
            
            self.log(f"📋 Casos com dados BDD: {casos_com_bdd}")
            
            # Extrair casos de teste
            casos_teste = []
            for item in dados_csv:
                if 'Nome' in item and item['Nome'].strip():
                    casos_teste.append(item)
            
            if not casos_teste:
                messagebox.showerror("Erro", "Não foi possível encontrar casos de teste no CSV")
                self.gerar_btn.config(state='normal')
                return
            
            # Determinar modo de operação
            use_default_template = True
            if template_path and Path(template_path).exists():
                use_default_template = False
                self.log("📁 Usando template personalizado...")
            else:
                self.log("📝 Gerando documentos com template padrão...")
            
            self._process_test_cases(casos_teste, dados_fixos, template_path, output_folder, use_default_template)
            
        except Exception as e:
            self.log(f"❌ Erro inesperado: {e}")
            messagebox.showerror("Erro", f"Erro inesperado: {e}")
            self.gerar_btn.config(state='normal')

    def _garantir_template_valido(self, template_path: str) -> str:
        """Garante que temos um template válido, criando automaticamente se necessário"""
        if not template_path or not Path(template_path).exists():
            self.log("📝 Nenhum template válido encontrado, criando automaticamente...")
            if self._criar_template_exemplo_automatico():
                new_template_path = 'template_evidencias.docx'
                self.template_entry.delete(0, tk.END)
                self.template_entry.insert(0, new_template_path)
                self.log("✅ Template padrão criado e configurado automaticamente")
                return new_template_path
            else:
                self.log("⚠️ Não foi possível criar template, usando geração padrão...")
                return ""
        return template_path

    def _process_test_cases(self, casos_teste: List[Dict], dados_fixos: Dict[str, str], 
                           template_path: str, output_folder: str, use_default_template: bool) -> None:
        """Processa cada caso de teste individualmente"""
        self.log(f"📊 Encontrados {len(casos_teste)} casos de teste\n")
        self.progress['maximum'] = len(casos_teste)
        
        sucessos = 0
        erros = []
        arquivos_gerados = set()
        
        # Criar mapeamento de campos
        field_mapping = {campo['key']: campo['label'].rstrip(':').strip() 
                        for campo in self.campos_config}
        
        campo_nome = next(iter(dados_fixos.keys()))
        
        for i, caso_teste in enumerate(casos_teste, 1):
            try:
                self.progress['value'] = i
                nome_caso = caso_teste.get('Nome', f'Caso_{i}')
                self.log(f"🔄 Processando: {nome_caso}")
                
                if self._generate_single_document(caso_teste, dados_fixos, template_path, 
                                                output_folder, field_mapping, campo_nome, 
                                                arquivos_gerados, use_default_template):
                    sucessos += 1
                else:
                    erros.append((nome_caso, "Erro na geração"))
                    
            except Exception as e:
                nome_caso = caso_teste.get('Nome', f'Caso_{i}')
                self.log(f"❌ Erro no caso '{nome_caso}': {e}\n")
                erros.append((nome_caso, str(e)))
            
            time.sleep(0.05)  # Pequena pausa para não sobrecarregar
            
        self._show_final_results(sucessos, len(erros), len(casos_teste), 
                               output_folder, arquivos_gerados)
        self.gerar_btn.config(state='normal')

    def _generate_single_document(self, caso_teste: Dict, dados_fixos: Dict[str, str], 
                                template_path: str, output_folder: str, 
                                field_mapping: Dict[str, str], campo_nome: str,
                                arquivos_gerados: set, use_default_template: bool) -> bool:
        """Gera um único documento - versão robusta com suporte a BDD"""
        try:
            nome_caso = caso_teste.get('Nome', 'Caso_Desconhecido')
            
            # Combinar dados fixos com dados do CSV - CORREÇÃO CRÍTICA
            dados_completos = dados_fixos.copy()  # Começa com dados da interface
            
            # Adicionar TODOS os dados do CSV (sobrescrevendo se necessário)
            for key, value in caso_teste.items():
                if value and value != 'N/A':  # Ignorar campos vazios ou "N/A"
                    dados_completos[key] = value
            
            # Garantir que o nome do caso de teste está correto
            dados_completos['Caso de Teste'] = nome_caso
            
            # Log dos dados BDD para debug
            bdd_fields = ['Given', 'When', 'Then', 'And', 'Script de Teste (BDD)']
            bdd_dados = {k: v for k, v in caso_teste.items() if k in bdd_fields and v}
            
            if bdd_dados:
                self.log(f"   📋 Dados BDD encontrados: {list(bdd_dados.keys())}")
                for campo, valor in bdd_dados.items():
                    if campo == 'Script de Teste (BDD)':
                        self.log(f"      {campo}: {valor[:100]}{'...' if len(valor) > 100 else ''}")
                    else:
                        self.log(f"      {campo}: {valor}")
            
            # Usar template se fornecido e existir, caso contrário criar documento padrão
            if not use_default_template:
                try:
                    doc = Document(template_path)
                    # Tenta preencher o template e verifica se algum campo BDD foi preenchido
                    bdd_was_filled = self.doc_processor.fill_template(doc, dados_completos, field_mapping)
                    self.log("   📝 Template personalizado preenchido")
                    
                    # SE HÁ DADOS BDD MAS NENHUM CAMPO BDD FOI PREENCHIDO NO TEMPLATE, ADICIONA SEÇÃO BDD
                    if bdd_dados and not bdd_was_filled:
                        self.log("   ➕ Adicionando seção BDD ao template personalizado")
                        self.doc_processor.add_bdd_section_to_template(doc, dados_completos)
                    
                except Exception as e:
                    self.log(f"⚠️ Erro ao usar template personalizado: {e}. Usando template padrão...")
                    doc = self.default_doc_generator.create_default_document(
                        dados_completos, self.campos_config)
            else:
                # Criar documento padrão com todos os dados
                doc = self.default_doc_generator.create_default_document(
                    dados_completos, self.campos_config)
                self.log("   📝 Documento padrão gerado")
            
            # Gerar nome do arquivo
            nome_base = self.doc_processor.clean_filename(nome_caso)
            nome_arquivo = self._generate_unique_filename(
                f"Evidencia_{dados_fixos[campo_nome]}_{nome_base}.docx", arquivos_gerados)
            
            caminho_completo = Path(output_folder) / nome_arquivo
            
            # Tentar salvar o documento
            try:
                doc.save(caminho_completo)
                self.log(f"✅ Salvo: {nome_arquivo}")
                
                # Log de campos BDD processados se existirem
                if bdd_dados:
                    self.log(f"   📋 Cenário BDD incluído no documento")
                    
                return True
            except Exception as e:
                # Fallback: tentar salvar com nome diferente
                try:
                    nome_alternativo = f"Evidencia_{len(arquivos_gerados)}_{datetime.now().strftime('%H%M%S')}.docx"
                    caminho_alternativo = Path(output_folder) / nome_alternativo
                    doc.save(caminho_alternativo)
                    self.log(f"✅ Salvo (nome alternativo): {nome_alternativo}")
                    arquivos_gerados.add(nome_alternativo)
                    return True
                except Exception as e2:
                    self.log(f"❌ Erro ao salvar documento: {e2}")
                    return False
            
        except Exception as e:
            self.log(f"❌ Erro crítico ao gerar documento: {e}")
            return False

    def _generate_unique_filename(self, filename: str, existing_files: set) -> str:
        """Gera um nome de arquivo único"""
        contador = 1
        nome_original = filename
        
        while filename in existing_files:
            nome, extensao = os.path.splitext(nome_original)
            filename = f"{nome}_{contador}{extensao}"
            contador += 1
        
        existing_files.add(filename)
        return filename

    def _show_final_results(self, sucessos: int, erros: int, total: int, 
                           output_folder: str, arquivos_gerados: set) -> None:
        """Mostra os resultados finais do processamento"""
        self.log("\n" + "="*50)
        self.log("📋 RESUMO DA EXECUÇÃO")
        self.log("="*50)
        self.log(f"✅ Documentos gerados com sucesso: {sucessos}")
        self.log(f"❌ Documentos com erro: {erros}")
        self.log(f"📊 Total processado: {total}")
        
        if sucessos > 0:
            self.log(f"📁 Pasta de saída: {output_folder}")
            self.log(f"📄 Arquivos gerados: {len(arquivos_gerados)}")
            
            if messagebox.askyesno("Concluído", 
                                 f"Processamento concluído!\n"
                                 f"Sucessos: {sucessos}\n"
                                 f"Erros: {erros}\n\n"
                                 f"Deseja abrir a pasta de saída?"):
                try:
                    os.startfile(output_folder)
                except:
                    self.log("⚠️ Não foi possível abrir a pasta automaticamente")
        else:
            messagebox.showwarning("Atenção", "Nenhum documento foi gerado com sucesso!")

    def iniciar_processamento(self) -> None:
        """Inicia o processamento em thread separada"""
        self.gerar_btn.config(state='disabled')
        self._clear_log()
        
        thread = threading.Thread(target=self.processar_documentos, daemon=True)
        thread.start()


# Função de compatibilidade para manter a interface antiga
def create_template_generator(parent, settings=None):
    """Função de fábrica para criar o módulo"""
    return TemplateGeneratorModule(parent, settings)


# Teste local do módulo
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Teste Template Generator")
    root.geometry("800x600")
    
    app = TemplateGeneratorModule(root)
    app.show()
    
    root.mainloop()