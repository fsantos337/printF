# PrintF v1.0 - Sistema Unificado de Captura de Evidências

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-green)
![License](https://img.shields.io/badge/license-MIT-orange)

## 📋 Índice

- [Sobre o Projeto](#-sobre-o-projeto)
- [Características Principais](#-características-principais)
- [Instalação](#-instalação)
- [Guia de Uso](#-guia-de-uso)
- [Módulos](#-módulos)
- [Compilação](#-compilação)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Atalhos de Teclado](#%EF%B8%8F-atalhos-de-teclado)
- [Solução de Problemas](#-solução-de-problemas)
- [Licença](#-licença)

## 🎯 Sobre o Projeto

O **PrintF** é uma solução profissional e integrada para captura, organização e documentação de evidências em processos de teste de software. A versão 1.0 representa uma evolução significativa, unificando todos os módulos em uma única aplicação com interface moderna e responsiva.

### 🌟 Novidades da Versão 1.0

- ✅ **Aplicação Unificada** - Todos os módulos integrados em um único executável
- 🎨 **Interface Moderna** - Tema visual "Liquid Glass" com efeitos translúcidos
- 📱 **Design Responsivo** - Adapta-se automaticamente ao tamanho da tela
- 🏗️ **Arquitetura Modular** - Código organizado, escalável e de fácil manutenção
- ⚡ **Performance Otimizada** - Sistema leve e eficiente
- 🔧 **Configuração Flexível** - Sistema de configurações persistentes

## ✨ Características Principais

### 🎨 Interface do Usuário

- **Tema Liquid Glass**: Visual moderno com efeitos translúcidos e animações suaves
- **Layout Responsivo**: Adapta-se automaticamente a diferentes resoluções (mobile, tablet, desktop)
- **Navegação Intuitiva**: Cards interativos com feedback visual
- **Sistema de Temas**: Suporte para múltiplos temas visuais
- **Atalhos Globais**: Acesso rápido a todas as funcionalidades

### 📷 Sistema de Captura

- **Multi-Monitor**: Suporte completo para configurações com múltiplos monitores
- **Timestamp Automático**: Todas as capturas incluem data/hora
- **Marcador de Clique**: Destaque visual do ponto de interação
- **Metadados Completos**: Informações técnicas automáticas (resolução, sistema, etc)
- **Modos de Captura**: "Ocultar" ou "Manter" barra de tarefas
- **Controles Flexíveis**: Pausar, retomar e finalizar gravações

### 📄 Geração de Documentos

- **Templates DOCX Personalizáveis**: Crie seus próprios modelos
- **Geração em Lote**: Processe múltiplos documentos simultaneamente
- **Campos Dinâmicos**: Preenchimento automático de variáveis
- **Backup de Metadados**: Preservação de informações importantes
- **Editor Integrado**: Adicione comentários e observações
- **Navegação Avançada**: Percorra evidências com facilidade

### 🗑️ Gestão de Arquivos

- **Análise de Disco**: Visualize o uso de espaço detalhadamente
- **Filtros Inteligentes**: Organize por tipo, tamanho e data
- **Exclusão Segura**: Confirmação antes de remover arquivos
- **Backup Automático**: Proteção contra perda de dados

## 🚀 Instalação

### Opção 1: Executável Pronto (Recomendado)

1. Baixe o arquivo `PrintF.exe` da [página de releases](https://github.com/usuario/printf/releases)
2. Execute o arquivo (não requer instalação ou permissões especiais)
3. Pronto! Todos os módulos estão incluídos e prontos para uso

### Opção 2: Executar via Código Fonte

**Pré-requisitos:**
- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

**Passo a passo:**

```bash
# 1. Clone o repositório ou baixe os arquivos
git clone https://github.com/usuario/printf.git
cd printf

# 2. Instale as dependências
pip install -r requirements.txt

# Ou instale manualmente:
pip install pillow python-docx pyautogui pynput psutil screeninfo

# 3. Execute a aplicação
python main.py
```

### Dependências Necessárias

```
pillow>=8.0.0          # Processamento de imagens
python-docx>=0.8.11    # Geração de documentos DOCX
pyautogui>=0.9.53      # Automação de interface
pynput>=1.7.6          # Captura de eventos de teclado/mouse
psutil>=5.8.0          # Informações do sistema
screeninfo>=0.8        # Detecção de monitores
```

## 📖 Guia de Uso

### 🏠 Tela Principal

Ao iniciar o PrintF, você encontrará quatro módulos principais organizados em cards interativos:

1. **📷 Capturar Evidências (F8)** - Sistema de captura de screenshots
2. **📄 Gerar Templates (F9)** - Criação de templates de documentos
3. **📋 Gerar Documentos (F10)** - Geração de documentação completa
4. **🗑️ Limpar Arquivos (F11)** - Gestão e limpeza de arquivos

### 🔄 Fluxo de Trabalho Básico

```
1. Iniciar PrintF
   ↓
2. Selecionar módulo (clique ou atalho F8-F11)
   ↓
3. Executar ações do módulo
   ↓
4. Resultados salvos automaticamente em Output/
```

### ⚙️ Configuração Automática

Na primeira execução, o sistema cria automaticamente:

```
PrintF/
├── 📁 Output/             # Evidências capturadas
├── 📁 Templates/          # Templates de documentos
├── 📁 Config/             # Configurações do usuário
├── 📁 Logs/              # Logs de execução
└── 📄 user_settings.json  # Preferências salvas
```

## 🏗️ Módulos

### 📷 Capturar Evidências (F8)

**Objetivo:** Captura sistemática e organizada de screenshots durante testes.

**Recursos:**
- Captura com um clique do mouse
- Detecção automática de múltiplos monitores
- Timestamp e metadados em cada captura
- Marcador visual do ponto clicado
- Dois modos de operação:
  - **Ocultar**: Remove barra de tarefas das capturas
  - **Manter**: Preserva a interface completa

**Como usar:**
1. Pressione `F8` para iniciar a gravação
2. Clique em qualquer lugar da tela para capturar
3. Use `F6` para pausar/retomar
4. Pressione `F9` para finalizar e salvar

**Arquivos gerados:**
- Imagens PNG com timestamp
- Arquivo JSON com metadados
- Log de sessão de captura

### 📄 Gerar Templates (F9)

**Objetivo:** Criação de templates DOCX personalizados para documentação.

**Recursos:**
- Templates com campos dinâmicos
- Suporte a arquivos CSV para dados em lote
- Campos personalizáveis: `[NOME_CAMPO]`
- Preservação de formatação original

**Como usar:**
1. Prepare um arquivo CSV com coluna 'Nome'
2. Crie ou edite um template DOCX
3. Insira campos dinâmicos: `[PROJETO]`, `[MÓDULO]`, etc
4. Execute o módulo para gerar documentos

**Exemplo de campos:**
```
[NOME_DO_PROJETO]
[MÓDULO]
[VERSÃO]
[RESPONSÁVEL]
[DATA]
[AMBIENTE]
```

### 📋 Gerar Documentos (F10)

**Objetivo:** Transformar evidências capturadas em documentação profissional.

**Recursos:**
- Navegação interativa entre evidências
- Editor integrado para comentários
- Adição de observações e anotações
- Geração de DOCX final formatado
- Inclusão automática de metadados

**Como usar:**
1. Selecione a pasta com evidências
2. Navegue pelas capturas usando os controles
3. Adicione comentários e observações
4. Clique em "Gerar Documento" para criar o DOCX

**Estrutura do documento gerado:**
- Cabeçalho com informações do projeto
- Seção de detalhes do teste
- Evidências com timestamps
- Comentários e observações
- Metadados técnicos

### 🗑️ Limpar Arquivos (F11)

**Objetivo:** Gerenciamento e organização eficiente do espaço em disco.

**Recursos:**
- Análise detalhada de uso de disco
- Filtros por tipo, tamanho e data
- Exclusão segura com confirmação
- Proteção contra remoção acidental
- Estatísticas de limpeza

**Como usar:**
1. Selecione a pasta para análise
2. Revise os arquivos listados
3. Aplique filtros conforme necessário
4. Marque itens para exclusão
5. Confirme a operação

## 🔨 Compilação

### Gerar Executável com PyInstaller

#### Método 1: Comando Direto (Rápido)

```bash
# 1. Instalar PyInstaller
pip install pyinstaller

# 2. Compilar aplicação
pyinstaller --onefile --windowed --name "PrintF" \
  --add-data "modules;modules" \
  --add-data "config.py;." \
  --hidden-import=modules.capture \
  --hidden-import=modules.template_gen \
  --hidden-import=modules.evidence_gen \
  --hidden-import=modules.cleanup \
  --hidden-import=modules.styles \
  --hidden-import=PIL._tkinter_finder \
  main.py
```

#### Método 2: Script Automatizado (Windows)

Execute o arquivo `gerarEXE.bat` incluído no projeto:

```batch
gerarEXE.bat
```

#### Método 3: Configuração Personalizada

Crie um arquivo `build.spec`:

```python
# -*- mode: python ; coding: utf-8 -*-
import os

a = Analysis(
    ['main.py'],
    pathex=[os.getcwd()],
    binaries=[],
    datas=[
        ('modules/*.py', 'modules'),
        ('config.py', '.'),
        ('assets/icon.ico', 'assets')
    ],
    hiddenimports=[
        'modules.capture',
        'modules.template_gen', 
        'modules.evidence_gen',
        'modules.cleanup',
        'modules.styles',
        'PIL._tkinter_finder',
        'docx',
        'pyautogui',
        'pynput.keyboard',
        'pynput.mouse'
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PrintF',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='assets/icon.ico'
)
```

Execute:
```bash
pyinstaller build.spec
```

### 📦 Localização do Executável

Após compilação bem-sucedida:

```
dist/
└── PrintF.exe  ← SEU EXECUTÁVEL AQUI
```

## 📁 Estrutura do Projeto

```
PrintF/
│
├── 🐍 main.py                  # Aplicação principal e orquestração
├── ⚙️ config.py                # Configurações globais do sistema
├── 📄 requirements.txt         # Dependências do projeto
├── 📄 README.md               # Esta documentação
├── 🔨 gerarEXE.bat            # Script de compilação (Windows)
│
├── 📁 modules/                 # Módulos da aplicação
│   ├── 🎯 capture.py          # Sistema de captura de evidências
│   ├── 📑 template_gen.py     # Gerador de templates DOCX
│   ├── 📄 evidence_gen.py     # Gerador de documentos finais
│   ├── 🗑️ cleanup.py          # Gerenciador de limpeza
│   └── 🎨 styles.py           # Sistema de temas visuais
│
├── 📁 assets/                  # Recursos visuais (opcional)
│   └── icon.ico               # Ícone da aplicação
│
├── 📁 Output/                  # ✅ Criado automaticamente
│   └── [evidências]           # Screenshots e metadados
│
├── 📁 Templates/               # ✅ Criado automaticamente
│   └── template_evidencias.docx
│
├── 📁 Config/                  # ✅ Criado automaticamente
│   └── user_settings.json     # Preferências do usuário
│
└── 📁 Logs/                    # ✅ Criado automaticamente
    └── printf_[DATA].log       # Logs de execução
```

## ⌨️ Atalhos de Teclado

### Atalhos Globais

| Tecla | Função |
|-------|--------|
| `F8` | Abrir módulo Capturar Evidências |
| `F9` | Abrir módulo Gerar Templates |
| `F10` | Abrir módulo Gerar Documentos |
| `F11` | Abrir módulo Limpar Arquivos |
| `F12` | Fechar aplicação |

### Módulo de Captura

| Tecla | Função |
|-------|--------|
| `F8` | Iniciar gravação |
| `F6` | Pausar/Retomar gravação |
| `F9` | Finalizar e salvar gravação |
| `Mouse Click` | Capturar screenshot |

## 🔧 Solução de Problemas

### ❌ Executável não é gerado

**Possíveis causas:**
- Falta de permissões de administrador
- Espaço insuficiente em disco
- PyInstaller não instalado corretamente

**Soluções:**
```bash
# 1. Executar terminal como administrador
# 2. Verificar espaço em disco
# 3. Reinstalar PyInstaller
pip uninstall pyinstaller
pip install pyinstaller --upgrade
```

### ❌ "ModuleNotFoundError" ao executar

**Solução:** Adicione os módulos faltantes ao `hiddenimports` no comando PyInstaller:

```python
--hidden-import=nome_do_modulo_faltante
```

### ❌ Interface não carrega ou apresenta erros

**Diagnóstico:**
```bash
# Execute via terminal para ver erros detalhados
PrintF.exe

# Ou no código fonte
python main.py
```

**Verifique:**
- Dependências instaladas corretamente
- Logs em `Logs/printf_[DATA].log`
- Permissões de escrita nas pastas

### ❌ Captura de tela não funciona (Linux)

**Solução:**
```bash
# Instalar dependências do sistema
sudo apt-get update
sudo apt-get install python3-tk python3-dev scrot

# Instalar dependências Python
pip install python-xlib
```

### ❌ Tema Liquid Glass não aparece

**Verificações:**
1. Confirme que `modules/styles.py` existe
2. Verifique o arquivo `Config/user_settings.json`:
```json
{
  "theme": "liquid_glass"
}
```
3. Reinicie a aplicação

### 📊 Logs de Diagnóstico

Os logs detalhados estão disponíveis em:
```
Logs/printf_[DATA_HORA].log
```

Eles incluem:
- Erros de execução
- Avisos de configuração
- Operações realizadas
- Performance do sistema

## 📄 Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

```
MIT License

Copyright (c) 2024 PrintF Team

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 👥 Autores

- **Fernanda Maria dos Santos Braga** - [GitHub](https://github.com/fernanda)
- **Thiago Gomes Rocha** - [GitHub](https://github.com/thiago)

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abra um Pull Request

## 📞 Suporte

- 🐛 [Reportar Bug](https://github.com/usuario/printf/issues)
- 💡 [Solicitar Feature](https://github.com/usuario/printf/issues)
- 💬 [Discussões](https://github.com/usuario/printf/discussions)

## 🗺️ Roadmap

- [ ] Suporte para vídeos de evidência
- [ ] Integração com ferramentas de gestão de testes
- [ ] Exportação para PDF
- [ ] Modo escuro/claro alternável
- [ ] Tradução para outros idiomas
- [ ] API REST para integração
- [ ] Plugin para navegadores

---

<p align="center">
  <strong>Desenvolvido com ❤️ para a comunidade de testes de software</strong>
</p>

<p align="center">
  <a href="#-índice">⬆️ Voltar ao topo</a>
</p>