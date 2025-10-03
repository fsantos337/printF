# PrintF v1.0 - Sistema Unificado de Captura de Evidências

![Version](https://img.shields.io/badge/version-2.0.0-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-green)
![License](https://img.shields.io/badge/license-MIT-orange)

## 📋 Sumário

- [Sobre](#sobre)
- [Características](#características)
- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Uso](#uso)
- [Compilação](#compilação)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Atalhos de Teclado](#atalhos-de-teclado)
- [Solução de Problemas](#solução-de-problemas)
- [Contribuindo](#contribuindo)
- [Changelog](#changelog)
- [Licença](#licença)

## 🎯 Sobre

PrintF é um sistema profissional para captura e documentação de evidências de testes de software. Desenvolvido para simplificar o processo de documentação, permite capturar screenshots com apenas um clique, adicionar anotações e gerar relatórios automatizados em formato DOCX.

### Por que PrintF v2.0?

A versão 2.0 representa uma reformulação completa do sistema:
- **Aplicativo unificado** - Todas as funcionalidades em um único executável
- **Arquitetura modular** - Código organizado e manutenível
- **Performance otimizada** - Captura mais rápida e eficiente
- **Interface moderna** - Design intuitivo e responsivo
- **Melhor tratamento de erros** - Sistema mais robusto e confiável

## ✨ Características

### Principais Funcionalidades

- 📷 **Captura Inteligente**: Suporte multi-monitor com detecção automática
- 🎨 **Editor Integrado**: Adicione anotações, setas e destaques nas evidências
- 📝 **Geração Automática**: Crie documentos DOCX formatados automaticamente
- 🕒 **Timestamp Automático**: Data/hora adicionada automaticamente nas capturas
- 🎯 **Marcador de Clique**: Destaque visual do ponto clicado
- 💾 **Metadados Inteligentes**: Rastreamento completo de todas as evidências
- 🔒 **Instância Única**: Previne execuções múltiplas acidentais

### Modos de Captura

1. **Ocultar Barra de Tarefas**: Captura apenas a área de trabalho, ideal para documentação limpa
2. **Manter Barra de Tarefas**: Captura tela completa, preservando o contexto do sistema

## 📦 Requisitos

### Sistema Operacional
- Windows 10/11 (64-bit)
- Linux (Ubuntu 20.04+)
- macOS 10.15+

### Python e Dependências

```bash
Python 3.8 ou superior
```

#### Dependências Principais:
```txt
# requirements.txt
tkinter              # Interface gráfica (geralmente incluído no Python)
Pillow>=9.0.0       # Manipulação de imagens
python-docx>=0.8.11 # Geração de documentos
pynput>=1.7.0       # Captura de eventos de mouse/teclado
pyautogui>=0.9.53   # Automação e captura de tela
psutil>=5.9.0       # Gerenciamento de processos

# Dependências Opcionais (melhor performance)
mss>=7.0.0          # Captura de tela otimizada
pywin32>=304        # APIs Windows (apenas Windows)
screeninfo>=0.8     # Informações de múltiplos monitores
```

## 🚀 Instalação

### Método 1: Instalação Rápida (Executável)

1. Baixe o executável mais recente em [Releases](https://github.com/seu-usuario/printf/releases)
2. Execute `PrintF.exe`
3. Pronto! Não requer instalação

### Método 2: Instalação do Código-Fonte

```bash
# 1. Clone o repositório
git clone https://github.com/seu-usuario/printf.git
cd printf

# 2. Crie um ambiente virtual
python -m venv venv

# 3. Ative o ambiente virtual
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

# 4. Instale as dependências
pip install -r requirements.txt

# 5. Execute o aplicativo
python printF_v2.py
```

### Método 3: Instalação com Script

```bash
# Windows
install.bat

# Linux/macOS
chmod +x install.sh
./install.sh
```

## ⚙️ Configuração

### Primeira Execução

Na primeira execução, o PrintF criará automaticamente:
- Diretório de configuração: `~/.printf/`
- Arquivo de configuração: `~/.printf/config.json`
- Diretório padrão de saída: `~/Documents/PrintF/`

### Arquivo de Configuração

```json
{
  "version": "2.0.0",
  "output_dir": "~/Documents/PrintF",
  "template_dir": "~/Documents/PrintF/Templates",
  "capture_mode": "ocultar",
  "keep_evidence_files": true,
  "timestamp_position": [0.75, 0.90],
  "timestamp_color": "#FFFFFF",
  "timestamp_background": "#000000B2",
  "timestamp_size": 24
}
```

### Personalização

#### Alterar Diretório de Saída
1. Menu `Arquivo` → `Configurações`
2. Altere o campo "Diretório de Saída"
3. Clique em `Salvar`

#### Configurar Modo de Captura
1. Na tela inicial, clique em `Iniciar Gravação`
2. Selecione o modo desejado:
   - **Ocultar barra**: Para documentação limpa
   - **Manter barra**: Para contexto completo

## 📖 Uso

### Workflow Básico

1. **Iniciar o PrintF**
   ```bash
   python printF_v2.py
   # ou
   PrintF.exe
   ```

2. **Configurar Sessão**
   - Clique em `Iniciar Gravação (F8)`
   - Selecione o template (opcional)
   - Escolha o diretório de saída
   - Configure o modo de captura

3. **Capturar Evidências**
   - Clique em qualquer lugar da tela para capturar
   - Use `F6` para pausar temporariamente
   - Use `F7` para retomar

4. **Finalizar e Editar**
   - Pressione `F9` para finalizar
   - Revise e edite as evidências
   - Adicione comentários

5. **Gerar Documento**
   - Clique em `Gerar Documento`
   - O DOCX será salvo automaticamente

### Funcionalidades Avançadas

#### Editor de Evidências
- **Adicionar Setas**: Destaque elementos importantes
- **Inserir Texto**: Adicione anotações diretamente na imagem
- **Desenhar Retângulos**: Enquadre áreas específicas
- **Círculos de Destaque**: Realce pontos de interesse

#### Templates Personalizados
1. Crie um documento DOCX com sua formatação
2. Use como template para manter consistência
3. O PrintF preservará estilos e formatação

## 🔨 Compilação

### Criar Executável Único

#### Usando PyInstaller

```bash
# Instalar PyInstaller
pip install pyinstaller

# Compilar aplicativo
pyinstaller --onefile \
            --windowed \
            --name="PrintF" \
            --icon="assets/icon.ico" \
            --add-data="assets;assets" \
            --hidden-import="PIL._tkinter_finder" \
            printF_v2.py

# O executável estará em: dist/PrintF.exe
```

#### Script de Build Automatizado

```bash
# Windows
build.bat

# Linux/macOS
chmod +x build.sh
./build.sh
```

### Build com Configurações Avançadas

```python
# build_config.spec
a = Analysis(
    ['printF_v2.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        ('templates', 'templates')
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'pynput.keyboard._win32',
        'pynput.mouse._win32'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'scipy'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PrintF',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico',
)
```

Execute com:
```bash
pyinstaller build_config.spec
```

## 📁 Estrutura do Projeto

```
printf/
│
├── printF_v2.py          # Aplicativo principal unificado
├── README.md             # Documentação
├── requirements.txt      # Dependências Python
├── LICENSE              # Licença do projeto
│
├── assets/              # Recursos visuais
│   ├── icon.ico        # Ícone do aplicativo
│   └── screenshots/    # Capturas de tela para documentação
│
├── templates/           # Templates DOCX
│   ├── default.docx    # Template padrão
│   └── custom/         # Templates personalizados
│
├── scripts/            # Scripts auxiliares
│   ├── build.bat      # Script de build Windows
│   ├── build.sh       # Script de build Linux/macOS
│   ├── install.bat    # Instalador Windows
│   └── install.sh     # Instalador Linux/macOS
│
└── dist/              # Executáveis compilados
    └── PrintF.exe     # Aplicativo compilado
```

## ⌨️ Atalhos de Teclado

| Atalho | Função |
|--------|--------|
| `F8` | Iniciar gravação |
| `F6` | Pausar gravação |
| `F7` | Retomar gravação |
| `F9` | Finalizar gravação |
| `F12` | Fechar aplicativo |
| `Ctrl+N` | Nova sessão |
| `Ctrl+Q` | Sair |
| `Ctrl+Z` | Desfazer (no editor) |

## 🔧 Solução de Problemas

### Problema: "PrintF já está em execução"
**Solução:**
```bash
# Windows: Abrir Gerenciador de Tarefas e finalizar PrintF.exe
# Linux/macOS:
pkill -f printF_v2.py
# Ou remover arquivo de lock:
rm ~/.printf/printf.lock
```

### Problema: Erro de importação de módulos
**Solução:**
```bash
# Reinstalar dependências
pip install --upgrade -r requirements.txt
```

### Problema: Captura de tela preta ou incorreta
**Solução:**
1. Instale a biblioteca `mss` para melhor suporte:
   ```bash
   pip install mss
   ```
2. No Windows, instale `pywin32`:
   ```bash
   pip install pywin32
   ```

### Problema: Documento DOCX não é gerado
**Solução:**
```bash
# Verificar instalação do python-docx
pip install --upgrade python-docx
```

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add: nova funcionalidade'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

### Diretrizes de Código

- Siga PEP 8
- Adicione docstrings
- Escreva testes quando aplicável
- Mantenha a compatibilidade com Python 3.8+

## 📝 Changelog

### v2.0.0 (2024-01-XX)
- 🎉 Reformulação completa do sistema
- ✨ Aplicativo unificado (single executable)
- 🏗️ Arquitetura modular com classes bem definidas
- 🚀 Performance otimizada na captura
- 🛡️ Melhor tratamento de erros
- 📚 Documentação completa
- 🔧 Sistema de configuração aprimorado
- 🎨 Interface modernizada

### v1.0.0 (2024-01-XX)
- Versão inicial
- Múltiplos executáveis
- Funcionalidades básicas

## 📄 Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 👥 Autores

- **Fernanda Maria dos Santos Braga** - *Desenvolvimento Principal* - [GitHub](https://github.com/fsantos337)
- **Thiago Gomes Rocha** - *Desenvolvimento Principal* - [GitHub](https://github.com/thiagogomesrocha)

## 🙏 Agradecimentos

- Comunidade Python
- Contribuidores do projeto
- Usuários que reportaram bugs e sugeriram melhorias

---

<p align="center">
  Desenvolvido com ❤️ para facilitar a documentação de testes
</p>

<p align="center">
  <a href="https://github.com/seu-usuario/printf/issues">Reportar Bug</a> •
  <a href="https://github.com/seu-usuario/printf/issues">Solicitar Feature</a>
</p>
