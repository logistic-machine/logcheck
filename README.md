# LogCheck 📋🚀

> Automação de auditorias logísticas com captura de evidências, geração de relatórios PDF e integração Outlook.

<img width="472" height="354" alt="Login" src="https://github.com/user-attachments/assets/3c59202d-3dd4-4a28-8d99-83395b9487b6" />
<img width="1021" height="817" alt="Painel Master" src="https://github.com/user-attachments/assets/c13adb6c-b9ce-4a2b-8c32-48a35368f316" />
<img width="987" height="870" alt="Check-list" src="https://github.com/user-attachments/assets/c4e978cc-8988-492e-a1ec-edcfebb3bc81" />


## Visão Geral

O **LogCheck** é uma ferramenta desktop desenvolvida para **automatizar o processo de inventário de bases logísticas**, substituindo planilhas Excel manuais por um sistema estruturado, seguro e rastreável.

### 📊 Impacto

| Métrica | Antes | Depois | Redução |
|---------|-------|--------|---------|
| **Tempo por auditoria** | 3-4 horas | 45 minutos | **-80%** |
| **Erros de transcrição** | ~5-10% | 0% | **100%** |
| **Rastreabilidade** | Nenhuma | Completa | ✅ |
| **Evidências fotográficas** | Pasta separada | Integrada no PDF | ✅ |

---

## ✨ Principais Funcionalidades

### 1️⃣ Autenticação Segura
- Criptografia PBKDF2-HMAC-SHA256 com salt aleatório
- Sistema de perfis (Master/Auditor)
- Controle de acesso granular
- Gerenciamento de senhas padrão

### 2️⃣ Formulário Inteligente
- **Listas suspensas em cascata**: Área → Centro → Empresa → Localidade
- **Perguntas estruturadas**: SIM / NÃO / N/A
- **Upload obrigatório de PNG** quando resposta = NÃO
- **Placar live**: Cálculo automático de pontuação conforme preenchimento

### 3️⃣ Processamento Automático
- Cálculo de notas por seção (regra: SIM = peso da pergunta; NÃO/N/A = 0)
- Cálculo de aproveitamento (%)
- Geração de **PDF profissional** com evidências integradas
- Fórmulas Excel dinâmicas (VLOOKUP, SUM)

### 4️⃣ Integração Corporativa
- Envio automático de relatórios via **Outlook** (pywin32)
- Pergunta SIM/NÃO antes de enviar
- Tratamento de erros robusto
- Armazenamento centralizado em Excel

### 5️⃣ Deploy Facilitado
- Compatível com **PyInstaller** (executável .exe standalone)
- Sem dependências externas no ambiente
- Funciona offline
- Caminhos relativos para portabilidade

---

## 🛠️ Stack Técnico

```
┌─────────────────────────────────────────────────────────┐
│ Frontend (Tkinter)                                      │
│ └─ Interface responsiva para campo / desktop           │
├─────────────────────────────────────────────────────────┤
│ Backend (Python)                                        │
│ ├─ Autenticação (PBKDF2-HMAC-SHA256)                   │
│ ├─ Processamento de dados (pontuar, validar)           │
│ └─ Integração com serviços (Outlook, PDF)              │
├─────────────────────────────────────────────────────────┤
│ Armazenamento (Excel + Arquivos)                        │
│ ├─ openpyxl (leitura/escrita com fórmulas)            │
│ ├─ PNG (evidências fotográficas)                       │
│ └─ PDF (relatórios)                                    │
├─────────────────────────────────────────────────────────┤
│ Integração (Outlook/COM)                               │
│ └─ pywin32 (envio de e-mail corporativo)              │
└─────────────────────────────────────────────────────────┘
```

### Pacotes Principais

```python
# Interface & Desktop
tkinter              # UI (stdlib)
Pillow (PIL)         # Processamento de imagens

# Dados & Persistência
openpyxl             # Excel com fórmulas dinâmicas
pathlib              # Gerenciamento de caminhos (stdlib)

# PDF & Relatórios
ReportLab            # Geração de PDF com layout profissional
reportlab.platypus   # Componentes de página

# Segurança & Autenticação
hashlib              # PBKDF2-HMAC-SHA256 (stdlib)
hmac                 # Verificação segura de hashes (stdlib)
os                   # Geração de salt aleatório (stdlib)

# Integração Corporativa
pywin32              # COM com Outlook (Windows)

# Deploy
PyInstaller          # Empacotamento em .exe
```

---

## 📁 Estrutura do Projeto

```
inventario/
├── src/
│   ├── __main__.py          # Entry point
│   ├── app.py               # Interface Tkinter (Etapas 2-4)
│   ├── auth.py              # Autenticação & perfis
│   ├── config.py            # Caminhos & constantes
│   ├── data_loader.py       # Carregamento de dados (Etapa 1)
│   ├── email_utils.py       # Integração Outlook
│   └── results.py           # Cálculo & gravação (Etapas 3-4)
├── data/
│   ├── Cronograma_Atualizado.xlsx   # Parâmetros (Ação 1.2)
│   └── Checklist_Avaliacao.xlsx     # Perguntas (Ação 2.3)
├── resultados/
│   ├── Inventario_Resultados.xlsx   # BD de auditorias
│   ├── Relatorios_PDF/              # PDFs gerados
│   ├── evidencias/                  # PNGs por auditoria
│   └── usuarios.json                # Base de usuários
└── requirements.txt
```

---

## 🚀 Instalação

### Pré-requisitos
- Python 3.9+
- Outlook instalado e logado (para envio de e-mail)
- Windows 7+ (Tkinter + pywin32)

### Setup Desenvolvimento

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/logcheck.git
cd logcheck

# Crie um ambiente virtual
python -m venv venv
venv\Scripts\activate  # Windows

# Instale as dependências
pip install -r requirements.txt

# Execute o app
python -m inventario.src.app
```

### Deploy (PyInstaller)

```bash
# Instale PyInstaller
pip install pyinstaller

# Gere o executável
pyinstaller --name=LogCheck \
  --onefile \
  --windowed \
  --add-data "inventario/data:data" \
  inventario/src/__main__.py

# O .exe fica em: dist/LogCheck.exe
```

---

## 💻 Como Usar

### 1️⃣ Login
```
E-mail: seu_email@empresa.com
Senha: (padrão fornecido pelo Master)
```

### 2️⃣ Preenchimento do Checklist
1. Preencha os campos obrigatórios (Responsável, Data, Inspetor)
2. Selecione os parâmetros em cascata (Área → Centro → Empresa → Localidade)
3. Responda as perguntas (SIM / NÃO / N/A)
4. **Quando resposta = NÃO**: Anexe evidência PNG + justificativa

### 3️⃣ Salvamento
1. Clique em **"Salvar"**
2. Responda SIM/NÃO para enviar relatório por e-mail
3. Checklist salva em `Inventario_Resultados.xlsx`
4. PDF gerado automaticamente em `Relatorios_PDF/`

### 4️⃣ Painel Master (se tiver permissão)
- Cadastro de novos auditores
- Liberar/bloquear acesso
- Conceder/retirar permissão Master
- Resetar senhas

---

## 🔐 Segurança

### Autenticação
- **Algoritmo**: PBKDF2-HMAC-SHA256
- **Iterações**: 200.000 (OWASP recomenda 100k+)
- **Salt**: 16 bytes aleatórios por usuário
- **Comparação**: `hmac.compare_digest()` (protege contra timing attacks)

### Controle de Acesso
```python
usuario.ativo    # SIM/NÃO: pode abrir o checklist?
usuario.master   # Pode cadastrar outros? Sempre ativo.
```

### Armazenamento
- Usuários: `usuarios.json` (JSON plano, não é DB sensível)
- Auditorias: `Inventario_Resultados.xlsx` (Excel com fórmulas)
- Evidências: Pastas por ID de auditoria

---

## 🤝 Contribuição

Contribuições são bem-vindas! Por favor:

---

## 👤 Autor

**Gabriel Passos**  
Assistente de Logística | Desenvolvedor Python  
[LinkedIn](https://linkedin.com/in/seu-perfil) | [GitHub](https://github.com/seu-usuario)

---
