# Middleware Microsoft Graph API

![CI](https://github.com/brunodesamachado/Middleware-Microsoft-Graph-API/actions/workflows/ci.yml/badge.svg)
![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

API Gateway / Middleware para centralizar a comunicação das automações (RPA) com a Microsoft Graph API.

Abstrai a complexidade de autenticação OAuth2, gerencia renovação de tokens e contextos de segurança, permitindo que robôs e scripts interajam com SharePoint e Exchange (E-mail) utilizando apenas uma chave de API interna e chamadas REST simples.

---

## Sumário

- [Funcionalidades](#funcionalidades)
- [Tech Stack](#tech-stack)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Configuração e Segurança](#configuração-e-segurança)
- [Instalação e Execução](#instalação-e-execução)
- [Autenticação nos Endpoints](#autenticação-nos-endpoints)
- [Referência de Endpoints](#referência-de-endpoints)
  - [Health Check](#health-check)
  - [E-mail](#e-mail)
  - [SharePoint — Arquivos](#sharepoint--arquivos)
  - [SharePoint — Excel](#sharepoint--excel)
- [Usando o Client Python (client.py)](#usando-o-client-python-clientpy)
- [Logs](#logs)
- [Testes](#testes)
- [Contribuindo](#contribuindo)
- [Autor](#autor)

---

## Funcionalidades

- **Gestão de Tokens**: Cache e renovação automática de tokens OAuth2 (padrão Singleton).
- **Segurança Zero-Trust**: Integração nativa com Azure Key Vault via RBAC — nenhuma credencial fica exposta no código ou em disco.
- **Contexto Duplo**: Credenciais segregadas para E-mail e SharePoint (Princípio do Privilégio Mínimo).
- **Proteção contra Timing Attack**: Comparação de API Key via `secrets.compare_digest()`.
- **Rate Limiting**: Limites por IP em todos os endpoints (via `slowapi`), diferenciados por peso da operação.
- **Validação de Path Traversal**: Todos os caminhos de arquivo são validados contra ataques `../`.
- **Limite de Upload**: Uploads rejeitados acima de 25 MB (Base64 e multipart).
- **Alta Performance**: HTTPX com suporte a HTTP/2 e processamento 100% assíncrono.
- **Smart Streaming**: Downloads transmitidos via stream direto da Microsoft para o cliente (sem disco intermediário).
- **Upload via Base64**: Aceita o conteúdo do arquivo codificado em Base64, sem precisar de um caminho local no servidor.
- **Auto-criação de Pastas**: Cria automaticamente a hierarquia de pastas no SharePoint durante uploads, se ela não existir.
- **Operações Excel via Graph**: Leitura e escrita direta em planilhas Excel hospedadas no SharePoint, sem precisar baixar o arquivo.
- **Documentação Automática**: Swagger UI integrado em `/docs`.

---

## Tech Stack

| Componente | Tecnologia |
|---|---|
| Linguagem | Python 3.9+ |
| Framework Web | FastAPI + Uvicorn |
| Cliente HTTP | HTTPX (Async, HTTP/2) |
| Validação | Pydantic V2 |
| Segurança | azure-identity, azure-keyvault-secrets |
| Configuração | PyYAML + Pydantic Settings + python-dotenv |
| I/O Assíncrono | aiofiles |

---

## Estrutura do Projeto

```text
middleware-graph-api/
├── main.py                   # Ponto de entrada da aplicação FastAPI
├── client.py                 # Biblioteca Python para consumir o middleware
├── config.yaml               # Configurações não-sensíveis (porta, paths, tenant)
├── requirements.txt          # Dependências de produção
├── requirements-dev.txt      # Dependências de desenvolvimento/testes
├── .env.example              # Template de variáveis de ambiente
├── .env                      # Credenciais locais (ignorado pelo Git)
├── logs/                     # Logs rotacionados (apenas .gitkeep versionado)
├── tests/
│   ├── conftest.py           # Fixtures e mocks globais (Azure, Graph API)
│   ├── test_api.py           # Testes funcionais dos endpoints
│   ├── test_security.py      # Testes de autenticação (X-API-Key)
│   └── test_validation.py    # Testes de validação (path traversal, upload)
└── src/
    ├── config/
    │   ├── settings.py       # Pydantic Settings: lê config.yaml + Key Vault
    │   └── security.py       # Validação da API Key (secrets.compare_digest)
    ├── core/
    │   ├── graph_client.py   # GraphTokenManager e GraphService (lógica principal)
    │   └── vault_client.py   # Singleton para Azure Key Vault
    └── utils/
        └── logger.py         # Logging centralizado com rotação de arquivos
```

---

## Configuração e Segurança

A configuração é dividida em camadas para garantir que nenhuma credencial fique no código ou no repositório.

### Camada 1 — `config.yaml` (não-sensível, versionado)

```yaml
app:
  api_port: 8001
  api_producao: false        # true em produção (desativa o --reload do uvicorn)
  allowed_origins: ["*"]

azure:
  tenant_name: "suaempresa.sharepoint.com"

paths:
  logs: logs
  data:
    input:   data\input_data
    output:  data\output_data
    backup:  data\backup_data
    results: data\results
```

### Camada 2 — Segredos no Azure Key Vault

A aplicação lê automaticamente os seguintes segredos do cofre configurado:

| Segredo no Key Vault | Descrição |
|---|---|
| `TENANT-ID` | Azure Tenant ID |
| `CLIENT-ID-SHAREPOINT` | Client ID do App Registration para SharePoint |
| `CLIENT-SECRET-SHAREPOINT` | Client Secret do App Registration para SharePoint |
| `CLIENT-ID-EMAIL` | Client ID do App Registration para E-mail |
| `CLIENT-SECRET-EMAIL` | Client Secret do App Registration para E-mail |
| `API-KEY-MIDDLEWARE` | Chave de API interna usada pelos consumidores |

### Camada 3 — Identidade de Serviço (como o middleware acessa o Key Vault)

**Ambiente de Desenvolvimento com Azure** — crie um `.env` na raiz do projeto:

```ini
KEY_VAULT_URL=https://nome-do-seu-cofre.vault.azure.net/
AZURE_TENANT_ID=seu_tenant_id
AZURE_CLIENT_ID=seu_client_id
AZURE_CLIENT_SECRET=sua_client_secret
```

**Ambiente de Desenvolvimento sem Azure (fallback local)** — deixe `KEY_VAULT_URL` vazio e forneça os segredos diretamente no `.env`:

```ini
# KEY_VAULT_URL vazio = Key Vault ignorado, segredos lidos do .env
KEY_VAULT_URL=

API_KEY_MIDDLEWARE=minha-chave-local
TENANT_ID=seu-tenant-id
CLIENT_ID_SHAREPOINT=seu-client-id-sp
CLIENT_SECRET_SHAREPOINT=seu-client-secret-sp
CLIENT_ID_EMAIL=seu-client-id-email
CLIENT_SECRET_EMAIL=seu-client-secret-email
```

> O servidor sobe normalmente neste modo. Endpoints que chamam o Microsoft Graph falharão com erro de autenticação, mas `/health` e o Swagger UI (`/docs`) funcionam sem restrições.

**Ambiente de Produção (NSSM / Windows Service)** — não use `.env`. Cadastre as variáveis diretamente na aba **Environment** da configuração do serviço no NSSM:

```
KEY_VAULT_URL=https://...
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
```

---

## Instalação e Execução

### 1. Localmente (Python)

```bash
# Instale as dependências
pip install -r requirements.txt

# Execute a aplicação
python main.py
# ou
uvicorn main:app --host 0.0.0.0 --reload
```

A aplicação iniciará na porta configurada no `config.yaml` (padrão: **8001**).  
Swagger UI disponível em: `http://127.0.0.1:8001/docs`

### 2. Docker

```bash
# Build da imagem
docker build -t middleware-graph-api .

# Run do container (injeta credenciais via --env-file)
docker run -d -p 8001:8001 --name middleware-graph-api \
  --env-file .env \
  -v /caminho/para/config.yaml:/app/config.yaml \
  middleware-graph-api
```

### 3. Produção — Windows Service (NSSM)

```bash
nssm install MiddlewareGraphAPI "C:\Python\python.exe" "C:\middleware\main.py"
# Configure as variáveis de ambiente na aba Environment do NSSM
nssm start MiddlewareGraphAPI
```

---

## Autenticação nos Endpoints

Todos os endpoints (exceto `/health`) exigem autenticação via header HTTP:

```
X-API-Key: <valor do segredo API-KEY-MIDDLEWARE no Key Vault>
```

### Autenticar no Swagger UI

1. Acesse `/docs`.
2. Clique no botão **Authorize** (cadeado) no canto superior direito.
3. No campo **APIKeyHeader**, insira o valor da chave.
4. Clique em **Authorize** → **Close**.

---

## Referência de Endpoints

### Health Check

#### `GET /health`

Verifica se o serviço está ativo. Não requer autenticação.

**Resposta:**
```json
{
  "status": "online",
  "service": "Graph Middleware"
}
```

---

### E-mail

#### `POST /email/send`

Envia e-mails com ou sem anexos. Aceita conteúdo HTML ou texto simples.

**Request Body:**

```json
{
  "sender": "automacao@suaempresa.com",
  "recipients": ["destinatario1@suaempresa.com", "destinatario2@suaempresa.com"],
  "subject": "Relatório Diário",
  "body": "<h1>Olá!</h1><p>Segue o relatório em anexo.</p>",
  "is_html": true,
  "attachments_data": [
    {
      "name": "relatorio.pdf",
      "content_b64": "JVBERi0xLjQg...",
      "content_type": "application/pdf"
    }
  ]
}
```

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `sender` | string | Sim | Endereço de e-mail do remetente |
| `recipients` | string[] | Sim | Lista de destinatários |
| `subject` | string | Sim | Assunto do e-mail |
| `body` | string | Sim | Corpo do e-mail (HTML ou texto simples) |
| `is_html` | boolean | Não (padrão: `true`) | Define se o corpo é HTML |
| `attachment_folder` | string | Não | Caminho de pasta local **no servidor** (uso legado) |
| `attachments_data` | array | Não | Anexos fornecidos pelo cliente em Base64 |
| `attachments_data[].name` | string | Sim | Nome do arquivo com extensão |
| `attachments_data[].content_b64` | string | Sim | Conteúdo do arquivo em Base64 |
| `attachments_data[].content_type` | string | Não (padrão: `application/octet-stream`) | MIME type do arquivo |

**Resposta de Sucesso (`200`):**
```json
{
  "success": true,
  "message": "Email enviado com sucesso"
}
```

---

### SharePoint — Arquivos

Todos os endpoints de SharePoint requerem `site_name` e `drive_name` para identificar o contexto.

#### `POST /sharepoint/list`

Lista arquivos e pastas dentro de um diretório do SharePoint.

**Request Body:**
```json
{
  "site_name": "TI-Industrial",
  "drive_name": "Documentos",
  "folder_path": "Relatorios/2026"
}
```

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `site_name` | string | Sim | Nome do site SharePoint |
| `drive_name` | string | Sim | Nome da biblioteca de documentos |
| `folder_path` | string | Não (padrão: `""`) | Caminho da pasta (vazio = raiz) |

**Resposta de Sucesso (`200`):**
```json
{
  "success": true,
  "data": [
    {
      "id": "01ABCD...",
      "name": "janeiro.xlsx",
      "size": 20480,
      "createdDateTime": "2026-01-02T10:00:00Z",
      "lastModifiedDateTime": "2026-01-15T14:30:00Z"
    }
  ]
}
```

---

#### `POST /sharepoint/download`

Faz o download de um arquivo do SharePoint. A resposta é transmitida como stream binário diretamente ao cliente.

**Request Body:**
```json
{
  "site_name": "TI-Industrial",
  "drive_name": "Documentos",
  "file_path": "Relatorios/2026/janeiro.xlsx"
}
```

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `site_name` | string | Sim | Nome do site SharePoint |
| `drive_name` | string | Sim | Nome da biblioteca de documentos |
| `file_path` | string | Sim | Caminho completo do arquivo |

**Resposta de Sucesso (`200`):** Stream binário com `Content-Disposition: attachment`.

**Exemplo com `requests`:**
```python
import requests

headers = {"X-API-Key": "sua-chave-aqui"}
payload = {
    "site_name": "TI-Industrial",
    "drive_name": "Documentos",
    "file_path": "Relatorios/2026/janeiro.xlsx"
}

with requests.post("http://localhost:8001/sharepoint/download",
                   json=payload, headers=headers, stream=True) as r:
    r.raise_for_status()
    with open("janeiro.xlsx", "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)
```

---

#### `POST /sharepoint/upload`

Faz o upload de um arquivo para o SharePoint. Aceita duas formas:

**Forma 1 — Arquivo Base64 (recomendado para consumidores externos):**
```json
{
  "site_name": "TI-Industrial",
  "drive_name": "Documentos",
  "file_content_b64": "JVBERi0xLjQg...",
  "filename": "relatorio.xlsx",
  "target_path": "Financeiro/Notas"
}
```

**Forma 2 — Arquivo local no servidor (uso interno):**
```json
{
  "site_name": "TI-Industrial",
  "drive_name": "Documentos",
  "local_file_path": "C:\\RPA\\output\\relatorio.xlsx",
  "target_path": "Financeiro/Notas"
}
```

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `site_name` | string | Sim | Nome do site SharePoint |
| `drive_name` | string | Sim | Nome da biblioteca de documentos |
| `file_content_b64` | string | Condicional | Conteúdo do arquivo em Base64 |
| `filename` | string | Se usar `file_content_b64` | Nome do arquivo com extensão |
| `local_file_path` | string | Condicional | Caminho local no servidor |
| `target_path` | string | Não (padrão: `""`) | Pasta de destino no SharePoint (criada automaticamente se não existir) |

> Forneça **apenas um** entre `file_content_b64` ou `local_file_path`.

**Resposta de Sucesso (`200`):**
```json
{
  "success": true,
  "message": "Upload concluído",
  "data": {
    "id": "01ABCD...",
    "name": "relatorio.xlsx",
    "size": 20480
  }
}
```

---

#### `POST /sharepoint/move`

Move ou renomeia um arquivo dentro do SharePoint.

**Request Body:**
```json
{
  "site_name": "TI-Industrial",
  "drive_name": "Documentos",
  "ori_path": "Entrada/nota_fiscal.pdf",
  "dest_folder_path": "Processados",
  "conflict_behavior": "rename"
}
```

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `site_name` | string | Sim | Nome do site SharePoint |
| `drive_name` | string | Sim | Nome da biblioteca de documentos |
| `ori_path` | string | Sim | Caminho atual do arquivo |
| `dest_folder_path` | string | Sim | Pasta de destino |
| `conflict_behavior` | enum | Não (padrão: `"fail"`) | O que fazer se o arquivo já existir no destino: `"fail"`, `"replace"` ou `"rename"` |

**Resposta de Sucesso (`200`):**
```json
{
  "success": true,
  "message": "Arquivo movido com sucesso",
  "new_name": "nota_fiscal.pdf",
  "behavior_used": "rename"
}
```

---

### SharePoint — Excel

Permite ler e escrever em planilhas Excel hospedadas no SharePoint diretamente via Graph API, sem precisar baixar o arquivo.

#### `POST /sharepoint/excel/read`

Lê todos os valores da área utilizada de uma planilha.

**Request Body:**
```json
{
  "site_name": "TI-Industrial",
  "drive_name": "Documentos",
  "file_path": "Relatorios/financeiro.xlsx",
  "sheet_name": "Planilha1"
}
```

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `site_name` | string | Sim | Nome do site SharePoint |
| `drive_name` | string | Sim | Nome da biblioteca de documentos |
| `file_path` | string | Sim | Caminho completo do arquivo Excel |
| `sheet_name` | string | Sim | Nome da planilha (aba) |

**Resposta de Sucesso (`200`):**
```json
{
  "values": [
    ["ID", "Nome", "Valor"],
    [1, "Item A", 100.50],
    [2, "Item B", 200.75]
  ]
}
```

> `values` é uma matriz (lista de listas) onde cada lista interna representa uma linha.

---

#### `POST /sharepoint/excel/write`

Escreve valores em células específicas de uma planilha.

**Request Body:**
```json
{
  "site_name": "TI-Industrial",
  "drive_name": "Documentos",
  "file_path": "Relatorios/financeiro.xlsx",
  "sheet_name": "Planilha1",
  "values": {
    "A1": "ID",
    "B1": "Nome",
    "C1": "Valor",
    "A2": 1,
    "B2": "Item A",
    "C2": 500.50
  }
}
```

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `site_name` | string | Sim | Nome do site SharePoint |
| `drive_name` | string | Sim | Nome da biblioteca de documentos |
| `file_path` | string | Sim | Caminho completo do arquivo Excel |
| `sheet_name` | string | Sim | Nome da planilha (aba) |
| `values` | object | Sim | Mapa de célula → valor (ex: `"A1": "Texto"`, `"B2": 42`) |

**Resposta de Sucesso (`200`):**
```json
{
  "summary": [
    {"cell": "A1", "success": true, "error": null},
    {"cell": "B1", "success": true, "error": null},
    {"cell": "C2", "success": false, "error": "Session expired"}
  ]
}
```

> Cada célula é processada individualmente com até 3 tentativas automáticas.

---

## Usando o Client Python (`client.py`)

O projeto inclui um módulo `client.py` pronto para uso em scripts RPA. Ele encapsula todas as chamadas HTTP, incluindo encoding Base64 para uploads e download atômico com arquivo temporário.

### Configuração

```python
from client import MiddlewareGraphClient

client = MiddlewareGraphClient(
    base_url="http://10.0.0.50:8001",   # URL do servidor onde o middleware está rodando
    api_key="sua-chave-do-key-vault"
)
```

### Exemplos de Uso

#### Health Check
```python
status = client.get_health()
print(status)  # {"status": "online", "service": "Graph Middleware"}
```

#### Enviar E-mail com Anexo
```python
import base64

with open("relatorio.pdf", "rb") as f:
    conteudo_b64 = base64.b64encode(f.read()).decode("utf-8")

client.send_email(
    sender="automacao@suaempresa.com",
    recipients=["gestor@suaempresa.com"],
    subject="Relatório Automático",
    body="<p>Segue o relatório diário.</p>",
    is_html=True,
    attachments_data=[
        {
            "name": "relatorio.pdf",
            "content_b64": conteudo_b64,
            "content_type": "application/pdf"
        }
    ]
)
```

#### Listar Arquivos
```python
resultado = client.list_files(
    site_name="TI-Industrial",
    drive_name="Documentos",
    folder_path="Relatorios/2026"
)
for item in resultado["data"]:
    print(item["name"], item["size"])
```

#### Download de Arquivo
```python
# O client salva em disco de forma atômica (arquivo .tmp → renomeia ao concluir)
client.download_file(
    site_name="TI-Industrial",
    drive_name="Documentos",
    file_path_sp="Relatorios/2026/janeiro.xlsx",
    save_as_path="C:\\RPA\\downloads\\janeiro.xlsx"
)
```

#### Upload via Base64
```python
resultado = client.upload_file(
    site_name="TI-Industrial",
    drive_name="Documentos",
    local_file_path="C:\\RPA\\output\\processado.xlsx",
    target_path="Processados/2026"
)
print(resultado)
```

#### Mover Arquivo
```python
client.move_file(
    site_name="TI-Industrial",
    drive_name="Documentos",
    ori_path="Entrada/nota.pdf",
    dest_folder_path="Processados",
    conflict_behavior="rename"  # "fail" | "replace" | "rename"
)
```

#### Ler Planilha Excel
```python
resultado = client.read_excel(
    site_name="TI-Industrial",
    drive_name="Documentos",
    file_path="Controle/base.xlsx",
    sheet_name="Dados"
)
# resultado["values"] é uma lista de listas
for linha in resultado["values"]:
    print(linha)
```

#### Escrever em Planilha Excel
```python
client.write_excel(
    site_name="TI-Industrial",
    drive_name="Documentos",
    file_path="Controle/base.xlsx",
    sheet_name="Dados",
    values={
        "A2": "Bot-001",
        "B2": "Concluído",
        "C2": "2026-05-15"
    }
)
```

---

## Logs

Logs são gerados automaticamente na pasta `logs/` com rotação automática:

- **Tamanho máximo por arquivo:** 5 MB
- **Arquivos de backup mantidos:** 3
- **Nível padrão:** `INFO`

Cada requisição, erro e evento de autenticação é registrado com timestamp e contexto.

---

## Testes

Os testes não requerem credenciais Azure — os serviços externos são mockados automaticamente.

```bash
# Instalar dependências de desenvolvimento
pip install -r requirements.txt -r requirements-dev.txt

# Rodar todos os testes
pytest tests/ -v

# Rodar por categoria
pytest tests/test_validation.py -v   # validação de DTOs
pytest tests/test_security.py -v     # autenticação
pytest tests/test_api.py -v          # endpoints funcionais
```

A suíte cobre:
- **Validação**: path traversal, limite de upload, campos obrigatórios
- **Segurança**: rejeição de chaves ausentes/inválidas, aceitação da chave correta
- **Endpoints**: todos os recursos (Email, SharePoint, Excel)

---

## Contribuindo

Contribuições são bem-vindas! Consulte o [CONTRIBUTING.md](CONTRIBUTING.md) para instruções de configuração do ambiente e envio de PRs.

---

## Autor

**Bruno Machado**
