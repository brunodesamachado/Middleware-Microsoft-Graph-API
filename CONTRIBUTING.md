# Contribuindo com o Middleware Graph API

Obrigado pelo interesse em contribuir! Este documento descreve como configurar o ambiente, rodar os testes e enviar contribuições.

## Sumário

- [Configuração do Ambiente](#configuração-do-ambiente)
- [Executando os Testes](#executando-os-testes)
- [Abrindo Issues](#abrindo-issues)
- [Enviando Pull Requests](#enviando-pull-requests)
- [Estilo de Código](#estilo-de-código)

---

## Configuração do Ambiente

**Requisitos:** Python 3.11+

```bash
git clone https://github.com/brunodesamachado/Middleware-Microsoft-Graph-API.git
cd Middleware-Microsoft-Graph-API

python -m venv .venv
# Linux/macOS
source .venv/bin/activate
# Windows
.venv\Scripts\activate

pip install -r requirements.txt -r requirements-dev.txt

cp .env.example .env
# Edite .env com suas credenciais Azure
```

---

## Executando os Testes

Os testes **não requerem credenciais Azure** — todos os serviços externos são mockados automaticamente pelo `conftest.py`.

```bash
# Todos os testes
pytest tests/ -v

# Apenas testes de validação
pytest tests/test_validation.py -v

# Apenas testes de segurança
pytest tests/test_security.py -v
```

---

## Abrindo Issues

Antes de abrir, verifique se a issue já existe. Ao abrir, inclua:

- Descrição clara do problema ou da feature
- Versão do Python e sistema operacional
- Passos para reproduzir (se for um bug)
- Logs relevantes (remova dados sensíveis)

---

## Enviando Pull Requests

1. Faça um fork e crie um branch descritivo:
   ```bash
   git checkout -b feat/minha-feature
   # ou
   git checkout -b fix/descricao-do-bug
   ```
2. Escreva testes para as alterações
3. Confirme que todos os testes passam: `pytest tests/ -v`
4. Abra o PR com uma descrição clara do que foi feito e por quê

---

## Estilo de Código

- Python 3.11+, tipagem explícita onde possível
- Siga o padrão existente — async/await, Pydantic v2, FastAPI
- Nunca use `print()` em código de produção — use o logger: `from src.utils.logger import Logger`
- Segredos nunca entram no código — use `.env` localmente ou Azure Key Vault em produção
- Caminhos de arquivo devem ser validados contra path traversal (`..`)
