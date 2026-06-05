"""
Bootstrap de testes: injeta mocks no sys.modules ANTES de qualquer import da
aplicação, evitando chamadas reais ao Azure Key Vault e à Microsoft Graph API.
"""
import sys
from unittest.mock import MagicMock, AsyncMock

TEST_API_KEY = "test-api-key-middleware-12345"

# ── 1. Mock do Azure Key Vault ────────────────────────────────────────────────
# Substitui src.core.vault_client antes que qualquer módulo da app seja importado.
# Todos os get_secret() retornam TEST_API_KEY, permitindo que settings carregue
# sem precisar de credenciais Azure reais.
_vault_instance = MagicMock()
_vault_instance.get_secret.return_value = TEST_API_KEY

_vault_module = MagicMock()
_vault_module.VaultClient.return_value = _vault_instance
_vault_module.vault_client = _vault_instance

sys.modules["src.core.vault_client"] = _vault_module

# ── 2. Mock do GraphService (Microsoft Graph API) ────────────────────────────
# Substitui src.core.graph_client para que nenhum endpoint tente contatar a API.

async def _mock_stream():
    yield b"conteudo-de-teste"

_service_instance = MagicMock()
_service_instance.send_email = AsyncMock(
    return_value={"success": True, "message": "Email enviado com sucesso"}
)
_service_instance.list_files = AsyncMock(
    return_value={"success": True, "data": [{"name": "arquivo.xlsx", "size": 1024}]}
)
_service_instance.upload_file = AsyncMock(
    return_value={"success": True, "message": "Upload concluído", "data": {"id": "abc123"}}
)
_service_instance.upload_file_bytes = AsyncMock(
    return_value={"success": True, "message": "Upload concluído", "data": {"id": "abc123"}}
)
_service_instance.move_file = AsyncMock(
    return_value={"success": True, "message": "Arquivo movido", "new_name": "nota.pdf", "behavior_used": "fail"}
)
_service_instance.read_excel_values = AsyncMock(
    return_value={"values": [["ID", "Nome", "Valor"], [1, "Item A", 100]]}
)
_service_instance.write_excel_cells = AsyncMock(
    return_value={"summary": [{"cell": "A1", "success": True, "error": None}]}
)
_service_instance.get_file_stream = AsyncMock(
    return_value=lambda: _mock_stream()
)

_graph_module = MagicMock()
_graph_module.GraphService.return_value = _service_instance

sys.modules["src.core.graph_client"] = _graph_module
# ─────────────────────────────────────────────────────────────────────────────

import pytest
from fastapi.testclient import TestClient


@pytest.fixture(scope="session")
def client():
    """
    TestClient com autenticação bypassed via dependency_overrides.
    Usar nos testes funcionais de endpoints.
    """
    from main import app
    from src.config.security import get_api_key

    async def _mock_api_key():
        return TEST_API_KEY

    app.dependency_overrides[get_api_key] = _mock_api_key
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()
