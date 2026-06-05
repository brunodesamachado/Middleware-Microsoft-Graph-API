"""
Testa a camada de autenticação (X-API-Key):
- Requisições sem chave devem receber 403
- Requisições com chave errada devem receber 403
- Requisições com chave correta devem passar
- /health não requer autenticação
"""
import pytest
from fastapi.testclient import TestClient

TEST_API_KEY = "test-api-key-middleware-12345"

SP_PAYLOAD = {"site_name": "TI-Industrial", "drive_name": "Documentos"}


@pytest.fixture()
def no_auth_client():
    """
    TestClient sem dependency_overrides — usa o fluxo real de validação da API Key.
    TEST_API_KEY corresponde ao valor retornado pelo vault mock em conftest.py,
    portanto settings.app.api_key == TEST_API_KEY durante os testes.
    """
    from main import app
    from src.config.security import get_api_key

    saved_overrides = dict(app.dependency_overrides)
    app.dependency_overrides.pop(get_api_key, None)

    with TestClient(app) as c:
        yield c

    app.dependency_overrides.update(saved_overrides)


def test_sem_api_key_retorna_403(no_auth_client):
    r = no_auth_client.post("/sharepoint/list", json=SP_PAYLOAD)
    assert r.status_code == 403


def test_api_key_invalida_retorna_403(no_auth_client):
    r = no_auth_client.post(
        "/sharepoint/list",
        json=SP_PAYLOAD,
        headers={"X-API-Key": "chave-completamente-errada"},
    )
    assert r.status_code == 403


def test_api_key_correta_e_aceita(no_auth_client):
    r = no_auth_client.post(
        "/sharepoint/list",
        json=SP_PAYLOAD,
        headers={"X-API-Key": TEST_API_KEY},
    )
    assert r.status_code == 200


def test_health_nao_requer_autenticacao(no_auth_client):
    """/health deve responder 200 sem nenhum header."""
    r = no_auth_client.get("/health")
    assert r.status_code == 200


def test_api_key_vazia_retorna_403(no_auth_client):
    r = no_auth_client.post(
        "/sharepoint/list",
        json=SP_PAYLOAD,
        headers={"X-API-Key": ""},
    )
    assert r.status_code == 403
