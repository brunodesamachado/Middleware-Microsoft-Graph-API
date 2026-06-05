"""
Testes funcionais dos endpoints (autenticação bypassed via conftest).
Valida status codes, estrutura de resposta e comportamento dos endpoints.
"""
import base64

SP = {"site_name": "TI-Industrial", "drive_name": "Documentos"}


def test_health(client):
    r = client.get("/health")
    assert r.status_code == 200
    assert r.json() == {"status": "online", "service": "Graph Middleware"}


class TestEmailEndpoints:
    def test_send_email_simples(self, client):
        payload = {
            "sender": "bot@empresa.com",
            "recipients": ["destino@empresa.com"],
            "subject": "Teste automatizado",
            "body": "<p>Relatório em anexo.</p>",
        }
        r = client.post("/email/send", json=payload)
        assert r.status_code == 200
        assert r.json()["success"] is True

    def test_send_email_com_anexo_base64(self, client):
        conteudo = base64.b64encode(b"conteudo do pdf").decode()
        payload = {
            "sender": "bot@empresa.com",
            "recipients": ["destino@empresa.com"],
            "subject": "Com Anexo",
            "body": "Segue em anexo.",
            "is_html": False,
            "attachments_data": [
                {"name": "relatorio.pdf", "content_b64": conteudo, "content_type": "application/pdf"}
            ],
        }
        r = client.post("/email/send", json=payload)
        assert r.status_code == 200

    def test_send_email_campos_obrigatorios_ausentes(self, client):
        r = client.post("/email/send", json={"sender": "bot@empresa.com"})
        assert r.status_code == 422


class TestSharePointEndpoints:
    def test_list_files_raiz(self, client):
        r = client.post("/sharepoint/list", json={**SP, "folder_path": ""})
        assert r.status_code == 200
        assert "data" in r.json()

    def test_list_files_subpasta(self, client):
        r = client.post("/sharepoint/list", json={**SP, "folder_path": "Relatorios/2026"})
        assert r.status_code == 200

    def test_download_arquivo(self, client):
        r = client.post("/sharepoint/download", json={**SP, "file_path": "Relatorios/janeiro.xlsx"})
        assert r.status_code == 200

    def test_upload_base64(self, client):
        conteudo = base64.b64encode(b"dados do arquivo").decode()
        payload = {**SP, "file_content_b64": conteudo, "filename": "relatorio.xlsx", "target_path": "Financeiro"}
        r = client.post("/sharepoint/upload", json=payload)
        assert r.status_code == 200
        assert r.json()["success"] is True

    def test_move_arquivo(self, client):
        payload = {**SP, "ori_path": "Entrada/nota.pdf", "dest_folder_path": "Processados", "conflict_behavior": "rename"}
        r = client.post("/sharepoint/move", json=payload)
        assert r.status_code == 200
        assert r.json()["success"] is True

    def test_move_conflict_behavior_invalido(self, client):
        payload = {**SP, "ori_path": "Entrada/nota.pdf", "dest_folder_path": "Processados", "conflict_behavior": "invalido"}
        r = client.post("/sharepoint/move", json=payload)
        assert r.status_code == 422


class TestExcelEndpoints:
    def test_read_excel(self, client):
        r = client.post("/sharepoint/excel/read", json={**SP, "file_path": "dados.xlsx", "sheet_name": "Planilha1"})
        assert r.status_code == 200
        assert "values" in r.json()
        assert isinstance(r.json()["values"], list)

    def test_write_excel(self, client):
        payload = {**SP, "file_path": "dados.xlsx", "sheet_name": "Planilha1", "values": {"A1": "ID", "B1": "Nome", "C1": 42}}
        r = client.post("/sharepoint/excel/write", json=payload)
        assert r.status_code == 200
        assert "summary" in r.json()

    def test_read_excel_path_traversal_rejeitado(self, client):
        r = client.post("/sharepoint/excel/read", json={**SP, "file_path": "../config.yaml", "sheet_name": "Planilha1"})
        assert r.status_code == 422
