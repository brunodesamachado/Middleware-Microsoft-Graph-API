"""
Testa a camada de validação dos DTOs:
- Proteção contra path traversal (..)
- Limite de tamanho de upload (Base64)
- Campos obrigatórios e regras de negócio
"""
import base64

import pytest
from pydantic import ValidationError

from main import (
    MAX_UPLOAD_BYTES,
    SPDownloadRequest,
    SPExcelReadRequest,
    SPExcelWriteRequest,
    SPListFilesRequest,
    SPMoveRequest,
    SPUploadRequest,
)

BASE = {"site_name": "TI-Industrial", "drive_name": "Documentos"}


class TestPathTraversal:
    """Caminhos com '..' ou caracteres nulos devem ser rejeitados."""

    @pytest.mark.parametrize("path", [
        "../etc/passwd",
        "pasta/../segredo",
        "a/../../b",
        "normal/../../fora",
    ])
    def test_list_rejeita_traversal(self, path):
        with pytest.raises(ValidationError, match="inválido"):
            SPListFilesRequest(**BASE, folder_path=path)

    @pytest.mark.parametrize("path", ["../config.yaml", "pasta/../outro"])
    def test_download_rejeita_traversal(self, path):
        with pytest.raises(ValidationError, match="inválido"):
            SPDownloadRequest(**BASE, file_path=path)

    @pytest.mark.parametrize("path", ["../etc/passwd", "subdir/../escape"])
    def test_move_rejeita_traversal_na_origem(self, path):
        with pytest.raises(ValidationError, match="inválido"):
            SPMoveRequest(**BASE, ori_path=path, dest_folder_path="Destino")

    @pytest.mark.parametrize("path", ["../pasta-proibida"])
    def test_move_rejeita_traversal_no_destino(self, path):
        with pytest.raises(ValidationError, match="inválido"):
            SPMoveRequest(**BASE, ori_path="Entrada/file.pdf", dest_folder_path=path)

    @pytest.mark.parametrize("path", ["../config.yaml"])
    def test_excel_read_rejeita_traversal(self, path):
        with pytest.raises(ValidationError, match="inválido"):
            SPExcelReadRequest(**BASE, file_path=path, sheet_name="Planilha1")

    @pytest.mark.parametrize("path", ["../config.yaml"])
    def test_excel_write_rejeita_traversal(self, path):
        with pytest.raises(ValidationError, match="inválido"):
            SPExcelWriteRequest(**BASE, file_path=path, sheet_name="Planilha1", values={"A1": "x"})

    def test_caminho_valido_e_aceito(self):
        """Caminhos normais não devem ser rejeitados."""
        req = SPDownloadRequest(**BASE, file_path="Relatorios/2026/janeiro.xlsx")
        assert req.file_path == "Relatorios/2026/janeiro.xlsx"

    def test_caminho_vazio_e_aceito_em_list(self):
        """folder_path vazio (raiz) é válido."""
        req = SPListFilesRequest(**BASE, folder_path="")
        assert req.folder_path == ""


class TestUploadValidation:
    """Testa limites de tamanho e campos obrigatórios no upload."""

    def test_base64_acima_do_limite_e_rejeitado(self):
        conteudo_gigante = base64.b64encode(b"x" * (MAX_UPLOAD_BYTES + 1)).decode()
        with pytest.raises(ValidationError, match="25 MB"):
            SPUploadRequest(**BASE, file_content_b64=conteudo_gigante, filename="big.bin")

    def test_base64_dentro_do_limite_e_aceito(self):
        conteudo_ok = base64.b64encode(b"dados de teste").decode()
        req = SPUploadRequest(**BASE, file_content_b64=conteudo_ok, filename="teste.bin")
        assert req.filename == "teste.bin"

    def test_sem_fonte_de_arquivo_e_rejeitado(self):
        """Deve fornecer local_file_path ou file_content_b64."""
        with pytest.raises(ValidationError):
            SPUploadRequest(**BASE)

    def test_base64_sem_filename_e_rejeitado(self):
        """filename é obrigatório quando file_content_b64 é fornecido."""
        conteudo = base64.b64encode(b"dados").decode()
        with pytest.raises(ValidationError):
            SPUploadRequest(**BASE, file_content_b64=conteudo)

    def test_local_file_path_aceito(self):
        req = SPUploadRequest(**BASE, local_file_path="/app/uploads/relatorio.xlsx")
        assert req.local_file_path == "/app/uploads/relatorio.xlsx"
