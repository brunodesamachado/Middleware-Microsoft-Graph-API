import os
import base64
import requests
from typing import List, Dict, Any, Literal, Optional


class MiddlewareGraphClient:
    """
    Cliente Python para consumir o Middleware Microsoft Graph API.

    Encapsula todas as chamadas HTTP, incluindo encoding Base64 para uploads
    e download atômico (arquivo .tmp → renomeia ao concluir com sucesso).

    Exemplo de uso:
        client = MiddlewareGraphClient(
            base_url="http://10.0.0.50:8001",
            api_key="sua-chave-do-key-vault"
        )
        print(client.get_health())
    """

    def __init__(self, base_url: str, api_key: str):
        """
        :param base_url: URL base do middleware (ex: "http://10.0.0.50:8001").
        :param api_key:  Valor do segredo API-KEY-MIDDLEWARE no Azure Key Vault.
        """
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.headers = {
            "X-API-Key": self.api_key,
            "Content-Type": "application/json",
        }

    def get_health(self) -> dict:
        """Verifica se o serviço está ativo."""
        try:
            resp = requests.get(f"{self.base_url}/health")
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise ConnectionError(f"Erro ao conectar ao middleware: {e}") from e

    def send_email(
        self,
        sender: str,
        recipients: List[str],
        subject: str,
        body: str,
        is_html: bool = True,
        attachment_folder: Optional[str] = None,
        attachments_data: Optional[List[Dict]] = None,
    ) -> dict:
        """
        Envia um email usando o serviço de middleware.

        :param sender:            Email do remetente.
        :param recipients:        Lista de emails dos destinatários.
        :param subject:           Assunto do email.
        :param body:              Corpo do email.
        :param is_html:           Se True, envia como HTML (padrão: True).
        :param attachment_folder: Caminho de pasta local no servidor (uso legado).
        :param attachments_data:  Lista de dicionários com dados dos anexos em Base64.
        :return: Resposta da API em formato JSON.
        """
        payload = {
            "sender": sender,
            "recipients": recipients,
            "subject": subject,
            "body": body,
            "is_html": is_html,
            "attachment_folder": attachment_folder,
            "attachments_data": attachments_data,
        }
        try:
            resp = requests.post(
                f"{self.base_url}/email/send", headers=self.headers, json=payload
            )
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise ValueError(
                f"Erro ao enviar email: {e.response.text if e.response else e}"
            ) from e

    def send_email_with_files(
        self,
        sender: str,
        recipients: List[str],
        subject: str,
        body: str,
        is_html: bool = True,
        attachment_paths: Optional[List[str]] = None,
    ) -> dict:
        """
        Envia email com arquivos locais do cliente via multipart/form-data.
        Não requer codificação Base64 manual — o servidor faz a conversão internamente.

        :param sender:            Email do remetente.
        :param recipients:        Lista de emails dos destinatários.
        :param subject:           Assunto do email.
        :param body:              Corpo do email (HTML ou texto simples).
        :param is_html:           Se True, envia como HTML (padrão: True).
        :param attachment_paths:  Lista de caminhos de arquivos locais a anexar.
        :return: Resposta da API em formato JSON.
        :raises FileNotFoundError: Se algum arquivo não existir.
        :raises ValueError:        Em caso de erro na requisição.
        """
        headers = {"X-API-Key": self.api_key}

        data = {
            "sender": sender,
            "recipients": recipients,
            "subject": subject,
            "body": body,
            "is_html": str(is_html).lower(),
        }

        files = []
        try:
            for path in attachment_paths or []:
                if not os.path.exists(path):
                    raise FileNotFoundError(f"Arquivo não encontrado: {path}")
                files.append(
                    ("attachments", (os.path.basename(path), open(path, "rb")))
                )

            resp = requests.post(
                f"{self.base_url}/email/send-files",
                headers=headers,
                data=data,
                files=files if files else None,
            )
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise ValueError(
                f"Erro ao enviar email com arquivos: {e.response.text if e.response else e}"
            ) from e
        finally:
            for _, (_, file_obj) in files:
                file_obj.close()

    def upload_file(
        self,
        site_name: str,
        drive_name: str,
        local_file_path: str,
        target_path: str = "",
    ) -> dict:
        """
        Faz o upload de um arquivo local para o SharePoint.
        O arquivo é lido localmente e enviado como Base64 no corpo da requisição.

        :param site_name:       Nome do site SharePoint.
        :param drive_name:      Nome do Drive (biblioteca de documentos).
        :param local_file_path: Caminho do arquivo local a ser enviado.
        :param target_path:     Pasta de destino no SharePoint (criada automaticamente se não existir).
        :return: Resposta da API em formato JSON.
        :raises FileNotFoundError: Se o arquivo não existir.
        :raises IOError:           Em caso de erro no upload.
        """
        if not os.path.exists(local_file_path):
            raise FileNotFoundError(f"Arquivo não encontrado: {local_file_path}")

        with open(local_file_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("utf-8")

        payload = {
            "site_name": site_name,
            "drive_name": drive_name,
            "file_content_b64": content_b64,
            "filename": os.path.basename(local_file_path),
            "target_path": target_path,
        }
        try:
            resp = requests.post(
                f"{self.base_url}/sharepoint/upload", headers=self.headers, json=payload
            )
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise IOError(
                f"Falha no upload para o SharePoint: {e.response.text if e.response else e}"
            ) from e

    def list_files(
        self, site_name: str, drive_name: str, folder_path: str = ""
    ) -> dict:
        """
        Lista arquivos em uma pasta do SharePoint.

        :param site_name:   Nome do site SharePoint.
        :param drive_name:  Nome do Drive.
        :param folder_path: Caminho da pasta (vazio = raiz).
        :return: Resposta da API com a lista de arquivos.
        """
        payload = {
            "site_name": site_name,
            "drive_name": drive_name,
            "folder_path": folder_path,
        }
        try:
            resp = requests.post(
                f"{self.base_url}/sharepoint/list", headers=self.headers, json=payload
            )
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise ValueError(
                f"Erro ao listar arquivos: {e.response.text if e.response else e}"
            ) from e

    def download_file(
        self, site_name: str, drive_name: str, file_path_sp: str, save_as_path: str
    ):
        """
        Baixa um arquivo do SharePoint usando estratégia de escrita atômica.
        Só cria o arquivo final se o download for concluído com sucesso.

        :param site_name:     Nome do site SharePoint.
        :param drive_name:    Nome do Drive.
        :param file_path_sp:  Caminho do arquivo no SharePoint.
        :param save_as_path:  Caminho local onde o arquivo será salvo.
        :raises IOError:      Em caso de falha no download.
        """
        payload = {
            "site_name": site_name,
            "drive_name": drive_name,
            "file_path": file_path_sp,
        }

        local_dir = os.path.dirname(save_as_path)
        if local_dir:
            os.makedirs(local_dir, exist_ok=True)

        temp_path = f"{save_as_path}.tmp"
        try:
            with requests.post(
                f"{self.base_url}/sharepoint/download",
                headers=self.headers,
                json=payload,
                stream=True,
            ) as r:
                if r.status_code != 200:
                    raise IOError(f"Erro Middleware ({r.status_code}): {r.text}")

                with open(temp_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

            if os.path.exists(save_as_path):
                os.remove(save_as_path)
            os.rename(temp_path, save_as_path)

        except Exception as e:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            raise IOError(f"Falha no download: {e}") from e

    def move_file(
        self,
        site_name: str,
        drive_name: str,
        ori_path: str,
        dest_folder_path: str,
        conflict_behavior: Literal["fail", "replace", "rename"] = "fail",
    ) -> dict:
        """
        Move ou renomeia um arquivo dentro do SharePoint.

        :param site_name:          Nome do site SharePoint.
        :param drive_name:         Nome do Drive.
        :param ori_path:           Caminho atual do arquivo.
        :param dest_folder_path:   Pasta de destino.
        :param conflict_behavior:  Comportamento em conflito: "fail", "replace" ou "rename".
        :return: Resposta da API.
        """
        payload = {
            "site_name": site_name,
            "drive_name": drive_name,
            "ori_path": ori_path,
            "dest_folder_path": dest_folder_path,
            "conflict_behavior": conflict_behavior,
        }
        try:
            resp = requests.post(
                f"{self.base_url}/sharepoint/move", headers=self.headers, json=payload
            )
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise ValueError(
                f"Erro ao mover arquivo: {e.response.text if e.response else e}"
            ) from e

    def read_excel(
        self, site_name: str, drive_name: str, file_path: str, sheet_name: str
    ) -> dict:
        """
        Lê todos os valores da área utilizada de uma aba do Excel no SharePoint.

        :param site_name:  Nome do site SharePoint.
        :param drive_name: Nome do Drive.
        :param file_path:  Caminho do arquivo Excel no SharePoint.
        :param sheet_name: Nome da aba (worksheet).
        :return: Dicionário com chave "values" contendo uma matriz (lista de listas).
        """
        payload = {
            "site_name": site_name,
            "drive_name": drive_name,
            "file_path": file_path,
            "sheet_name": sheet_name,
        }
        try:
            resp = requests.post(
                f"{self.base_url}/sharepoint/excel/read",
                headers=self.headers,
                json=payload,
            )
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise ValueError(
                f"Erro ao ler Excel: {e.response.text if e.response else e}"
            ) from e

    def write_excel(
        self,
        site_name: str,
        drive_name: str,
        file_path: str,
        sheet_name: str,
        values: Dict[str, Any],
    ) -> dict:
        """
        Escreve valores em células específicas de uma planilha Excel no SharePoint.

        :param site_name:  Nome do site SharePoint.
        :param drive_name: Nome do Drive.
        :param file_path:  Caminho do arquivo Excel no SharePoint.
        :param sheet_name: Nome da aba (worksheet).
        :param values:     Dicionário no formato {"A1": valor, "B2": valor, ...}.
        :return: Resposta da API com o resumo por célula.
        """
        payload = {
            "site_name": site_name,
            "drive_name": drive_name,
            "file_path": file_path,
            "sheet_name": sheet_name,
            "values": values,
        }
        try:
            resp = requests.post(
                f"{self.base_url}/sharepoint/excel/write",
                headers=self.headers,
                json=payload,
            )
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            raise ValueError(
                f"Erro ao escrever no Excel: {e.response.text if e.response else e}"
            ) from e
