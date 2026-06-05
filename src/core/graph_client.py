import os
import base64
import time
import httpx
import aiofiles
import asyncio
from fastapi import HTTPException
from src.config.settings import settings
from src.utils.logger import Logger

# Logger
logger = Logger().get_logger()

# ==============================================================================
# 1. Gerenciador de Tokens (Singleton)
# ==============================================================================
class GraphTokenManager:
    """
    Gerencia tokens separadamente para contextos 'sharepoint' e 'email'.
    Armazena em cache e renova automaticamente se expirar.
    """
    def __init__(self):
        self._cache = {}
        self._token_url = f"https://login.microsoftonline.com/{settings.azure.tenant_id}/oauth2/v2.0/token"

    async def get_valid_token(self, context: str) -> str:
        if context not in ['email', 'sharepoint']:
            logger.error(f"Contexto de token inválido solicitado: {context}")
            raise ValueError(f"Contexto inválido: {context}. Use 'email' ou 'sharepoint'.")

        cached = self._cache.get(context)
        if cached and time.time() < (cached['expires'] - 300):
            logger.debug(f"Token em cache encontrado para o contexto '{context}'.")
            return cached['token']
        
        logger.info(f"Token para o contexto '{context}' expirado ou não encontrado. Obtendo um novo.")
        return await self._fetch_new_token(context)

    async def _fetch_new_token(self, context: str):
        if context == 'email':
            client_id = settings.azure.client_id_email
            client_secret = settings.azure.client_secret_email
            scope = "https://graph.microsoft.com/.default"
        else:
            client_id = settings.azure.client_id_sharepoint
            client_secret = settings.azure.client_secret_sharepoint
            scope = "https://graph.microsoft.com/.default"
        
        payload = {
            "client_id": client_id,
            "scope": scope,
            "client_secret": client_secret,
            "grant_type": "client_credentials",
        }
        
        async with httpx.AsyncClient(http2=True) as client:
            try:
                logger.debug(f"Solicitando novo token para o contexto '{context}' na URL: {self._token_url}")
                response = await client.post(self._token_url, data=payload)
                response.raise_for_status()
                data = response.json()
                
                self._cache[context] = {
                    'token': data["access_token"],
                    'expires': time.time() + int(data["expires_in"])
                }
                logger.info(f"Novo token adquirido com sucesso para o contexto '{context}'.")
                return data["access_token"]
            except httpx.HTTPStatusError as e:
                logger.critical(f"Falha ao obter token para '{context}'. Status: {e.response.status_code}. Resposta: {e.response.text}", exc_info=True)
                raise HTTPException(status_code=401, detail=f"Erro crítico de autenticação com o Azure para o contexto '{context}'.")
            except Exception as e:
                logger.critical(f"Erro inesperado ao buscar token para '{context}': {e}", exc_info=True)
                raise HTTPException(status_code=500, detail="Erro interno ao contatar o serviço de autenticação.")

# Instância global
token_manager = GraphTokenManager()

# ==============================================================================
# 2. Serviço Graph (Lógica de Negócio)
# ==============================================================================
class GraphService:
    def __init__(self):
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.tenant_name = settings.azure.tenant_name

    async def _get_headers(self, context: str):
        token = await token_manager.get_valid_token(context)
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

    # ... (demais métodos com logging) ...
    async def _get_site_id(self, site_name: str, headers: dict) -> str:
        url = f"{self.base_url}/sites/{self.tenant_name}:/sites/{site_name}"
        logger.debug(f"Buscando ID do site: '{site_name}'")
        async with httpx.AsyncClient(http2=True) as client:
            resp = await client.get(url, headers=headers)
            if resp.status_code != 200:
                logger.error(f"Falha ao buscar ID do site '{site_name}'. Status: {resp.status_code}, Detalhe: {resp.text}")
                raise HTTPException(status_code=404, detail=f"Site '{site_name}' não encontrado ou acesso negado.")
            site_id = resp.json().get("id")
            logger.info(f"ID do site '{site_name}' encontrado: {site_id}")
            return site_id

    async def list_files(self, site_name: str, drive_name: str, folder_path: str):
        logger.info(f"Listando arquivos em: Site='{site_name}', Drive='{drive_name}', Pasta='{folder_path}'")
        try:
            headers = await self._get_headers('sharepoint')
            site_id = await self._get_site_id(site_name, headers)
            drive_id = await self._get_drive_id(site_id, drive_name, headers)

            path_format = f":/{folder_path}:" if folder_path else ""
            url = f"{self.base_url}/drives/{drive_id}/root{path_format}/children"
            
            logger.debug(f"URL para listar arquivos: {url}")
            async with httpx.AsyncClient(http2=True) as client:
                resp = await client.get(url, headers=headers)
                resp.raise_for_status()
                logger.info("Listagem de arquivos bem-sucedida.")
                return {"success": True, "data": resp.json().get('value')}
        except httpx.HTTPStatusError as e:
            logger.error(f"Erro HTTP ao listar arquivos: {e.response.text}", exc_info=True)
            raise HTTPException(status_code=e.response.status_code, detail=e.response.text)
        except Exception as e:
            logger.error(f"Erro inesperado ao listar arquivos: {e}", exc_info=True)
            if isinstance(e, HTTPException):
                raise e
            raise HTTPException(status_code=500, detail=str(e))
    
    async def _ensure_folder_path_exists(self, drive_id: str, folder_path: str, headers: dict):
        """Cria cada nível do caminho de pastas no SharePoint, ignorando se já existir."""
        if not folder_path:
            return

        parts = [p for p in folder_path.strip("/").split("/") if p]
        current_path = ""

        async with httpx.AsyncClient(http2=True, timeout=30.0) as client:
            for part in parts:
                url = (
                    f"{self.base_url}/drives/{drive_id}/root/children"
                    if not current_path
                    else f"{self.base_url}/drives/{drive_id}/root:/{current_path}:/children"
                )
                resp = await client.post(url, headers=headers, json={
                    "name": part,
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "fail"
                })
                # 201 = criada, 409 = já existia — ambos são sucesso
                if resp.status_code not in [201, 409]:
                    logger.error(f"Falha ao criar pasta '{part}': {resp.status_code} - {resp.text}")
                    raise HTTPException(
                        status_code=resp.status_code,
                        detail=f"Falha ao criar pasta '{part}' no SharePoint: {resp.text}"
                    )
                current_path = f"{current_path}/{part}" if current_path else part

        logger.info(f"Estrutura de pastas '{folder_path}' garantida no SharePoint.")

    async def _upload_content(self, drive_id: str, content: bytes, filename: str, target_path: str, headers: dict):
        """Lógica central de upload — compartilhada por upload_file e upload_file_bytes."""
        final_path = f"{target_path}/{filename}" if target_path else filename
        url = f"{self.base_url}/drives/{drive_id}/root:/{final_path}:/content"
        logger.debug(f"URL de upload: {url}")

        async with httpx.AsyncClient(http2=True, timeout=60.0) as client:
            resp = await client.put(url, headers=headers, data=content)

            if resp.status_code == 404:
                logger.warning(f"Pasta '{target_path}' não encontrada. Criando estrutura e repetindo upload...")
                await self._ensure_folder_path_exists(drive_id, target_path, headers)
                resp = await client.put(url, headers=headers, data=content)

            resp.raise_for_status()
            logger.info(f"Upload de '{filename}' concluído com sucesso.")
            return {"success": True, "message": "Upload concluído", "data": resp.json()}

    async def upload_file(self, site_name: str, drive_name: str, local_file_path: str, target_path: str):
        if not os.path.isabs(local_file_path):
            local_file_path = os.path.join(str(settings.paths.base), local_file_path)

        logger.info(f"Upload (servidor): Arquivo='{local_file_path}', Destino='{site_name}/{drive_name}/{target_path}'")
        if not os.path.exists(local_file_path):
            logger.error(f"Arquivo local para upload não encontrado: {local_file_path}")
            raise HTTPException(status_code=400, detail=f"Arquivo local não encontrado: {local_file_path}")

        try:
            headers = await self._get_headers('sharepoint')
            site_id = await self._get_site_id(site_name, headers)
            drive_id = await self._get_drive_id(site_id, drive_name, headers)

            filename = os.path.basename(local_file_path)
            async with aiofiles.open(local_file_path, 'rb') as f:
                content = await f.read()

            return await self._upload_content(drive_id, content, filename, target_path, headers)
        except httpx.HTTPStatusError as e:
            logger.error(f"Erro HTTP durante o upload: {e.response.text}", exc_info=True)
            raise HTTPException(status_code=e.response.status_code, detail=e.response.text)
        except Exception as e:
            logger.error(f"Erro inesperado durante o upload: {e}", exc_info=True)
            if isinstance(e, HTTPException):
                raise e
            raise HTTPException(status_code=500, detail=str(e))

    async def upload_file_bytes(self, site_name: str, drive_name: str, file_content_b64: str, filename: str, target_path: str):
        logger.info(f"Upload (cliente): Arquivo='{filename}', Destino='{site_name}/{drive_name}/{target_path}'")

        try:
            content = base64.b64decode(file_content_b64)
        except Exception:
            raise HTTPException(status_code=400, detail="Conteúdo 'file_content_b64' não é um Base64 válido.")

        try:
            headers = await self._get_headers('sharepoint')
            site_id = await self._get_site_id(site_name, headers)
            drive_id = await self._get_drive_id(site_id, drive_name, headers)

            return await self._upload_content(drive_id, content, filename, target_path, headers)
        except httpx.HTTPStatusError as e:
            logger.error(f"Erro HTTP durante o upload: {e.response.text}", exc_info=True)
            raise HTTPException(status_code=e.response.status_code, detail=e.response.text)
        except Exception as e:
            logger.error(f"Erro inesperado durante o upload: {e}", exc_info=True)
            if isinstance(e, HTTPException):
                raise e
            raise HTTPException(status_code=500, detail=str(e))

    async def send_email(self, sender: str, recipients: list, subject: str, body: str, 
                         attachment_folder: str = None, is_html: bool = True, 
                         attachments_data: list = None): # <--- Anexo do e-mail
        
        headers = await self._get_headers('email')
        
        # Estrutura base da mensagem
        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body
            },
            "toRecipients": [{"emailAddress": {"address": email}} for email in recipients],
            "attachments": []
        }

        # 1. Processa anexos de pasta LOCAL do servidor (Legado)
        if attachment_folder and os.path.exists(attachment_folder):
            for filename in os.listdir(attachment_folder):
                file_path = os.path.join(attachment_folder, filename)
                if os.path.isfile(file_path):
                    async with aiofiles.open(file_path, "rb") as f:
                        content = await f.read()
                        b64_content = base64.b64encode(content).decode("utf-8")
                        message["attachments"].append({
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": filename,
                            "contentBytes": b64_content
                        })

        # 2. Processa anexos enviados pelo CLIENTE (Novo)
        if attachments_data:
            for att in attachments_data:
                # O Pydantic já entrega o objeto, basta acessar os atributos
                # Se att for dict (depende de como chama), acesse via ['key']
                # Aqui assumindo objeto Pydantic ou dict vindo do controller
                name = att.name if hasattr(att, 'name') else att.get('name')
                content = att.content_b64 if hasattr(att, 'content_b64') else att.get('content_b64')
                
                message["attachments"].append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": name,
                    "contentBytes": content 
                })

        # Valida se o Sender tem permissão
        url = f"{self.base_url}/users/{sender}/sendMail"
        
        payload = {"message": message, "saveToSentItems": "true"}

        async with httpx.AsyncClient(http2=True) as client:
            resp = await client.post(url, headers=headers, json=payload)
            if resp.status_code != 202:
                 raise HTTPException(status_code=resp.status_code, detail=f"Erro envio email: {resp.text}")
            
            return {"success": True, "message": "Email enviado com sucesso"}
    
    async def _get_drive_id(self, site_id: str, drive_name: str, headers: dict) -> str:
        url = f"{self.base_url}/sites/{site_id}/drives"
        logger.debug(f"Buscando ID do drive: '{drive_name}'")
        async with httpx.AsyncClient(http2=True) as client:
            resp = await client.get(url, headers=headers)
            if resp.status_code != 200:
                logger.error(f"Falha ao buscar drives. Status: {resp.status_code}, Detalhe: {resp.text}")
                raise HTTPException(status_code=resp.status_code, detail="Não foi possível listar os drives do site.")
            
            drives = resp.json().get("value", [])
            drive = next((d for d in drives if d.get("name") == drive_name), None)
            
            if not drive:
                logger.error(f"Drive '{drive_name}' não encontrado na lista de drives.")
                raise HTTPException(status_code=404, detail=f"O Drive (biblioteca) '{drive_name}' não foi encontrado no site.")
            
            drive_id = drive["id"]
            logger.info(f"ID do Drive '{drive_name}' encontrado: {drive_id}")
            return drive_id
    
    async def _get_root_folder_id(self, drive_id: str, headers: dict) -> str:
        return await self._get_item_id(drive_id, "", headers)

    async def _get_item_id(self, drive_id: str, path: str, headers: dict) -> str:
        if not path or path == "/":
            logger.debug("Buscando ID do item: root")
            url = f"{self.base_url}/drives/{drive_id}/root"
        else:
            logger.debug(f"Buscando ID do item: '{path}'")
            url = f"{self.base_url}/drives/{drive_id}/root:/{path}"
        
        async with httpx.AsyncClient(http2=True) as client:
            resp = await client.get(url, headers=headers)
            if resp.status_code == 200:
                item_id = resp.json().get("id")
                logger.info(f"ID do item '{path}' encontrado: {item_id}")
                return item_id
            elif resp.status_code == 404:
                logger.warning(f"Item não encontrado no caminho: '{path}'")
                return ''
            else:
                logger.error(f"Erro ao buscar ID do item '{path}'. Status: {resp.status_code}, Detalhe: {resp.text}")
                resp.raise_for_status()
                return ''
    
    async def get_file_stream(self, site_name: str, drive_name: str, file_path: str):
        logger.info(f"Iniciando download de stream para: '{site_name}/{drive_name}/{file_path}'")
        client = httpx.AsyncClient(http2=True, follow_redirects=True, timeout=60.0)
        
        try:
            headers = await self._get_headers('sharepoint')
            site_id = await self._get_site_id(site_name, headers)
            drive_id = await self._get_drive_id(site_id, drive_name, headers)
            
            url = f"{self.base_url}/drives/{drive_id}/root:/{file_path}:/content"
            logger.debug(f"URL de download do stream: {url}")

            req = client.build_request("GET", url, headers=headers)
            response = await client.send(req, stream=True)

            if response.status_code != 200:
                error_content = await response.aread()
                logger.error(f"Falha ao iniciar stream. Status: {response.status_code}. Detalhe: {error_content.decode('utf-8')}")
                await client.aclose()
                raise HTTPException(status_code=response.status_code, detail=f"Microsoft Graph Error: {error_content.decode('utf-8')}")

            logger.info(f"Stream para '{file_path}' estabelecido com sucesso.")
            
            async def stream_generator():
                try:
                    async for chunk in response.aiter_bytes(chunk_size=16384):
                        yield chunk
                    logger.debug(f"Stream para '{file_path}' finalizado.")
                except Exception as e:
                    logger.error(f"Erro durante o streaming do arquivo '{file_path}': {e}", exc_info=True)
                finally:
                    await response.aclose()
                    await client.aclose()
                    logger.info(f"Conexão de stream para '{file_path}' fechada.")

            return stream_generator

        except Exception as e:
            logger.critical(f"Erro crítico antes de iniciar o stream para '{file_path}': {e}", exc_info=True)
            await client.aclose()
            if isinstance(e, HTTPException):
                raise e
            raise HTTPException(status_code=502, detail=f"Erro de conexão com Microsoft ao tentar baixar: {str(e)}")

    async def move_file(self, site_name: str, drive_name: str, ori_path: str, dest_folder_path: str, conflict_behavior: str = "fail"):
        headers = await self._get_headers('sharepoint')
        site_id = await self._get_site_id(site_name, headers)
        drive_id = await self._get_drive_id(site_id, drive_name, headers)

        # 1. Pega ID do arquivo original
        file_id = await self._get_item_id(drive_id, ori_path, headers)
        if not file_id:
            raise HTTPException(status_code=404, detail=f"Arquivo de origem não encontrado: {ori_path}")

        # 2. Pega ID da pasta de destino (O Graph exige o ID do PARENT para mover)
        # Se dest_folder_path for vazio, assume raiz
        if not dest_folder_path:
            dest_folder_id = await self._get_root_folder_id(drive_id, headers)
        else:
            dest_folder_id = await self._get_item_id(drive_id, dest_folder_path, headers)
            if not dest_folder_id:
                # Opcional: Criar pasta se não existir (requer mais lógica), aqui vamos dar erro
                raise HTTPException(status_code=404, detail=f"Pasta de destino não encontrada: {dest_folder_path}")

        url = f"{self.base_url}/drives/{drive_id}/items/{file_id}"

        # 3. Payload com a instrução de conflito
        payload = {
            "parentReference": {"id": dest_folder_id},
            "name": os.path.basename(ori_path), # Mantém o mesmo nome
            "@microsoft.graph.conflictBehavior": conflict_behavior 
        }

        async with httpx.AsyncClient(http2=True) as client:
            # Para mover, usamos PATCH
            resp = await client.patch(url, headers=headers, json=payload)
            
            if resp.status_code not in [200, 201]:
                # Se for "fail" e já existir, vai cair aqui com erro 409 (Conflict)
                raise HTTPException(status_code=resp.status_code, detail=f"Erro ao mover arquivo: {resp.text}")
            
            # Retorna o novo nome (útil no caso de rename)
            novo_nome = resp.json().get("name")
            return {
                "success": True, 
                "message": "Arquivo movido com sucesso", 
                "new_name": novo_nome,
                "behavior_used": conflict_behavior
            }
        
    async def read_excel_values(self, site_name: str, drive_name: str, file_path: str, sheet_name: str):
        headers = await self._get_headers('sharepoint')
        site_id = await self._get_site_id(site_name, headers)
        drive_id = await self._get_drive_id(site_id, drive_name, headers)
        
        # Obtém ID do arquivo
        item_id = await self._get_item_id(drive_id, file_path, headers)
        if not item_id:
            raise HTTPException(status_code=404, detail=f"Arquivo excel não encontrado: {file_path}")

        # Endpoint: usedRange (Pega apenas a área com dados)
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_name}/usedRange"

        async with httpx.AsyncClient(http2=True, timeout=60.0) as client:
            resp = await client.get(url, headers=headers)

            if resp.status_code != 200:
                raise HTTPException(status_code=resp.status_code, detail=f"Erro ao ler Excel: {resp.text}")

            # Retorna apenas a matriz de valores
            return {"values": resp.json().get('values')}

    async def write_excel_cells(self, site_name: str, drive_name: str, file_path: str, sheet_name: str, values_dict: dict):
        headers = await self._get_headers('sharepoint')
        site_id = await self._get_site_id(site_name, headers)
        drive_id = await self._get_drive_id(site_id, drive_name, headers)
        
        item_id = await self._get_item_id(drive_id, file_path, headers)
        if not item_id:
            raise HTTPException(status_code=404, detail=f"Arquivo excel não encontrado: {file_path}")

        base_url = f"{self.base_url}/drives/{drive_id}/items/{item_id}/workbook/worksheets/{sheet_name}"
        
        results = []
        
        # Abre uma sessão única para processar as escritas (mais rápido)
        async with httpx.AsyncClient(http2=True, timeout=30.0) as client:
            for cell, value in values_dict.items():
                url = f"{base_url}/range(address='{cell}')"
                body = {"values": [[value]]} # Formato exigido pela MS: Matriz 2D
                
                success = False
                last_error = ""

                # Retry logic simples (transcrita do seu código original)
                for attempt in range(3): 
                    try:
                        resp = await client.patch(url, headers=headers, json=body)
                        if resp.status_code == 200:
                            success = True
                            break
                        else:
                            last_error = resp.text
                            await asyncio.sleep(1) # Backoff simples
                    except Exception as e:
                        last_error = str(e)
                        await asyncio.sleep(1)
                
                results.append({
                    "cell": cell, 
                    "success": success, 
                    "error": last_error if not success else None
                })
        
        return {"summary": results}

    