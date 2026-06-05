from fastapi import FastAPI, Depends, HTTPException, Request, File, UploadFile, Form
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
from pydantic import BaseModel, Field, model_validator, field_validator
from typing import List, Optional, Dict, Any, Literal
import uvicorn
import base64
import os
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded
from src.config.settings import settings
from src.core.graph_client import GraphService
from src.config.security import get_api_key
from src.utils.logger import Logger


MAX_UPLOAD_BYTES = 25 * 1024 * 1024  # 25 MB

def _validate_sp_path(v: str) -> str:
    """Rejeita caminhos com traversal (../) ou caracteres nulos."""
    if v and (".." in v or "\x00" in v):
        raise ValueError("Caminho inválido: '..' e caracteres nulos não são permitidos.")
    return v

limiter = Limiter(key_func=get_remote_address)

# ==============================================================================
# Configurações Globais & Lifespan
# ==============================================================================

# Logger
logger = Logger().get_logger()
service = GraphService() # Singleton

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Executado na inicialização
    IP = settings.app.api_url
    PORT = settings.app.api_port
    logger.info(f"Servidor iniciando em http://{IP}:{PORT} | Produção: {settings.app.api_producao}")
    yield
    # Executado no desligamento
    logger.info("Servidor desligando...")

def validar_segredos_azure():
    """
    Testa a conexão com o Key Vault no início da aplicação.
    Se falhar, loga o erro e encerra, evitando falhas silenciosas.
    """
    logger.info("Inicializando conexão com Azure Key Vault...")
    
    try:
        # Acessar uma propriedade força a chamada ao vault.get_secret()
        # Testamos o Tenant ID pois é uma chave 'segura' de testar
        _ = settings.azure.tenant_id
        logger.info("Conexão com Key Vault estabelecida e credenciais validadas.")
        
    except RuntimeError as e:
        logger.critical(f"ERRO CRÍTICO DE CONFIGURAÇÃO: {e}")
        # Opcional: raise e (para parar o robô) ou sys.exit(1)
    except PermissionError as e:
        logger.critical(f"ERRO DE PERMISSÃO AZURE: {e}")
    except FileNotFoundError as e:
        logger.error(f"SEGREDO NÃO ENCONTRADO: {e}")
    except Exception as e:
        logger.error(f"Erro desconhecido no Vault: {e}")

# --- INICIO DO APP ---
validar_segredos_azure()

app = FastAPI(
    title="Middleware Microsoft Graph API",
    description="API Gateway para centralizar automações de Email e SharePoint via Microsoft Graph API.",
    version="2.0.0",
    lifespan=lifespan
)

# CORS (Recomendado se for acessado via browser/n8n web)
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.app.allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# ==============================================================================
# Global Exception Handlers (DRY - Limpeza de Código)
# ==============================================================================
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    """Captura qualquer erro não tratado nos endpoints."""
    logger.error(f"Erro não tratado em {request.url}: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"detail": "Erro interno do servidor. Consulte os logs."}
    )

@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    """Captura erros HTTP explícitos e loga antes de retornar."""
    logger.warning(f"Erro HTTP {exc.status_code} em {request.url}: {exc.detail}")
    return JSONResponse(
        status_code=exc.status_code,
        content={"detail": exc.detail}
    )
# ==============================================================================
# Modelos de Dados (DTOs)
# ==============================================================================

# --- EMAIL DTOs ---

class AttachmentItem(BaseModel):
    name: str = Field(..., description="Nome do arquivo com extensão", examples=["relatorio.pdf"])
    content_b64: str = Field(..., description="Conteúdo do arquivo em Base64")
    content_type: str = Field("application/octet-stream", description="MimeType do arquivo")

class EmailRequest(BaseModel):
    sender: str = Field(
        ...,
        description="Email do remetente (deve ter permissão no Azure)",
        examples=["automacao@suaempresa.com"]
    )
    recipients: List[str] = Field(
        ...,
        description="Lista de emails destinatários",
        examples=[["usuario1@suaempresa.com", "usuario2@suaempresa.com"]]
    )
    subject: str = Field(
        ...,
        description="Assunto do email",
        examples=["Relatório Diário"]
    )
    body: str = Field(
        ...,
        description="Corpo do email (HTML ou Texto)",
        examples=["<h1>Olá, segue anexo.</h1>"]
    )
    is_html: bool = Field(
        True,
        description="Se True, envia como HTML. Se False, texto puro."
    )
    attachment_folder: Optional[str] = None
    attachments_data: Optional[List[AttachmentItem]] = Field(
        None,
        description="Lista de arquivos enviados pelo cliente (Base64)"
    )

    @field_validator("attachment_folder")
    @classmethod
    def check_attachment_folder(cls, v): return _validate_sp_path(v) if v else v

# --- SHAREPOINT DTOs ---
class SharePointBase(BaseModel):
    site_name: str = Field(
        ...,
        description="Nome do site no SharePoint",
        examples=["TI-Industrial"]
    )
    drive_name: str = Field(
        ...,
        description="Nome da biblioteca de documentos",
        examples=["Documentos"]
    )

class SPListFilesRequest(SharePointBase):
    folder_path: str = Field(
        "",
        description="Caminho da pasta para listar (vazio = raiz)",
        examples=["Relatorios/2026"]
    )

    @field_validator("folder_path")
    @classmethod
    def check_folder_path(cls, v): return _validate_sp_path(v)

class SPDownloadRequest(SharePointBase):
    file_path: str = Field(
        ...,
        description="Caminho do arquivo no SharePoint a ser baixado",
        examples=["Relatorios/2024/janeiro.xlsx"]
    )

    @field_validator("file_path")
    @classmethod
    def check_file_path(cls, v): return _validate_sp_path(v)

class SPUploadRequest(SharePointBase):
    local_file_path: Optional[str] = Field(
        None,
        description="Caminho do arquivo no servidor (use quando o arquivo já está na máquina do middleware)",
        examples=["/app/uploads/arquivo.pdf"]
    )
    file_content_b64: Optional[str] = Field(
        None,
        description="Conteúdo do arquivo em Base64 (use quando o arquivo está na máquina do cliente/RPA)"
    )
    filename: Optional[str] = Field(
        None,
        description="Nome do arquivo com extensão (obrigatório com file_content_b64)",
        examples=["relatorio.xlsx"]
    )
    target_path: str = Field(
        "",
        description="Pasta de destino no SharePoint",
        examples=["Financeiro/Notas"]
    )

    @field_validator("local_file_path", "target_path")
    @classmethod
    def check_paths(cls, v): return _validate_sp_path(v) if v else v

    @field_validator("file_content_b64")
    @classmethod
    def check_content_size(cls, v):
        if v and len(v) > MAX_UPLOAD_BYTES * 4 // 3:
            raise ValueError(f"Conteúdo excede o limite de 25 MB.")
        return v

    @model_validator(mode="after")
    def validate_source(self):
        if not self.local_file_path and not self.file_content_b64:
            raise ValueError("Forneça 'local_file_path' (arquivo no servidor) ou 'file_content_b64' (arquivo do cliente).")
        if self.file_content_b64 and not self.filename:
            raise ValueError("'filename' é obrigatório quando 'file_content_b64' é fornecido.")
        return self

class SPMoveRequest(SharePointBase):
    ori_path: str = Field(
        ...,
        description="Caminho atual do arquivo no SharePoint",
        examples=["Entrada/arquivo.pdf"]
    )
    dest_folder_path: str = Field(
        ...,
        description="Pasta de destino no SharePoint",
        examples=["Processados"]
    )
    conflict_behavior: Literal["fail", "replace", "rename"] = Field(
        "fail",
        description="O que fazer se o arquivo já existir? 'fail' (erro), 'replace' (substitui), 'rename' (novo nome)."
    )

    @field_validator("ori_path", "dest_folder_path")
    @classmethod
    def check_move_paths(cls, v): return _validate_sp_path(v)

class SPExcelReadRequest(SharePointBase):
    file_path: str = Field(
        ...,
        description="Caminho do arquivo Excel",
        examples=["Relatorios/financeiro.xlsx"]
    )
    sheet_name: str = Field(
        ...,
        description="Nome da aba (Worksheet)",
        examples=["Planilha1"]
    )

    @field_validator("file_path")
    @classmethod
    def check_file_path(cls, v): return _validate_sp_path(v)

class SPExcelWriteRequest(SharePointBase):
    file_path: str = Field(..., description="Caminho do arquivo Excel")
    sheet_name: str = Field(..., description="Nome da aba")
    values: Dict[str, Any] = Field(
        ...,
        description="Dicionário {Célula: Valor}",
        examples=[{"A1": "ID", "B2": 500.50, "C3": "Aprovado"}]
    )

    @field_validator("file_path")
    @classmethod
    def check_file_path(cls, v): return _validate_sp_path(v)

# ==============================================================================
# Endpoints
# ==============================================================================

@app.get("/health", tags=["Monitoramento"])
async def health_check():
    return {"status": "online", "service": "Graph Middleware"}

@app.post("/email/send", tags=["Email"], dependencies=[Depends(get_api_key)])
@limiter.limit("20/minute")
async def send_email_endpoint(request: Request, payload: EmailRequest):
    logger.info(f"Envio de email: {payload.sender} -> {payload.recipients}")
    
    # Se houver anexos vindos do cliente, loga a quantidade
    if payload.attachments_data:
        logger.info(f"Anexando {len(payload.attachments_data)} arquivos enviados pelo cliente.")

    return await service.send_email(
        sender=payload.sender,
        recipients=payload.recipients,
        subject=payload.subject,
        body=payload.body,
        attachment_folder=payload.attachment_folder,
        is_html=payload.is_html,
        attachments_data=payload.attachments_data 
    )

@app.post("/email/send-files", tags=["Email"], dependencies=[Depends(get_api_key)])
@limiter.limit("10/minute")
async def send_email_with_files_endpoint(
    request: Request,
    sender: str = Form(..., description="Email do remetente"),
    recipients: List[str] = Form(..., description="Lista de destinatários"),
    subject: str = Form(..., description="Assunto do email"),
    body: str = Form(..., description="Corpo do email (HTML ou texto)"),
    is_html: bool = Form(True, description="Se True, envia como HTML"),
    attachments: List[UploadFile] = File(default=[], description="Arquivos a anexar")
):
    """
    Envia email com arquivos anexados diretamente via multipart/form-data.
    O cliente envia os arquivos binários normalmente — sem necessidade de codificar em Base64.
    O servidor faz a conversão internamente antes de repassar ao Microsoft Graph.
    """
    logger.info(f"Envio de email com upload: {sender} -> {recipients} | {len(attachments)} anexo(s)")

    attachments_data = []
    for file in attachments:
        content = await file.read()
        if len(content) > MAX_UPLOAD_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"Arquivo '{file.filename}' excede o limite de 25 MB."
            )
        attachments_data.append(AttachmentItem(
            name=file.filename,
            content_b64=base64.b64encode(content).decode("utf-8"),
            content_type=file.content_type or "application/octet-stream"
        ))

    return await service.send_email(
        sender=sender,
        recipients=recipients,
        subject=subject,
        body=body,
        is_html=is_html,
        attachments_data=attachments_data if attachments_data else None
    )

@app.post("/sharepoint/list", tags=["SharePoint"], dependencies=[Depends(get_api_key)])
@limiter.limit("60/minute")
async def list_files_endpoint(request: Request, payload: SPListFilesRequest):
    logger.info(f"Listar: {payload.drive_name}/{payload.folder_path}")
    return await service.list_files(
        site_name=payload.site_name,
        drive_name=payload.drive_name,
        folder_path=payload.folder_path
    )

@app.post("/sharepoint/download", tags=["SharePoint"], dependencies=[Depends(get_api_key)])
@limiter.limit("60/minute")
async def download_file_endpoint(request: Request, payload: SPDownloadRequest):
    logger.info(f"Download solicitado: {payload.file_path}")
    
    file_iterator = await service.get_file_stream(
        site_name=payload.site_name,
        drive_name=payload.drive_name,
        file_path=payload.file_path
    )
    
    filename = os.path.basename(payload.file_path)
    
    return StreamingResponse(
        file_iterator(),
        media_type="application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

@app.post("/sharepoint/upload", tags=["SharePoint"], dependencies=[Depends(get_api_key)])
@limiter.limit("30/minute")
async def upload_file_endpoint(request: Request, payload: SPUploadRequest):
    if payload.file_content_b64:
        logger.info(f"Upload (cliente): {payload.filename} -> {payload.target_path}")
        return await service.upload_file_bytes(
            site_name=payload.site_name,
            drive_name=payload.drive_name,
            file_content_b64=payload.file_content_b64,
            filename=payload.filename,
            target_path=payload.target_path
        )
    logger.info(f"Upload (servidor): {payload.local_file_path} -> {payload.target_path}")
    return await service.upload_file(
        site_name=payload.site_name,
        drive_name=payload.drive_name,
        local_file_path=payload.local_file_path,
        target_path=payload.target_path
    )

@app.post("/sharepoint/move", tags=["SharePoint"], dependencies=[Depends(get_api_key)])
@limiter.limit("30/minute")
async def move_file_endpoint(request: Request, payload: SPMoveRequest):
    logger.info(f"Mover: {payload.ori_path} -> {payload.dest_folder_path} (Modo: {payload.conflict_behavior})")
    
    return await service.move_file(
        site_name=payload.site_name,
        drive_name=payload.drive_name,
        ori_path=payload.ori_path,
        dest_folder_path=payload.dest_folder_path,
        conflict_behavior=payload.conflict_behavior # Passando o parametro
    )

@app.post("/sharepoint/excel/read", tags=["Excel"], dependencies=[Depends(get_api_key)])
@limiter.limit("60/minute")
async def read_excel_endpoint(request: Request, payload: SPExcelReadRequest):
    """Lê todos os valores usados de uma aba do Excel."""
    logger.info(f"Lendo Excel: {payload.file_path} (Aba: {payload.sheet_name})")
    return await service.read_excel_values(
        site_name=payload.site_name,
        drive_name=payload.drive_name,
        file_path=payload.file_path,
        sheet_name=payload.sheet_name
    )

@app.post("/sharepoint/excel/write", tags=["Excel"], dependencies=[Depends(get_api_key)])
@limiter.limit("30/minute")
async def write_excel_endpoint(request: Request, payload: SPExcelWriteRequest):
    """Escreve valores em células específicas do Excel."""
    logger.info(f"Escrevendo no Excel: {payload.file_path} - {len(payload.values)} células")
    return await service.write_excel_cells(
        site_name=payload.site_name,
        drive_name=payload.drive_name,
        file_path=payload.file_path,
        sheet_name=payload.sheet_name,
        values_dict=payload.values
    )

if __name__ == "__main__":
    # Reload deve ser controlado pelo config, não hardcoded no 'not'
    uvicorn.run(
        "main:app", 
        host=settings.app.api_url, 
        port=settings.app.api_port, 
        reload=not settings.app.api_producao
    )