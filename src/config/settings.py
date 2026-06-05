import yaml
import os
import socket
from pathlib import Path
from typing import Union, List
from pydantic import BaseModel, field_validator, model_validator, Field  # field_validator usado em PathConfigBase
from src.core.vault_client import VaultClient
from dotenv import load_dotenv

load_dotenv(override=True)
PROJETO_RAIZ = Path(__file__).resolve().parents[2]

# 1. Instanciar o VaultClient (Singleton) uma única vez
try:
    vault = VaultClient()
except Exception as e:
    print(f"AVISO: Não foi possível conectar ao Key Vault: {e}")
    vault = None

def _get_ip() -> str:
    try:
        nome_host = socket.gethostname()
        endereco_ip = socket.gethostbyname(nome_host)
        return str(endereco_ip)
    except Exception:
        return "127.0.0.1"

class AppSettings(BaseModel):
    api_url: str = Field(default_factory=_get_ip)
    api_port: int = 8001
    api_producao: bool = False
    allowed_origins: List[str] = ["*"]
    api_key: str = "API-KEY-MIDDLEWARE"

    @model_validator(mode="before")
    @classmethod
    def load_vault_secrets(cls, data):
        if not vault or not isinstance(data, dict):
            return data
        if "api_key" not in data:
            data["api_key"] = vault.get_secret("API-KEY-MIDDLEWARE")
        return data

class AzureSettings(BaseModel):
    tenant_name: str = "seu-tenant-padrao.onmicrosoft.com"
    tenant_id: str = "TENANT-ID"
    client_id_sharepoint: str = "CLIENT-ID-SHAREPOINT"
    client_secret_sharepoint: str = "CLIENT-SECRET-SHAREPOINT"
    client_id_email: str = "CLIENT-ID-EMAIL"
    client_secret_email: str = "CLIENT-SECRET-EMAIL"

    @model_validator(mode="before")
    @classmethod
    def load_vault_secrets(cls, data):
        if not vault or not isinstance(data, dict):
            return data
        secret_fields = {
            "tenant_id": "TENANT-ID",
            "client_id_sharepoint": "CLIENT-ID-SHAREPOINT",
            "client_secret_sharepoint": "CLIENT-SECRET-SHAREPOINT",
            "client_id_email": "CLIENT-ID-EMAIL",
            "client_secret_email": "CLIENT-SECRET-EMAIL",
        }
        for field_name, secret_name in secret_fields.items():
            if field_name not in data:
                data[field_name] = vault.get_secret(secret_name)
        return data

class PathConfigBase(BaseModel):
    """
    Converte caminhos relativos em absolutos baseados na raiz do projeto.
    """
    @field_validator('*', mode='after')
    @classmethod
    def make_absolute_path(cls, v):
        if isinstance(v, Path) and not v.is_absolute():
            return PROJETO_RAIZ / v
        return v

class DataPaths(PathConfigBase):
    input: Path = Path("input_data")
    output: Path = Path("data/output")
    backup: Path = Path("data/backup")
    results: Path = Path("data/results")

class SrcPaths(PathConfigBase):
    config: Path = Path("src/config")
    services: Path = Path("src/services")
    utils: Path = Path("src/utils")
    tests: Path = Path("src/tests")
    workflows: Path = Path("src/workflows")

class MainPaths(PathConfigBase):
    base: Union[Path, str] = PROJETO_RAIZ
    logs: Path = Path("logs")
    logs_print: Path = Path("logs/print")
    data: DataPaths = Field(default_factory=DataPaths)
    src: SrcPaths = Field(default_factory=SrcPaths)

class Settings(BaseModel):
    app: AppSettings
    azure: AzureSettings
    paths: MainPaths = Field(default_factory=MainPaths)

def load_settings() -> Settings:
    config_path = PROJETO_RAIZ / "config.yaml"
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        # O Pydantic vai ignorar campos extras se existirem no YAML, 
        # mas idealmente você deve removê-los do arquivo.
        return Settings(**data)
    except FileNotFoundError:
        raise RuntimeError(f"Arquivo de configuração não encontrado: {config_path}")
    except Exception as e:
        raise RuntimeError(f"Erro ao carregar config.yaml: {e}")

settings = load_settings()