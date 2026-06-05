import secrets
from fastapi import Security, HTTPException, status
from fastapi.security import APIKeyHeader
from src.config.settings import settings
from src.utils.logger import Logger

logger = Logger().get_logger()

api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)

async def get_api_key(api_key: str = Security(api_key_header)):
    """Valida se a chave enviada no Header bate com o segredo do Key Vault."""
    if api_key and secrets.compare_digest(api_key.strip(), settings.app.api_key):
        logger.info("API Key validada com sucesso.")
        return api_key.strip()

    logger.warning("Tentativa de acesso com X-API-Key inválida ou ausente.")
    raise HTTPException(
        status_code=status.HTTP_403_FORBIDDEN,
        detail="Credenciais inválidas. Verifique o header 'X-API-Key'."
    )
