import os
from functools import lru_cache
from azure.identity import ChainedTokenCredential, EnvironmentCredential, AzureCliCredential
from azure.keyvault.secrets import SecretClient
from dotenv import load_dotenv

load_dotenv(override=True)


class VaultClient:
    def __init__(self):
        self.vault_uri = os.getenv("KEY_VAULT_URL")
        if not self.vault_uri:
            raise ValueError("A variável de ambiente 'KEY_VAULT_URL' é obrigatória.")

        # Cadeia explícita e previsível:
        # - Produção (NSSM): usa AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET (EnvironmentCredential)
        # - Desenvolvimento: usa token do 'az login' (AzureCliCredential)
        self.credential = ChainedTokenCredential(
            EnvironmentCredential(),
            AzureCliCredential(),
        )
        self.client = SecretClient(vault_url=self.vault_uri, credential=self.credential)

    @lru_cache(maxsize=64)
    def get_secret(self, secret_name: str) -> str:
        """
        Busca o segredo no Azure e armazena em cache na memória.
        """
        try:
            # print(f"[DEBUG] Fetching '{secret_name}' from Key Vault...") 
            secret = self.client.get_secret(secret_name)
            return secret.value.strip() if secret.value else secret.value
        except Exception as e:
            raise RuntimeError(f"Falha ao recuperar segredo '{secret_name}': {e}")

# Instância Singleton para ser importada em outros módulos
# Garante que a conexão e o cache sejam reutilizados
vault_client = VaultClient()