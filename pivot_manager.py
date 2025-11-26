import logging
import shutil
from pathlib import Path
from typing import Optional
from openpyxl import load_workbook

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(name)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("PivotManager")

class PivotTableConfigurator:
    """
    Gerencia configurações de metadados de Tabelas Dinâmicas em arquivos Excel (OpenXML).
    """

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)

    def _create_backup(self) -> None:
        if not self.file_path.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {self.file_path}")
            
        backup_path = self.file_path.with_suffix(f".backup{self.file_path.suffix}")
        shutil.copy2(self.file_path, backup_path)
        logger.info(f"Backup criado em: {backup_path}")

    def set_refresh_on_load(self, sheet_name: str, pivot_name: Optional[str] = None) -> bool:
        """
        Habilita a flag 'RefreshOnLoad' nas tabelas dinâmicas especificadas.
        """
        wb = None
        try:
            self._create_backup()
            
            logger.info(f"Processando arquivo: {self.file_path.name}")
            wb = load_workbook(self.file_path, keep_vba=True)
            
            if sheet_name not in wb.sheetnames:
                logger.error(f"Aba '{sheet_name}' não encontrada.")
                return False

            ws = wb[sheet_name]
            pivot_found = False

            for pivot in ws._pivots:
                if pivot_name and pivot.name != pivot_name:
                    continue
                
                pivot.cache.refreshOnLoad = True
                pivot_found = True
                logger.info(f"Configuração aplicada: {pivot.name}")

            if not pivot_found:
                logger.warning(f"Nenhuma tabela correspondente encontrada na aba {sheet_name}.")
                return False

            wb.save(self.file_path)
            logger.info("Alterações salvas com sucesso.")
            return True

        except Exception as e:
            logger.exception(f"Erro durante o processamento: {e}")
            return False
        finally:
            if wb:
                wb.close()

if __name__ == "__main__":
    ARQUIVO_ALVO = r"C:\Dados\Relatorio_Vendas.xlsx"
    NOME_ABA = "Geral"
    NOME_TABELA = "TabelaDinamica1"

    configurator = PivotTableConfigurator(ARQUIVO_ALVO)
    configurator.set_refresh_on_load(NOME_ABA, NOME_TABELA)
