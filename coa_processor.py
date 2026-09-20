import os

class COAProcessor:
    def __init__(self, vault_dir="Cleaned_Vault_Folder"):
        self.vault_dir = vault_dir

    def find_coa_by_lot(self, lot_number):
        """Searches the vault directory for a PDF containing the given lot number in its filename."""
        if not os.path.exists(self.vault_dir):
            return None
            
        lot_str = str(lot_number).strip().lower()
        if not lot_str:
            return None

        for root, dirs, files in os.walk(self.vault_dir):
            for file in files:
                if file.lower().endswith(".pdf") and lot_str in file.lower():
                    return os.path.join(root, file)
        return None
