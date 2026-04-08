
import pandas as pd


class StreamlitCoreAdapter:
    def __init__(self, language='tr', notes=None, responsibles=None):
        self.language = language
        self.notes = notes or []
        self.responsibles = responsibles or []
        self.tb_raw_df = None
        self.plcc_raw_df = None
        self.available_periods = []
        self.current_period = ''
        self.previous_period = ''

        self.regular_ft_faggl_df = pd.DataFrame()
        self.regular_ft_eba_df = pd.DataFrame()
        self.regular_ft_zfi052_df = pd.DataFrame()
        self.regular_ft_output_df = pd.DataFrame()
        self.regular_ft_base_output_df = pd.DataFrame()
        self.regular_ft_periods = []
        self.regular_ft_current_period = ''
        self.regular_ft_previous_period = ''
        self.regular_ft_user_period_map = {}
        self.regular_ft_eba_pending_period_dict = {}
        self.regular_ft_eba_invoice_period_dict = {}
        self.regular_ft_faggl_invoice_keys_map = {}
        self.regular_ft_zfi_invoice_map = {}
        self.regular_ft_zfi_period_map = {}
        self.regular_ft_zfi_vendor_map = {}

    def set_processing_message(self, message: str, progress: int = 0):
        return None

    def regular_ft_all_label(self) -> str:
        return 'Tümü' if self.language == 'tr' else 'All'

    def recalculate_regular_ft_current_fields(self):
        if self.regular_ft_base_output_df is None or self.regular_ft_base_output_df.empty:
            self.regular_ft_output_df = pd.DataFrame()
            return
        self.regular_ft_output_df = self.regular_ft_base_output_df.copy()


class MuavinAdapter:
    def __init__(self, language='tr'):
        self.language = language
