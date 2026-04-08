
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from .adapters import MuavinAdapter, StreamlitCoreAdapter
from .legacy_loader import load_legacy_module
from . import storage


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def get_legacy():
    return load_legacy_module(str(project_root()))


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = 'Rapor') -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        (df if isinstance(df, pd.DataFrame) else pd.DataFrame()).to_excel(writer, sheet_name=sheet_name[:31] or 'Rapor', index=False)
    return output.getvalue()


def temp_paths_from_uploaded_files(uploaded_files: List) -> List[str]:
    paths = []
    for uploaded in uploaded_files or []:
        suffix = Path(uploaded.name).suffix or '.xlsx'
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded.getvalue())
            paths.append(tmp.name)
    return paths


def uploaded_file_to_df(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    return pd.read_excel(BytesIO(uploaded_file.getvalue()))


def run_main_analysis(tb_df: Optional[pd.DataFrame], plcc_df: Optional[pd.DataFrame], language: str = 'tr',
                      current_period: Optional[str] = None, previous_period: Optional[str] = None) -> Dict:
    legacy = get_legacy()
    notes = storage.load_notes(project_root())
    responsibles = storage.load_responsibles(project_root())

    adapter = StreamlitCoreAdapter(language=language, notes=notes, responsibles=responsibles)
    adapter.tb_raw_df = tb_df if isinstance(tb_df, pd.DataFrame) else pd.DataFrame()
    adapter.plcc_raw_df = plcc_df if isinstance(plcc_df, pd.DataFrame) else pd.DataFrame()

    adapter.available_periods = legacy.TbPlCcControlWindow.get_available_periods_from_loaded_files(adapter)
    selectable = [p for p in adapter.available_periods if p != legacy.OPENING_PERIOD_TR]
    if not selectable:
        selectable = ['02-Mart', '03-Nisan']

    adapter.current_period = current_period if current_period in selectable else selectable[-1]
    valid_previous = [p for p in selectable if p != adapter.current_period]
    adapter.previous_period = previous_period if previous_period in valid_previous else (valid_previous[-1] if valid_previous else adapter.current_period)

    tb_rows = legacy.TbPlCcControlWindow.build_tb_rows(adapter) if not adapter.tb_raw_df.empty else []
    plcc_detail, plcc_subtotal = legacy.TbPlCcControlWindow.build_plcc_rows(adapter) if not adapter.plcc_raw_df.empty else ([], [])
    balance_rows = legacy.build_balance_sheet_summary(tb_rows, adapter.available_periods, adapter.current_period, adapter.previous_period) if tb_rows else []
    income_rows = legacy.build_income_statement_summary(tb_rows, adapter.available_periods, adapter.current_period, adapter.previous_period) if tb_rows else []

    return {
        'adapter': adapter,
        'tb_rows': pd.DataFrame(tb_rows),
        'plcc_detail': pd.DataFrame(plcc_detail),
        'plcc_subtotal': pd.DataFrame(plcc_subtotal),
        'balance_df': pd.DataFrame(balance_rows),
        'income_df': pd.DataFrame(income_rows),
    }


def get_muavin_field_definitions(language: str = 'tr') -> List[Dict]:
    tr = language == 'tr'
    return [
        {'key': 'yilay', 'label': 'Yıl/ay' if tr else 'Year/Month', 'required': True},
        {'key': 'ana_hesap', 'label': 'Ana hesap' if tr else 'Main Account', 'required': True},
        {'key': 'ana_hesap_adi', 'label': 'DK hesabı uzun metni' if tr else 'G/L account long text', 'required': True},
        {'key': 'up_tutar', 'label': 'UPB Tutar' if tr else 'Amount in local currency', 'required': False},
        {'key': 'belge_pb_tutar', 'label': 'Belge PB Tutar' if tr else 'Amount in doc currency', 'required': False},
        {'key': 'belge_pb', 'label': 'Belge PB' if tr else 'Document Currency', 'required': False},
        {'key': 'referans', 'label': 'Referans' if tr else 'Reference', 'required': False},
        {'key': 'belge_numarasi', 'label': 'Belge numarası' if tr else 'Document Number', 'required': False},
        {'key': 'belge_turu', 'label': 'Belge türü' if tr else 'Document Type', 'required': False},
        {'key': 'karsi_hesap', 'label': 'Satıcı/Karşı Kayıt Hesabı' if tr else 'Vendor/Contra Account', 'required': False},
        {'key': 'karsi_hesap_tanimi', 'label': 'Satıcı/Karşı Kayıt Hesabı Tanımı' if tr else 'Vendor/Contra Account Name', 'required': False},
        {'key': 'denklestirme', 'label': 'Denkleştirme' if tr else 'Clearing Document', 'required': False},
        {'key': 'metin', 'label': 'Metin' if tr else 'Text', 'required': False},
        {'key': 'ters_kayit', 'label': 'Ters kayıt blg.no.' if tr else 'Reversal Document No.', 'required': False},
        {'key': 'kullanici', 'label': 'Kullanıcı adı' if tr else 'User Name', 'required': False},
        {'key': 'vergi_gostergesi', 'label': 'Vergi göstergesi' if tr else 'Tax Indicator', 'required': False},
        {'key': 'belge_tarihi', 'label': 'Belge Tarihi' if tr else 'Document Date', 'required': False},
        {'key': 'kayit_tarihi', 'label': 'Kayıt Tarihi' if tr else 'Posting Date', 'required': False},
        {'key': 'giris_tarihi', 'label': 'Giriş Tarihi' if tr else 'Created Date', 'required': False},
        {'key': 'masraf_yeri', 'label': 'Masraf yeri' if tr else 'Cost Center', 'required': False},
        {'key': 'masraf_yeri_tanimi', 'label': 'Masraf yeri tanımı' if tr else 'Cost Center Description', 'required': False},
    ]


def load_muavin_headers(uploaded_files: List, language: str = 'tr') -> Dict:
    legacy = get_legacy()
    paths = temp_paths_from_uploaded_files(uploaded_files)
    return legacy.load_muavin_headers_payload(paths=paths, language=language)


def _muavin_tables_from_clean_df(clean_df: pd.DataFrame, language: str = 'tr') -> Dict[str, pd.DataFrame]:
    legacy = get_legacy()
    adapter = MuavinAdapter(language=language)
    df = legacy.ensure_muavin_derived_columns(clean_df.copy())
    df = legacy.build_muavin_audit_columns(df)

    findings = pd.DataFrame({'Bulgu': legacy.make_muavin_findings(df)})

    user_headers, user_rows, _ = legacy.TbPlCcControlWindow.build_muavin_user_based_result(adapter, df)
    user_df = pd.DataFrame(user_rows, columns=user_headers)

    tax_result = legacy.TbPlCcControlWindow.build_muavin_tax_based_result(adapter, df)
    tax_dup = pd.DataFrame(tax_result['dup'][1], columns=tax_result['dup'][0])
    tax_projection = pd.DataFrame(tax_result['projection'][1], columns=tax_result['projection'][0])
    tax_indicator = pd.DataFrame(tax_result['indicator'][1], columns=tax_result['indicator'][0])

    account_content = legacy.TbPlCcControlWindow.build_muavin_account_content_result(adapter, df)
    text_df = pd.DataFrame(account_content['text'][1], columns=account_content['text'][0])
    relation_df = pd.DataFrame(account_content['relation'][1], columns=account_content['relation'][0])
    cost_df = pd.DataFrame(account_content['cost'][1], columns=account_content['cost'][0])

    period_df = (
        df.groupby('donem', dropna=False)
          .agg(
              Satir=('ana_hesap', 'size'),
              Belge_Adedi=('belge_anahtar', lambda s: s.replace('', pd.NA).dropna().nunique()),
              Tutar=('up_tutar', 'sum'),
              Riskli_Satir=('combined_risk_flag', 'sum') if 'combined_risk_flag' in df.columns else ('risk_flag', 'sum'),
          )
          .reset_index()
          .rename(columns={'donem': 'Dönem'})
    ) if not df.empty else pd.DataFrame()

    doctype_df = (
        df.groupby('belge_turu', dropna=False)
          .agg(Satir=('ana_hesap', 'size'), Belge_Adedi=('belge_anahtar', lambda s: s.replace('', pd.NA).dropna().nunique()), Tutar=('up_tutar', 'sum'))
          .reset_index()
          .rename(columns={'belge_turu': 'Belge Türü'})
          .sort_values(['Satir', 'Tutar'], ascending=[False, False])
    ) if not df.empty else pd.DataFrame()

    contra_df = (
        df.groupby(['karsi_hesap', 'karsi_hesap_tanimi'], dropna=False)
          .agg(Satir=('ana_hesap', 'size'), Belge_Adedi=('belge_anahtar', lambda s: s.replace('', pd.NA).dropna().nunique()), Tutar=('up_tutar', 'sum'), Ortalama_Risk=('audit_risk_score', 'mean') if 'audit_risk_score' in df.columns else ('risk_skoru', 'mean'))
          .reset_index()
          .sort_values(['Ortalama_Risk', 'Tutar'], ascending=[False, False])
    ) if not df.empty else pd.DataFrame()

    risk_user_df = (
        df.groupby('kullanici', dropna=False)
          .agg(Satir=('ana_hesap', 'size'), Riskli_Satir=('combined_risk_flag', 'sum') if 'combined_risk_flag' in df.columns else ('risk_flag', 'sum'), Ortalama_Risk=('audit_risk_score', 'mean') if 'audit_risk_score' in df.columns else ('risk_skoru', 'mean'))
          .reset_index()
          .sort_values(['Ortalama_Risk', 'Riskli_Satir'], ascending=[False, False])
          .rename(columns={'kullanici': 'Kullanıcı'})
    ) if not df.empty else pd.DataFrame()

    late7_df = df[df.get('late_7day_flag', False) == True].copy() if not df.empty and 'late_7day_flag' in df.columns else pd.DataFrame()
    doccheck_df = df[df.get('doc_relation_status', '').astype(str) != 'Normal'].copy() if not df.empty and 'doc_relation_status' in df.columns else pd.DataFrame()
    dupref_df = df[df.get('duplicate_open_reference', False) == True].copy() if not df.empty and 'duplicate_open_reference' in df.columns else pd.DataFrame()
    taxref_df = df[df.get('tax_risk_flag', False) == True].copy() if not df.empty and 'tax_risk_flag' in df.columns else pd.DataFrame()

    tables = {
        'findings': findings,
        'period': period_df,
        'doctype': doctype_df,
        'risk_user': risk_user_df,
        'user': user_df,
        'contra': contra_df,
        'text': text_df,
        'account_doc_relation': relation_df,
        'cost': cost_df,
        'dupref': tax_dup,
        'tax_projection': tax_projection,
        'tax_indicator': tax_indicator,
        'late7': late7_df,
        'doccheck': doccheck_df,
        'duprefdetail': dupref_df,
        'taxref': taxref_df,
        'drilldown': df,
        'document_lines': df,
    }
    return tables


def run_muavin_analysis(uploaded_files: List, mapping: Dict[str, str], language: str = 'tr') -> Dict:
    legacy = get_legacy()
    paths = temp_paths_from_uploaded_files(uploaded_files)
    payload = legacy.build_muavin_analysis_payload_from_files(file_paths=paths, mapping=mapping, language=language)
    clean_df = payload.get('clean_df', pd.DataFrame())
    payload['tables'] = _muavin_tables_from_clean_df(clean_df, language=language)
    return payload


def run_regular_ft_analysis(faggl_df: pd.DataFrame, eba_df: pd.DataFrame, zfi_df: pd.DataFrame, language: str = 'tr') -> Dict:
    legacy = get_legacy()
    adapter = StreamlitCoreAdapter(language=language)
    adapter.regular_ft_faggl_df = faggl_df.copy() if isinstance(faggl_df, pd.DataFrame) else pd.DataFrame()
    adapter.regular_ft_eba_df = eba_df.copy() if isinstance(eba_df, pd.DataFrame) else pd.DataFrame()
    adapter.regular_ft_zfi052_df = zfi_df.copy() if isinstance(zfi_df, pd.DataFrame) else pd.DataFrame()

    legacy.TbPlCcControlWindow.prepare_regular_ft_analysis(adapter)

    return {
        'adapter': adapter,
        'output_df': adapter.regular_ft_output_df.copy() if isinstance(adapter.regular_ft_output_df, pd.DataFrame) else pd.DataFrame(),
        'periods': list(adapter.regular_ft_periods),
        'current_period': adapter.regular_ft_current_period,
        'previous_period': adapter.regular_ft_previous_period,
    }
