import hashlib
import importlib.util
import json
import os
import tempfile
from pathlib import Path
from types import SimpleNamespace
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
LEGACY_FILE = BASE_DIR / 'tb_pl_cc_control.py'
NOTES_JSON = BASE_DIR / 'notes.json'
RESPONSIBLES_JSON = BASE_DIR / 'responsibles.json'
USERS_JSON = BASE_DIR / 'users.json'


@st.cache_resource
def load_legacy_module():
    spec = importlib.util.spec_from_file_location('tbplcc_legacy', str(LEGACY_FILE))
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


mod = load_legacy_module()


def ensure_state():
    defaults = {
        'logged_in': False,
        'username': '',
        'permissions': {},
        'language': 'tr',
        'tb_df': None,
        'plcc_df': None,
        'tb_name': '',
        'plcc_name': '',
        'available_periods': [],
        'current_period': '',
        'previous_period': '',
        'tb_rows': None,
        'plcc_detail': None,
        'plcc_subtotal': None,
        'balance_df': None,
        'income_df': None,
        'muavin_payload': None,
        'regular_ft_df': None,
        'regular_ft_periods': [],
        'regular_ft_current': '',
        'regular_ft_previous': '',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def read_json(path: Path, default):
    if not path.exists():
        return default
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def write_json(path: Path, payload):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)



def hash_password(value: str) -> str:
    return hashlib.sha256(str(value).encode('utf-8')).hexdigest()


def verify_login(username: str, password: str) -> Optional[Dict]:
    users = read_json(USERS_JSON, [])
    uname = str(username).strip().lower()
    pwd_hash = hash_password(password)
    for item in users:
        if str(item.get('username', '')).strip().lower() == uname and str(item.get('password_hash', '')).strip() == pwd_hash:
            return item
    return None


def get_notes() -> List[Dict]:
    return read_json(NOTES_JSON, [])


def get_responsibles() -> List[Dict]:
    return read_json(RESPONSIBLES_JSON, [])


def get_users() -> List[Dict]:
    return read_json(USERS_JSON, [])


class StreamlitAdapter:
    def __init__(self, language='tr'):
        self.language = language
        self.notes = get_notes()
        self.responsibles = get_responsibles()
        self.tb_raw_df = None
        self.plcc_raw_df = None
        self.available_periods = []
        self.current_period = ''
        self.previous_period = ''
        self.regular_ft_faggl_df = None
        self.regular_ft_eba_df = None
        self.regular_ft_zfi052_df = None
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


def build_main_analysis(tb_df: Optional[pd.DataFrame], plcc_df: Optional[pd.DataFrame], language='tr'):
    adapter = StreamlitAdapter(language=language)
    adapter.tb_raw_df = tb_df
    adapter.plcc_raw_df = plcc_df
    adapter.available_periods = mod.TbPlCcControlWindow.get_available_periods_from_loaded_files(adapter)
    selectable = [p for p in adapter.available_periods if p != mod.OPENING_PERIOD_TR]
    if not selectable:
        selectable = ['02-Mart', '03-Nisan']
    adapter.current_period = selectable[-1]
    adapter.previous_period = selectable[-2] if len(selectable) >= 2 else selectable[-1]
    tb_rows = mod.TbPlCcControlWindow.build_tb_rows(adapter) if tb_df is not None and not tb_df.empty else []
    plcc_detail, plcc_subtotal = mod.TbPlCcControlWindow.build_plcc_rows(adapter) if plcc_df is not None and not plcc_df.empty else ([], [])
    balance_rows = mod.build_balance_sheet_summary(tb_rows, adapter.available_periods, current_period=adapter.current_period, previous_period=adapter.previous_period) if tb_rows else []
    income_rows = mod.build_income_statement_summary(tb_rows, adapter.available_periods, current_period=adapter.current_period, previous_period=adapter.previous_period) if tb_rows else []
    return adapter, tb_rows, plcc_detail, plcc_subtotal, balance_rows, income_rows


def df_from_rows(rows: List[Dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


def save_uploaded_to_temp(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix or '.xlsx'
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.flush()
    tmp.close()
    return tmp.name


def login_screen():
    st.title('TB & PL-CC Control - Streamlit')
    st.caption('Desktop PySide6 uygulamasının Streamlit uyarlaması. Mevcut kural seti, kullanıcı/izin mantığı, not-sorumlu kayıtları ve ana analiz akışları korunacak şekilde kurgulandı.')
    with st.form('login_form'):
        username = st.text_input('Kullanıcı adı')
        password = st.text_input('Şifre', type='password')
        submitted = st.form_submit_button('Giriş Yap', use_container_width=True)
    if submitted:
        user = verify_login(username, password)
        if user:
            st.session_state.logged_in = True
            st.session_state.username = user.get('username', '')
            st.session_state.permissions = user.get('permissions', {})
            st.rerun()
        st.error('Kullanıcı adı veya şifre hatalı.')


def sidebar_menu():
    st.sidebar.success(f"Kullanıcı: {st.session_state.username}")
    if st.sidebar.button('Çıkış Yap', use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()


def permission(name: str) -> bool:
    return bool(st.session_state.permissions.get(name, False))


def page_dashboard():
    st.header('1) Dashboard / Dosya Yükleme')
    c1, c2 = st.columns(2)
    tb_up = c1.file_uploader('TB dosyası', type=['xlsx', 'xls'], key='tb_uploader')
    plcc_up = c2.file_uploader('PL-CC dosyası', type=['xlsx', 'xls'], key='plcc_uploader')

    if tb_up is not None:
        st.session_state.tb_df = pd.read_excel(tb_up)
        st.session_state.tb_name = tb_up.name
    if plcc_up is not None:
        st.session_state.plcc_df = pd.read_excel(plcc_up)
        st.session_state.plcc_name = plcc_up.name

    st.write({
        'TB': st.session_state.tb_name or '-',
        'PL-CC': st.session_state.plcc_name or '-',
    })

    if st.button('Ana Analizi Hazırla', use_container_width=True):
        with st.spinner('Ana analiz hazırlanıyor...'):
            adapter, tb_rows, plcc_detail, plcc_subtotal, balance_rows, income_rows = build_main_analysis(
                st.session_state.tb_df, st.session_state.plcc_df, language='tr'
            )
            st.session_state.available_periods = adapter.available_periods
            st.session_state.current_period = adapter.current_period
            st.session_state.previous_period = adapter.previous_period
            st.session_state.tb_rows = df_from_rows(tb_rows)
            st.session_state.plcc_detail = df_from_rows(plcc_detail)
            st.session_state.plcc_subtotal = df_from_rows(plcc_subtotal)
            st.session_state.balance_df = df_from_rows(balance_rows)
            st.session_state.income_df = df_from_rows(income_rows)
        st.success('Ana analiz hazırlandı.')

    if st.session_state.available_periods:
        st.info(f"Kullanılabilir dönemler: {', '.join(st.session_state.available_periods)}")
        st.write({
            'Cari dönem': st.session_state.current_period,
            'Karşılaştırma dönemi': st.session_state.previous_period,
        })


def page_analysis():
    st.header('2) Analiz')
    if st.session_state.tb_rows is None and st.session_state.plcc_detail is None:
        st.warning('Önce Dashboard ekranında ana analizi hazırlayın.')
        return

    tabs = st.tabs(['TB', 'PL-CC Detay', 'PL-CC Alt Toplam', 'Bilanço', 'Gelir Tablosu'])
    with tabs[0]:
        df = st.session_state.tb_rows if isinstance(st.session_state.tb_rows, pd.DataFrame) else pd.DataFrame()
        q = st.text_input('TB içinde ara', key='tb_search')
        if q and not df.empty:
            mask = mod.df_text_search_mask(df, list(df.columns), q)
            df = df[mask]
        st.dataframe(df, use_container_width=True, height=600)
    with tabs[1]:
        df = st.session_state.plcc_detail if isinstance(st.session_state.plcc_detail, pd.DataFrame) else pd.DataFrame()
        q = st.text_input('PL-CC detay içinde ara', key='plcc_detail_search')
        if q and not df.empty:
            mask = mod.df_text_search_mask(df, list(df.columns), q)
            df = df[mask]
        st.dataframe(df, use_container_width=True, height=600)
    with tabs[2]:
        st.dataframe(st.session_state.plcc_subtotal if isinstance(st.session_state.plcc_subtotal, pd.DataFrame) else pd.DataFrame(), use_container_width=True, height=600)
    with tabs[3]:
        st.dataframe(st.session_state.balance_df if isinstance(st.session_state.balance_df, pd.DataFrame) else pd.DataFrame(), use_container_width=True, height=600)
    with tabs[4]:
        st.dataframe(st.session_state.income_df if isinstance(st.session_state.income_df, pd.DataFrame) else pd.DataFrame(), use_container_width=True, height=600)


def page_notes():
    st.header('3) Not Tanımları')
    notes = get_notes()
    with st.form('notes_form'):
        col1, col2, col3 = st.columns(3)
        hesap = col1.text_input('Hesap Kodu')
        ana_hesap = col2.text_input('Ana Hesap (opsiyonel)')
        masraf = col3.text_input('Masraf Yeri (opsiyonel)')
        not_tr = st.text_area('Not')
        note_en = st.text_area('Note EN')
        submitted = st.form_submit_button('Kaydet', use_container_width=True)
    if submitted:
        notes.append({'hesap': hesap, 'anaHesap': ana_hesap, 'masrafYeri': masraf, 'not': not_tr, 'noteEn': note_en or not_tr})
        write_json(NOTES_JSON, notes)
        st.success('Not kaydedildi.')
        st.rerun()
    st.dataframe(pd.DataFrame(notes), use_container_width=True)


def page_responsibles():
    st.header('4) Sorumlu Tayin')
    st.caption('Excel’den iki veya üç kolon yapıştırabilirsiniz: hesap, anaHesap, sorumlu')
    paste = st.text_area('Toplu yapıştırma alanı')
    if st.button('Toplu Kaydet', use_container_width=True):
        rows = []
        for line in paste.splitlines():
            parts = [x.strip() for x in line.split('\t')]
            if len(parts) >= 2:
                if len(parts) == 2:
                    rows.append({'hesap': parts[0], 'anaHesap': '', 'sorumlu': parts[1]})
                else:
                    rows.append({'hesap': parts[0], 'anaHesap': parts[1], 'sorumlu': parts[2]})
        write_json(RESPONSIBLES_JSON, rows)
        st.success(f'{len(rows)} satır kaydedildi.')
        st.rerun()
    st.dataframe(pd.DataFrame(get_responsibles()), use_container_width=True)


def page_muavin():
    st.header('5) Muavin Analiz')
    files = st.file_uploader('Muavin dosyaları', type=['xlsx', 'xls'], accept_multiple_files=True)
    if not files:
        st.info('Bir veya birden fazla muavin dosyası yükleyin.')
        return

    headers = []
    temp_paths = []
    for f in files:
        path = save_uploaded_to_temp(f)
        temp_paths.append(path)
        headers.extend(mod.read_excel_headers_only(path))
    unique_headers = list(dict.fromkeys([str(x).strip() for x in headers if str(x).strip()]))

    st.subheader('Kolon Eşleştirme')
    required = {
        'yilay': 'Yıl/ay',
        'ana_hesap': 'Ana hesap',
        'ana_hesap_adi': 'DK hesabı uzun metni',
        'up_tutar': 'Tutar',
        'referans': 'Referans',
        'belge_numarasi': 'Belge numarası',
        'belge_turu': 'Belge türü',
        'karsi_hesap': 'Karşıt hesap',
        'karsi_hesap_tanimi': 'Karşıt hesap tanımı',
        'denklestirme': 'Denkleştirme',
        'metin': 'Metin',
        'ters_kayit': 'Ters kayıt belge no',
        'kullanici': 'Kullanıcı adı',
        'vergi_gostergesi': 'Vergi göstergesi',
        'belge_tarihi': 'Belge tarihi',
        'kayit_tarihi': 'Kayıt tarihi',
        'giris_tarihi': 'Giriş tarihi',
        'masraf_yeri': 'Masraf yeri',
        'masraf_yeri_tanimi': 'Masraf yeri tanımı',
    }
    mapping = {}
    cols = st.columns(3)
    for i, (key, label) in enumerate(required.items()):
        with cols[i % 3]:
            mapping[key] = st.selectbox(label, options=[''] + unique_headers, key=f'muavin_map_{key}')

    if st.button('Muavin Analizini Çalıştır', use_container_width=True):
        with st.spinner('Muavin analizi hazırlanıyor...'):
            payload = mod.build_muavin_analysis_payload_from_files(temp_paths, mapping, language='tr')
            st.session_state.muavin_payload = payload
        st.success('Muavin analizi hazırlandı.')

    payload = st.session_state.muavin_payload
    if payload:
        clean_df = payload.get('clean_df', pd.DataFrame())
        st.metric('Satır', len(clean_df))
        tabs = st.tabs(['Temel Veri', 'Bulgular', 'Dönem Özeti', 'Belge Türü', 'Riskli Satırlar'])
        with tabs[0]:
            st.dataframe(clean_df, use_container_width=True, height=600)
        with tabs[1]:
            findings = mod.make_muavin_findings(clean_df)
            for f in findings:
                st.write('- ' + str(f))
        with tabs[2]:
            if not clean_df.empty:
                period_summary = clean_df.groupby('donem').agg(Satir=('ana_hesap', 'size'), Tutar=('up_tutar', 'sum')).reset_index()
                st.dataframe(period_summary, use_container_width=True)
        with tabs[3]:
            if not clean_df.empty:
                doc_summary = clean_df.groupby('belge_turu').agg(Satir=('ana_hesap', 'size'), Tutar=('up_tutar', 'sum')).reset_index()
                st.dataframe(doc_summary, use_container_width=True)
        with tabs[4]:
            if not clean_df.empty:
                risky = clean_df[clean_df['risk_flag'] == True]
                st.dataframe(risky, use_container_width=True, height=600)


def page_regular_ft():
    st.header('6) Düzenli Gelen Ft Analiz')
    c1, c2, c3 = st.columns(3)
    faggl = c1.file_uploader('FAGGL', type=['xlsx', 'xls'])
    eba = c2.file_uploader('EBA', type=['xlsx', 'xls'])
    zfi = c3.file_uploader('ZFI052', type=['xlsx', 'xls'])

    if st.button('Düzenli Gelen Ft Analizini Çalıştır', use_container_width=True):
        if not all([faggl, eba, zfi]):
            st.error('Üç dosya da gerekli: FAGGL, EBA, ZFI052')
        else:
            with st.spinner('Düzenli gelen fatura analizi hazırlanıyor...'):
                adapter = StreamlitAdapter(language='tr')
                adapter.regular_ft_faggl_df = pd.read_excel(faggl)
                adapter.regular_ft_eba_df = pd.read_excel(eba)
                adapter.regular_ft_zfi052_df = pd.read_excel(zfi)
                mod.TbPlCcControlWindow.prepare_regular_ft_analysis(adapter)
                st.session_state.regular_ft_df = adapter.regular_ft_output_df.copy()
                st.session_state.regular_ft_periods = list(adapter.regular_ft_periods)
                st.session_state.regular_ft_current = adapter.regular_ft_current_period
                st.session_state.regular_ft_previous = adapter.regular_ft_previous_period
            st.success('Düzenli gelen ft analizi hazırlandı.')

    df = st.session_state.regular_ft_df
    if isinstance(df, pd.DataFrame) and not df.empty:
        st.write({
            'Dönemler': ', '.join(st.session_state.regular_ft_periods),
            'Cari dönem': st.session_state.regular_ft_current,
            'Önceki dönem': st.session_state.regular_ft_previous,
        })
        st.dataframe(df, use_container_width=True, height=650)


def page_users():
    st.header('7) Kullanıcı Yönetimi')
    users = get_users()
    st.dataframe(pd.DataFrame(users), use_container_width=True)


def page_publish_guide():
    st.header('Yayımlama Rehberi')
    st.markdown('''
1. Bu dosyaları bir GitHub reposuna koy:
   - `app_streamlit_tb_pl_cc.py`
   - `tb_pl_cc_control.py`
   - `notes.json`
   - `responsibles.json`
   - `users.json`
   - `requirements.txt`

2. Lokal test:
   ```bash
   pip install -r requirements.txt
   streamlit run app_streamlit_tb_pl_cc.py
   ```

3. Streamlit Community Cloud:
   - GitHub repo bağla.
   - Main file path olarak `app_streamlit_tb_pl_cc.py` seç.
   - Deploy et.

4. Kurumsal kullanım için daha sağlam seçenek:
   - Docker + Render / Azure App Service / AWS ECS.
   - JSON dosyalarını yerel disk yerine veritabanına taşı.
   - Kullanıcı şifrelerini salt + hash ile yeniden düzenle.
   - Büyük Excel yüklemelerinde object storage ve job queue ekle.

5. Üretim önerileri:
   - JSON yerine PostgreSQL.
   - Audit log tablosu.
   - İndirilen dosya logu.
   - Role-based access control.
   - Background worker (Celery/RQ).
   - Büyük veri için Parquet cache.
''')


def main():
    st.set_page_config(page_title='TB & PL-CC Streamlit', layout='wide')
    ensure_state()
    if not st.session_state.logged_in:
        login_screen()
        return

    sidebar_menu()
    pages = ['Dashboard', 'Analiz', 'Yayımlama']
    if permission('notes'):
        pages.append('Not Tanımları')
    if permission('responsibles'):
        pages.append('Sorumlu Tayin')
    if permission('muavin'):
        pages.append('Muavin Analiz')
    if permission('regular_ft'):
        pages.append('Düzenli Gelen Ft')
    if permission('user_management'):
        pages.append('Kullanıcı Yönetimi')

    page = st.sidebar.radio('Menü', pages)
    if page == 'Dashboard':
        page_dashboard()
    elif page == 'Analiz':
        page_analysis()
    elif page == 'Not Tanımları':
        page_notes()
    elif page == 'Sorumlu Tayin':
        page_responsibles()
    elif page == 'Muavin Analiz':
        page_muavin()
    elif page == 'Düzenli Gelen Ft':
        page_regular_ft()
    elif page == 'Kullanıcı Yönetimi':
        page_users()
    else:
        page_publish_guide()


if __name__ == '__main__':
    main()
