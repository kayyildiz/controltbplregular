
import importlib.util
from functools import lru_cache
from pathlib import Path

from .pyside6_stub import install as install_pyside6_stub


@lru_cache(maxsize=4)
def load_legacy_module(project_root: str):
    install_pyside6_stub()
    root = Path(project_root).resolve()
    legacy_path = root / 'legacy' / 'tb_pl_cc_control.py'
    spec = importlib.util.spec_from_file_location('tbplcc_legacy_streamlit', str(legacy_path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f'Legacy module could not be loaded: {legacy_path}')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    data_dir = root / 'data'
    data_dir.mkdir(parents=True, exist_ok=True)

    module.BASE_DIR = str(data_dir)
    module.NOTES_JSON = str(data_dir / 'notes.json')
    module.RESPONSIBLES_JSON = str(data_dir / 'responsibles.json')
    module.USERS_JSON = str(data_dir / 'users.json')
    module.MUAVIN_EXPORT_DEFAULT = str(data_dir / 'Muavin_Analiz_Raporu.xlsx')
    return module
