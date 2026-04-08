
import hashlib
import json
from pathlib import Path
from typing import Any, Dict, List


def data_dir(project_root: Path) -> Path:
    target = Path(project_root).resolve() / 'data'
    target.mkdir(parents=True, exist_ok=True)
    return target


def users_path(project_root: Path) -> Path:
    return data_dir(project_root) / 'users.json'


def notes_path(project_root: Path) -> Path:
    return data_dir(project_root) / 'notes.json'


def responsibles_path(project_root: Path) -> Path:
    return data_dir(project_root) / 'responsibles.json'


def read_json(path: Path, default: Any):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except Exception:
        return default


def write_json(path: Path, payload: Any):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')


def sha256_text(value: str) -> str:
    return hashlib.sha256(str(value).encode('utf-8')).hexdigest()


def ensure_seed_files(project_root: Path, default_permissions: Dict[str, bool]):
    u_path = users_path(project_root)
    n_path = notes_path(project_root)
    r_path = responsibles_path(project_root)

    if not n_path.exists():
        write_json(n_path, [])
    if not r_path.exists():
        write_json(r_path, [])
    if not u_path.exists():
        write_json(u_path, [{
            'username': 'admin',
            'password_hash': sha256_text('admin'),
            'is_admin': True,
            'permissions': default_permissions,
        }])


def load_users(project_root: Path) -> List[Dict]:
    return read_json(users_path(project_root), [])


def save_users(project_root: Path, users: List[Dict]):
    write_json(users_path(project_root), users)


def load_notes(project_root: Path) -> List[Dict]:
    return read_json(notes_path(project_root), [])


def save_notes(project_root: Path, notes: List[Dict]):
    write_json(notes_path(project_root), notes)


def load_responsibles(project_root: Path) -> List[Dict]:
    return read_json(responsibles_path(project_root), [])


def save_responsibles(project_root: Path, items: List[Dict]):
    write_json(responsibles_path(project_root), items)
