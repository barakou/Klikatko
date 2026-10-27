import json
import os
import re
import tempfile
import unicodedata
from pathlib import Path
from uuid import uuid4

ROOT = Path(__file__).resolve().parent.parent


def catalog():
    return json.loads((ROOT / 'klikatko/catalog.json').read_text(encoding='utf-8'))


def flatten(nodes=None):
    for node in catalog()['nodes'] if nodes is None else nodes:
        yield node
        yield from flatten(node.get('children', []))


def selected_nodes(selected):
    """Keep the selected nodes and ancestor headings, in source order."""
    ids = set(selected)
    known = {n['id'] for n in flatten()}
    if ids - known:
        raise ValueError('Výběr obsahuje neznámé kapitoly. Načtěte aktuální osnovu.')

    def visit(nodes):
        result = []
        for node in nodes:
            children = visit(node.get('children', []))
            if node['id'] in ids or children:
                result.append({**node, 'children': children, 'explicit': node['id'] in ids})
        return result
    return visit(catalog()['nodes'])


def safe_name(value):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '-', str(value)).strip(' .')
    value = value[:110].strip(' .') or 'Technicka-zprava'
    if value.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1, 10)), *(f'LPT{i}' for i in range(1, 10))}:
        value = '_' + value
    return value


def validate_project(data, require_name=False):
    if not isinstance(data, dict):
        raise ValueError('Neplatná zakázka.')
    fields = ('name', 'number', 'location', 'investor', 'author', 'date', 'stage', 'output_dir', 'filename', 'cover_image')
    result = {key: str(data.get(key, '')).strip() for key in fields}
    for value in result.values():
        if len(value) > 4000:
            raise ValueError('Zadaný text je příliš dlouhý.')
    if require_name and not result['name']:
        raise ValueError('Vyplňte název stavby.')
    selected = data.get('selected', [])
    if not isinstance(selected, list) or not all(isinstance(x, str) for x in selected):
        raise ValueError('Neplatný výběr kapitol.')
    selected_nodes(selected)
    result['selected'] = list(dict.fromkeys(selected))
    result['id'] = str(data.get('id', ''))
    return result


class Store:
    def __init__(self, directory):
        self.root = Path(directory).expanduser().resolve()
        self.root.mkdir(parents=True, exist_ok=True)

    def _path(self, kind, ident):
        if kind not in ('projects', 'presets') or not re.fullmatch(r'[a-f0-9]{32}', ident):
            raise ValueError('Neplatný identifikátor.')
        directory = self.root / kind
        directory.mkdir(exist_ok=True)
        return directory / (ident + '.json')

    def list(self, kind):
        directory = self.root / kind
        if not directory.exists():
            return []
        return sorted([json.loads(p.read_text(encoding='utf-8')) for p in directory.glob('*.json')], key=lambda x: x.get('name', '').casefold())

    def save(self, kind, data):
        if kind == 'projects':
            data = validate_project(data, require_name=True)
        elif kind == 'presets':
            if not str(data.get('name', '')).strip():
                raise ValueError('Zadejte název předvolby.')
            selected_nodes(data.get('selected', []))
            data = {'name': str(data['name']).strip()[:160], 'selected': data['selected'], 'id': data.get('id', '')}
        ident = data.get('id') or uuid4().hex
        path = self._path(kind, ident)
        data = {**data, 'id': ident, 'schema_version': 1}
        fd, temp = tempfile.mkstemp(dir=path.parent, suffix='.tmp')
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(temp, path)
        finally:
            Path(temp).unlink(missing_ok=True)
        return data

    def delete(self, kind, ident):
        self._path(kind, ident).unlink()
