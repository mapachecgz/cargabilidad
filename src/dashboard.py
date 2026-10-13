"""Dashboard HTML autocontenido: funciona localmente sin servicios externos."""
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path


def write_dashboard(path, payload):
    template = (Path(__file__).parent / 'dashboard.html').read_text(encoding='utf-8-sig')
    serialized = json.dumps(payload, ensure_ascii=True).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')
    path.write_text(template.replace('__DATA__', serialized), encoding='utf-8')


def generate_dashboard(directory, payload, timestamp=None):
    """Conserva un solo reporte HTML generado, fechado en hora de Bogota."""
    directory = Path(directory).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    timestamp = timestamp or datetime.now(timezone(timedelta(hours=-5))).strftime('%Y-%m-%d_%H-%M-%S')
    path = directory / f'cargabilidad_{timestamp}.html'
    write_dashboard(path, payload)
    for previous in directory.glob('cargabilidad*.html'):
        resolved = previous.resolve()
        if resolved == path or resolved.parent != directory or previous.is_symlink():
            continue
        content = previous.read_text(encoding='utf-8-sig')
        if '<script id="report-data" type="application/json">' in content and '<h1>Cargabilidad del equipo</h1>' in content:
            previous.unlink()
    return path
