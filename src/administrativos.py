"""Clasificacion de tareas administrativas por Title y referencias configurables."""
import json
import math
import re
import unicodedata


def normalized_title(value):
    plain = ''.join(c for c in unicodedata.normalize('NFKD', str(value or '').lower()) if not unicodedata.combining(c))
    return ' '.join(re.sub(r'[^a-z0-9]+', ' ', plain).split())


def load_rules(path):
    if not path.is_file():
        return {'activo': False, 'tipos_elemento': [], 'solo_pendientes': False, 'reglas': [], 'excluir_patrones': []}
    try:
        settings = json.loads(path.read_text(encoding='utf-8-sig'))
        if not isinstance(settings, dict):
            raise ValueError('la configuracion debe ser un objeto')
        for key in ('activo', 'solo_pendientes'):
            if not isinstance(settings.get(key), bool):
                raise ValueError(key + ' debe ser booleano')
        if not isinstance(settings.get('tipos_elemento'), list) or not settings['tipos_elemento'] or not all(isinstance(t, str) and t for t in settings['tipos_elemento']):
            raise ValueError('tipos_elemento debe ser una lista de tipos')
        if not isinstance(settings.get('reglas'), list):
            raise ValueError('reglas debe ser una lista')
        def patterns(values, name, required=False):
            if not isinstance(values, list) or required and not values or not all(isinstance(v, str) and v.strip() for v in values):
                raise ValueError(name + ' debe ser una lista de patrones no vacios')
            return [re.compile(v, re.IGNORECASE) for v in values]
        settings['_excluir'] = patterns(settings.get('excluir_patrones', []), 'excluir_patrones')
        identifiers = set()
        for rule in settings['reglas']:
            if not isinstance(rule, dict) or not isinstance(rule.get('id'), str) or not rule['id'] or rule['id'] in identifiers:
                raise ValueError('cada regla debe tener id unico')
            identifiers.add(rule['id'])
            if not isinstance(rule.get('nombre'), str) or not rule['nombre'] or not isinstance(rule.get('activo'), bool):
                raise ValueError('nombre y activo invalidos en ' + rule['id'])
            for key in ('horas_referencia', 'horas_maximas_sin_warning'):
                value = rule.get(key)
                if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
                    raise ValueError(key + ' debe ser un numero positivo finito en ' + rule['id'])
            if rule['horas_maximas_sin_warning'] < rule['horas_referencia']:
                raise ValueError('el umbral no puede ser menor que la referencia en ' + rule['id'])
            rule['_patrones'] = patterns(rule.get('patrones'), rule['id'] + '.patrones', True)
            rule['_excluir'] = patterns(rule.get('excluir_patrones', []), rule['id'] + '.excluir_patrones')
        return settings
    except (ValueError, TypeError, re.error) as error:
        raise ValueError(f'Configuracion administrativa invalida ({path}): {error}') from error


def classify(title, kind, state, hours, settings, final_states):
    if not settings['activo'] or kind not in settings['tipos_elemento'] or settings['solo_pendientes'] and state in final_states:
        return None
    text = normalized_title(title)
    matches = [r for r in settings['reglas'] if r['activo'] and any(p.search(text) for p in r['_patrones'])]
    if not matches:
        return None
    result = {'reglas': '; '.join(r['id'] for r in matches), 'categoria': '; '.join(r['nombre'] for r in matches),
              'referencia': None, 'umbral': None, 'diferencia': None, 'multiplicador': None}
    excluded = any(p.search(text) for p in settings['_excluir']) or any(p.search(text) for r in matches for p in r['_excluir'])
    if len(matches) > 1 or excluded:
        result.update(resultado='REVISION_MANUAL', detalle='Varias categorias en el titulo.' if len(matches) > 1 else 'El titulo coincide con una exclusion de complejidad o alcance. Revisar manualmente.')
        return result
    rule = matches[0]
    reference, limit = rule['horas_referencia'], rule['horas_maximas_sin_warning']
    result.update(referencia=reference, umbral=limit)
    if hours is None:
        result.update(resultado='SIN_ESTIMACION', detalle='No hay Story Points validos para comparar con la referencia.')
        return result
    result.update(diferencia=hours-reference, multiplicador=hours/reference)
    if hours > limit:
        result.update(resultado='WARNING_ESTIMACION_ALTA', detalle=f'{hours:g} h registradas frente a {reference:g} h de referencia; supera el umbral de {limit:g} h. Revisar alcance y estimacion; no demuestra sobreestimacion ni tiempo ejecutado.')
    else:
        result.update(resultado='DENTRO_UMBRAL', detalle=f'{hours:g} h registradas; no supera el umbral de {limit:g} h. Referencia orientativa: {reference:g} h.')
    return result
