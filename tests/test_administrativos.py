import copy
import json
import tempfile
import unittest
from pathlib import Path
from test_reglas import report
from administrativos import load_rules, classify


class AdministrativeTest(unittest.TestCase):
    def setUp(self):
        self.path = Path(__file__).resolve().parents[1]/'config'/'tareas_administrativas.json'
        self.settings = load_rules(self.path)
        self.finals = ['Closed', 'Removed', 'Resolved']

    def evaluate(self, title, hours, kind='Task', state='Active'):
        return classify(title, kind, state, hours, self.settings, self.finals)

    def test_accents_uppercase_and_warning_boundary(self):
        title = 'SP255 - CREACIÓN DE USUARIOS QA'
        self.assertEqual(self.evaluate(title, 2)['resultado'], 'DENTRO_UMBRAL')
        result = self.evaluate(title, 2.5)
        self.assertEqual(result['resultado'], 'WARNING_ESTIMACION_ALTA')
        self.assertEqual(result['referencia'], 1)
        self.assertEqual(result['multiplicador'], 2.5)

    def test_all_six_categories(self):
        examples = {'Crear usuario':'creacion_usuarios', 'Borrado de usuarios':'borrado_usuarios',
                    'Ejecución de script':'ejecucion_script', 'Otorgar permisos':'permisos',
                    'Cambiar contraseña':'contrasena', 'Desbloqueo del usuario':'desbloqueo'}
        for title, category in examples.items():
            self.assertEqual(self.evaluate(title, 3)['reglas'], category)

    def test_multiple_and_complex_titles_require_manual_review(self):
        for title in ('Creacion de usuario y otorgar permisos', 'Ejecucion masiva de script',
                      'Revision de permisos de auditoria', 'Creacion esquema y asignacion permisos',
                      'Crear usuarios DEV y QA'):
            result = self.evaluate(title, 8)
            self.assertEqual(result['resultado'], 'REVISION_MANUAL', title)
            self.assertIsNone(result['referencia'])

    def test_wrong_operation_and_unsupported_types_do_not_match(self):
        for title in ('Eliminar sesiones activas usuario', 'Generar script', 'Crear tabla', 'Gestion de usuarios'):
            self.assertIsNone(self.evaluate(title, 8), title)
        self.assertIsNone(self.evaluate('Crear usuario', 8, kind='Epic'))

    def test_missing_hours_and_closed_scope(self):
        self.assertEqual(self.evaluate('Ejecutar script', None)['resultado'], 'SIN_ESTIMACION')
        self.assertEqual(self.evaluate('Ejecutar script', 0)['resultado'], 'DENTRO_UMBRAL')
        self.assertEqual(self.evaluate('Crear usuario', 8, state='Closed')['resultado'], 'WARNING_ESTIMACION_ALTA')
        self.settings['solo_pendientes'] = True
        self.assertIsNone(self.evaluate('Crear usuario', 8, state='Closed'))

    def test_custom_threshold_and_disabled_rule(self):
        rule = self.settings['reglas'][0]
        rule['horas_referencia'] = 2
        rule['horas_maximas_sin_warning'] = 5
        result = self.evaluate('Crear usuario', 4)
        self.assertEqual(result['resultado'], 'DENTRO_UMBRAL')
        self.assertEqual(result['referencia'], 2)
        rule['activo'] = False
        self.assertIsNone(self.evaluate('Crear usuario', 8))

    def test_optional_file_and_invalid_configuration(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'tareas_administrativas.json'
            self.assertFalse(load_rules(path)['activo'])
            original = json.loads(self.path.read_text(encoding='utf-8'))
            for change in ('regex', 'duplicate', 'hours', 'threshold'):
                data = copy.deepcopy(original)
                if change == 'regex': data['reglas'][0]['patrones'] = ['[']
                elif change == 'duplicate': data['reglas'].append(copy.deepcopy(data['reglas'][0]))
                elif change == 'hours': data['reglas'][0]['horas_referencia'] = True
                else: data['reglas'][0]['horas_maximas_sin_warning'] = 0.5
                path.write_text(json.dumps(data),encoding='utf-8')
                with self.assertRaisesRegex(ValueError, 'Configuracion administrativa invalida'):
                    load_rules(path)
