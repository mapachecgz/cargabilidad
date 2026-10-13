import importlib.util
import unittest
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))

spec = importlib.util.spec_from_file_location('report', Path(__file__).resolve().parents[1] / 'src' / 'generar_reporte.py')
report = importlib.util.module_from_spec(spec)
spec.loader.exec_module(report)

class RulesTest(unittest.TestCase):
    def setUp(self):
        self.person = {'activo': True, 'horas_dia': 9, 'dedicacion': 1, 'ausencias': []}
        self.sprint = {'sprint': '255', 'inicio': '2026-10-07', 'fin': '2026-10-20', 'activo': True}

    def test_capacity_holidays_and_absence(self):
        self.assertEqual(report.capacity(self.person, self.sprint, []), 90)
        self.person['ausencias'] = [{'inicio': '2026-10-07', 'fin': '2026-10-09'}]
        self.assertEqual(report.capacity(self.person, self.sprint, ['2026-10-12']), 54)
        self.assertTrue(report.unavailable(self.person, date(2026, 10, 8)))
        self.assertFalse(report.unavailable(self.person, date(2026, 10, 10)))

    def test_inactive(self):
        self.person['activo'] = False
        self.assertEqual(report.capacity(self.person, self.sprint, []), 0)
        self.assertTrue(report.unavailable(self.person, date(2026, 10, 8)))

    def test_sprint_activities_and_example(self):
        times = report.DEFAULT_TIMES
        self.assertEqual(report.net_hours(42, 9, 0, sum(times['actividades_por_sprint'].values())), 20)
        breakdown = report.capacity_breakdown(self.person, self.sprint, ['2026-10-12'], times)
        self.assertEqual(breakdown, {'brutas': 90, 'festivos': 9, 'ausencias': 0, 'actividades': 13, 'disponibles': 68})
        self.assertEqual(report.capacity(self.person, self.sprint, ['2026-10-12','2026-10-12','2026-10-11'], times), 68)
        self.person['ausencias'] = [{'inicio':'2026-10-12','fin':'2026-10-12'}]
        self.assertEqual(report.capacity(self.person, self.sprint, ['2026-10-12'], times), 68)
        self.assertEqual(report.net_hours(9, 9, 0, 13), 0)

    def test_proportional_limits(self):
        limits = report.thresholds(64, report.DEFAULT_TIMES)
        self.assertEqual(limits, {'exceso_desde':64, 'punto_medio':55, 'capacidad':46})
        self.assertEqual(report.load_status(64,64,0,limits), 'EXCESO')
        self.assertEqual(report.load_status(55,64,0,limits), 'SOBRE PUNTO MEDIO')
        self.assertEqual(report.load_status(46,64,0,limits), 'SOBRE CAPACIDAD OBJETIVO')
        limits = report.thresholds(20, report.DEFAULT_TIMES)
        self.assertAlmostEqual(limits['punto_medio'],17.1875)
        self.assertAlmostEqual(limits['capacidad'],14.375)

    def test_infer_sprints_with_consistent_pattern(self):
        known=[{'sprint':'240','inicio':'2026-03-11','fin':'2026-03-24','activo':True},{'sprint':'241','inicio':'2026-03-25','fin':'2026-04-07','activo':True}]
        added=report.complete_sprints(known,[{'Iteration Level 3':'Sprint 239'},{'Iteration Level 3':'Sprint 242'}])
        self.assertEqual(added[0]['inicio'],'2026-02-25')
        self.assertEqual(added[1]['inicio'],'2026-04-08')
        self.assertEqual(added[1]['origen'],'inferido')
        self.assertEqual(report.complete_sprints(known,[{'Iteration Level 3':'Sprint 242'}]),[])
        known[1]['inicio']='2026-02-26'
        with self.assertRaises(ValueError): report.complete_sprints(known,[{'Iteration Level 3':'Sprint 243'}])

    def test_configured_vacation_and_sickness(self):
        self.person['nombre']='Ana'
        self.person['vacaciones']={'inicio':'2026-10-07','fin':'2026-10-09'}
        self.person['incapacidades']={'inicio':'2026-10-09','fin':'2026-10-12'}
        self.assertTrue(report.unavailable(self.person,date(2026,10,12)))
        self.assertEqual(report.capacity(self.person,self.sprint,[],report.DEFAULT_TIMES),41)
        self.person['incapacidades']['fin']=None
        with self.assertRaises(ValueError): report.resource_absences(self.person)

    def test_prior_sprint_pending(self):
        people=[{'nombre':'Ana','activo':True,'datos_base':{}}]
        sprints={'1':{'fin':'2026-10-06'},'2':{'fin':'2026-10-08'}}
        row={'persona':0,'pendiente':True,'sprint':'1','id':'1','tipo':'Task','estado':'Active','horas':3,'evc':'EVC','celula':'Celula'}
        records=[row,{**row,'id':'2','pendiente':False},{**row,'id':'3','sprint':'2'},{**row,'id':'4','persona':None},{**row,'id':'5','sprint':''}]
        result=report.previous_pending(records,people,sprints,date(2026,10,8))
        self.assertEqual(len(result),1)
        self.assertEqual(result[0]['dias'],2)
        self.assertEqual(result[0]['horas'],3)

    def test_identity(self):
        self.assertEqual(report.identity(' Alvaro Perez <TEST@example.com>'), ('Alvaro Perez', 'test@example.com'))
        self.assertEqual(report.normalize('  Alvaro  PEREZ '), report.normalize('Alvaro Perez'))

    def test_invalid_sprint(self):
        self.sprint['fin'] = '2026-10-01'
        with self.assertRaises(ValueError): report.capacity(self.person, self.sprint, [])

if __name__ == '__main__':
    unittest.main()
