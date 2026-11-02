"""
SUITE DE PRUEBAS AUTOMATIZADAS: test_sistema_hidrocarburos.py
=============================================================================
Verifica:
  1. Flujo integral de 3 capas, base de datos externa en BD/, 6 reportes y dictamen.
  2. Exportación del Modelo Estrella para Power BI (.dax/.json) y Tableau (.twb).
  3. Rechazo automático de contratos fuera de catálogo o saldos negativos.
  4. Inmutabilidad de Templates/Generico.xlsx y fórmulas vivas SUMIFS/SUM.
  5. Cero menciones a instituciones específicas (proyecto 100% genérico de hidrocarburos).
"""
import os
import re
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import unittest
import openpyxl

from src.constantes import (
    PROJECT_ROOT, DB_PATH, TEMPLATE_PATH, INSUMOS_DIR, REPORTES_DIR
)
from src.bo.hidrocarburos_bo import HidrocarburosBO
from src.ui.consola_ejecutiva import ejecutar_flujo_completo_hidrocarburos


class TestSistemaHidrocarburosGenerico(unittest.TestCase):
    def test_01_flujo_integral_y_exportacion_bi(self):
        res = ejecutar_flujo_completo_hidrocarburos()
        self.assertTrue(os.path.exists(DB_PATH))
        self.assertEqual(len(res["insumos_cargados"]), 6)
        self.assertEqual(len(res["reportes_generados"]), 6)
        self.assertEqual(len(res["dictamenes"]), 1)
        self.assertIn("Fact_Contraprestaciones_Hidrocarburos", res["bi_datasets"])
        self.assertTrue(os.path.exists(res["bi_datasets"]["Fact_Contraprestaciones_Hidrocarburos"]))

    def test_02_bloqueo_contrato_no_registrado_y_saldo_negativo(self):
        bo = HidrocarburosBO()
        ruta_err = os.path.join(
            INSUMOS_DIR, "INSUMO_ASIGNACIONES_INCORRECTO_CONTRATO_Y_SALDO_30_09_2026.xlsx"
        )
        ok, filas, n_err, detalle, nuevos = bo.validar_y_procesar_insumo_bo(ruta_err, "30/09/2026")
        self.assertFalse(ok)
        self.assertGreaterEqual(n_err, 2)
        self.assertIn("BLK-INEXISTENTE-999", nuevos)
        self.assertIn("Contraprestación negativa", detalle)

    def test_03_template_original_intacto_y_formulas_vivas(self):
        wb_tpl = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
        self.assertEqual(wb_tpl["DATOS"].max_row, 1)
        wb_tpl.close()

        ruta_cons = os.path.join(
            REPORTES_DIR, "Reporte_Hidrocarburos_Consolidado_Ingresos_Hidrocarburos_30092026.xlsx"
        )
        wb_rep = openpyxl.load_workbook(ruta_cons, data_only=False)
        self.assertGreater(wb_rep["DATOS"].max_row, 10)
        ws_r = wb_rep["REPORTE"]
        formulas = [
            str(ws_r.cell(r, c).value or "")
            for r in range(1, ws_r.max_row + 1)
            for c in range(1, 9)
            if str(ws_r.cell(r, c).value or "").startswith("=")
        ]
        wb_rep.close()
        self.assertTrue(any("SUMIFS(" in f for f in formulas))
        self.assertTrue(any("SUM(" in f for f in formulas))

    def test_04_proyecto_100_por_ciento_generico_sin_siglas_institucionales(self):
        patron_prohibido = re.compile(r"\b(shcp|cnh|bancomext|bmxt|upintsh|fmped|cnbv)\b", re.I)
        for root, dirs, files in os.walk(PROJECT_ROOT):
            if ".git" in dirs:
                dirs.remove(".git")
            for fn in files:
                if fn == "test_sistema_hidrocarburos.py":
                    continue
                m_fn = patron_prohibido.search(fn)
                self.assertIsNone(m_fn, f"Nombre de archivo prohibido: {fn}")
                p = os.path.join(root, fn)
                if fn.endswith((".py", ".md", ".dax", ".json", ".twb", ".bas", ".cls", ".vbs", ".bat", ".ps1", ".txt")):
                    with open(p, "rb") as _f:
                        raw = _f.read().decode("utf-8", errors="ignore")
                    m_txt = patron_prohibido.search(raw)
                    self.assertIsNone(m_txt, f"Término institucional encontrado en {p}: {m_txt.group(0) if m_txt else ''}")


if __name__ == "__main__":
    unittest.main()
