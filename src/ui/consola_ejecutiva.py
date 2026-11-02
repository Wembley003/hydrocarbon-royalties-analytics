"""
CAPA 1 (PRESENTACIÓN / EJECUCIÓN AUTOMATIZADA): consola_ejecutiva.py
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from src.constantes import INSUMOS_DIR, REPORTES_DIR, DICTAMENES_DIR, BI_DIR
from src.bo.hidrocarburos_bo import HidrocarburosBO


def ejecutar_flujo_completo_hidrocarburos() -> dict:
    bo = HidrocarburosBO()
    resultados = {"insumos_cargados": [], "reportes_generados": [], "dictamenes": [], "bi_datasets": {}}

    for fec_str, fec_tag in [("30/09/2026", "30_09_2026"), ("28/10/2026", "28_10_2026")]:
        for pref in ["INSUMO_ASIGNACIONES", "INSUMO_LICENCIAS", "INSUMO_PRODUCCION_COMPARTIDA"]:
            ruta_in = os.path.join(INSUMOS_DIR, f"{pref}_{fec_tag}.xlsx")
            ok, filas, n_err, detalle, _ = bo.validar_y_procesar_insumo_bo(ruta_in, fec_str)
            if not ok:
                raise RuntimeError(f"Error inesperado en insumo válido {ruta_in}: {detalle}")
            bo.guardar_insumo_validado_bo(ruta_in, fec_str)
            resultados["insumos_cargados"].append((os.path.basename(ruta_in), filas))

    fec_rep = "30/09/2026"
    for inc_a, inc_l, inc_p, inc_c in [
        (True, False, False, False),
        (False, True, False, False),
        (False, False, True, False),
        (True, True, True, True),
    ]:
        ok, r_out, msg = bo.generar_reporte_ejecutivo_bo(
            fec_rep, inc_a, inc_l, inc_p, inc_c, carpeta_destino=REPORTES_DIR
        )
        if not ok:
            raise RuntimeError(f"Error al generar reporte ejecutivo: {msg}")
        resultados["reportes_generados"].append(r_out)

    for n_ins in (2, 3):
        ok, r_cmp, msg = bo.generar_reporte_comparativo_bo(
            fec_rep, n_insumos=n_ins, carpeta_destino=REPORTES_DIR
        )
        if not ok:
            raise RuntimeError(f"Error al generar comparativo de {n_ins} insumos: {msg}")
        resultados["reportes_generados"].append(r_cmp)

    ok_d, r_d = bo.generar_dictamen_verificacion_bo(fec_rep, carpeta_destino=DICTAMENES_DIR)
    if ok_d:
        resultados["dictamenes"].append(r_d)

    resultados["bi_datasets"] = bo.exportar_modelo_estrella_bi_bo(BI_DIR)
    return resultados


if __name__ == "__main__":
    res = ejecutar_flujo_completo_hidrocarburos()
    print("=== FLUJO DE FISCALIZACIÓN DE HIDROCARBUROS COMPLETADO AL 100% ===")
    print(f"Insumos validados y cargados en BD/: {len(res['insumos_cargados'])}")
    print(f"Reportes generados desde Templates/Generico.xlsx: {len(res['reportes_generados'])}")
    print(f"Datasets Estrella exportados para Power BI / Tableau: {len(res['bi_datasets'])}")
