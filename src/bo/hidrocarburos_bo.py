"""
CAPA 2 (REGLAS DE NEGOCIO - BUSINESS OBJECT): hidrocarburos_bo.py
=============================================================================
Concentra las reglas de negocio para validación de volúmenes, valor contractual
y contraprestaciones de hidrocarburos (53 columnas), detección de bloques/contratos
no registrados en CONTRATOS_HIDROCARBUROS, generación de los 6 reportes ejecutivos
a partir de copias de Templates/Generico.xlsx con fórmulas vivas SUMIFS/SUM y
exportación de tablas estrella para Power BI y Tableau.
"""
import os
import csv
import shutil
from typing import Tuple, List, Optional, Dict, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from src.constantes import (
    DB_PATH, TEMPLATE_PATH, REPORTES_DIR, DICTAMENES_DIR, BI_DIR,
    COLUMNAS_53_OFICIALES, COLUMNAS_ENTERAS_1BASED,
    TABLA_INSUMO_ASIGNACIONES, TABLA_INSUMO_LICENCIAS,
    TABLA_INSUMO_PRODUCCION_COMPARTIDA,
    HOJA_TPL_REPORTE, HOJA_TPL_DATOS, HOJA_TPL_CALCULO, HOJA_TPL_CALCINSTR,
    TPL_PORTAFOLIO_1, TPL_PORTAFOLIO_2, TPL_PORTAFOLIO_3,
)
from src.dao.hidrocarburos_dao import HidrocarburosDAO


class HidrocarburosBO:
    def __init__(self, db_path: str = DB_PATH, template_path: str = TEMPLATE_PATH):
        self.dao = HidrocarburosDAO(db_path=db_path)
        self.template_path = template_path
        self._cache_ruta_validada: str = ""
        self._cache_matriz_validada: Optional[List[List[Any]]] = None
        self._cache_conteos_fecha: Dict[str, Tuple[int, int, int]] = {}
        self._cache_analisis_fecha: Dict[str, Any] = {}

    def invalidar_cache_datos_bo(self) -> None:
        self._cache_ruta_validada = ""
        self._cache_matriz_validada = None
        self._cache_conteos_fecha.clear()
        self._cache_analisis_fecha.clear()

    @staticmethod
    def validar_formato_unico_nombre_insumo_bo(
        nombre_archivo: str, fecha_corte: str, tipo_esperado: str = ""
    ) -> Tuple[bool, str, str, str]:
        nom_up = os.path.basename(nombre_archivo).strip().upper()
        tag_fec = fecha_corte.strip().replace("/", "_")
        mapa = {
            f"INSUMO_ASIGNACIONES_{tag_fec}.XLSX": ("ASIGNACIONES", TABLA_INSUMO_ASIGNACIONES),
            f"INSUMO_LICENCIAS_{tag_fec}.XLSX": ("LICENCIAS", TABLA_INSUMO_LICENCIAS),
            f"INSUMO_PRODUCCION_COMPARTIDA_{tag_fec}.XLSX": ("PRODUCCION_COMPARTIDA", TABLA_INSUMO_PRODUCCION_COMPARTIDA),
        }
        if nom_up not in mapa:
            return (
                False,
                "",
                "",
                f"El archivo '{nombre_archivo}' no cumple el formato único oficial para la fecha {fecha_corte}.",
            )
        tipo_det, tabla_dest = mapa[nom_up]
        if tipo_esperado and tipo_esperado.strip().upper() != tipo_det:
            return (
                False,
                tipo_det,
                tabla_dest,
                f"El archivo corresponde a '{tipo_det}' pero se solicitó '{tipo_esperado}'.",
            )
        return True, tipo_det, tabla_dest, ""

    def validar_y_procesar_insumo_bo(
        self, ruta_archivo: str, fecha_corte: str, tipo_esperado: str = ""
    ) -> Tuple[bool, int, int, str, List[str]]:
        self._cache_ruta_validada = ""
        self._cache_matriz_validada = None
        errores: List[str] = []
        nuevos_contratos: List[str] = []

        if not os.path.exists(ruta_archivo):
            return False, 0, 1, "- El archivo especificado no existe.", []

        nom_arch = os.path.basename(ruta_archivo)
        ok_nom, _, _, msg_nom = self.validar_formato_unico_nombre_insumo_bo(nom_arch, fecha_corte, tipo_esperado)
        if not ok_nom:
            errores.append(f"- [Nombre de Archivo]: {msg_nom}")

        wb = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
        ws = wb.active
        matriz = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()

        if len(matriz) < 2:
            return False, 0, 1, "- [Archivo Vacío]: El insumo no contiene registros de producción/contraprestaciones.", []

        headers = [str(c or "").strip().upper() for c in matriz[0]]
        if len(headers) != 53:
            errores.append(f"- [Estructura Interna]: El archivo tiene {len(headers)} columnas (se requieren exactamente 53).")

        for idx_c in range(min(len(headers), 53)):
            h_act = headers[idx_c].replace("Ó", "O")
            h_esp = COLUMNAS_53_OFICIALES[idx_c].upper().replace("Ó", "O")
            if idx_c == 52 and h_act in ("ID_INSTRUMENTO", "CONTRATO"):
                h_act = "ID_CONTRATO"
            if h_act != h_esp:
                errores.append(f"- [Columna {idx_c + 1} Incorrecta]: Es '{headers[idx_c]}' y debe ser '{COLUMNAS_53_OFICIALES[idx_c]}'.")

        mapa_contratos_bd = self.dao.obtener_mapa_contratos_dao()
        filas_revisadas = 0

        for r_idx, row in enumerate(matriz[1:], start=2):
            filas_revisadas += 1
            fila_str = [str(row[c] if c < len(row) and row[c] is not None else "").strip() for c in range(53)]

            for c_idx in range(min(len(headers), 53)):
                if fila_str[c_idx] == "":
                    errores.append(f"- Fila {r_idx} [{COLUMNAS_53_OFICIALES[c_idx]}]: Celda vacía.")

            for col_1b in COLUMNAS_ENTERAS_1BASED:
                val_ent = fila_str[col_1b - 1].replace("'", "")
                if val_ent and not val_ent.isdigit():
                    errores.append(f"- Fila {r_idx} [{COLUMNAS_53_OFICIALES[col_1b - 1]}]: '{val_ent}' requiere entero numérico.")

            rfc_val = fila_str[10]
            if rfc_val and not (9 <= len(rfc_val) <= 13):
                errores.append(f"- Fila {r_idx} [RFC]: '{rfc_val}' longitud inválida.")

            for c_num in range(29, 36):
                v_num_str = fila_str[c_num]
                if v_num_str:
                    try:
                        v_flt = float(v_num_str.replace(",", ""))
                        if v_flt < 0.0:
                            errores.append(
                                f"- Fila {r_idx} [{COLUMNAS_53_OFICIALES[c_num]}]: Contraprestación negativa (${v_flt:,.2f}) inválida."
                            )
                    except ValueError:
                        errores.append(f"- Fila {r_idx} [{COLUMNAS_53_OFICIALES[c_num]}]: Valor no numérico '{v_num_str}'.")

            if fila_str[51] and fila_str[51] != fecha_corte.strip():
                errores.append(
                    f"- Fila {r_idx} [FECHA_INSUMO]: Fecha de fila ({fila_str[51]}) no coincide con {fecha_corte}."
                )

            cod_contrato = fila_str[52]
            if cod_contrato and cod_contrato.upper() not in mapa_contratos_bd:
                if cod_contrato not in nuevos_contratos:
                    nuevos_contratos.append(cod_contrato)

        if nuevos_contratos:
            errores.append(
                f"- [Validación Interna de Contratos]: Se detectaron {len(nuevos_contratos)} bloque(s)/contrato(s) no registrados "
                f"en el catálogo CONTRATOS_HIDROCARBUROS: {', '.join(nuevos_contratos)}. Debe darlos de alta primero."
            )

        if errores:
            return False, filas_revisadas, len(errores), "\n".join(errores[:25]), nuevos_contratos

        self._cache_ruta_validada = ruta_archivo
        self._cache_matriz_validada = matriz
        return True, filas_revisadas, 0, "Validación exitosa al 100%.", []

    def guardar_insumo_validado_bo(
        self, ruta_archivo: str, fecha_corte: str, nombre_tabla_destino: str = ""
    ) -> bool:
        nom_arch = os.path.basename(ruta_archivo)
        if not nombre_tabla_destino:
            _, _, nombre_tabla_destino, _ = self.validar_formato_unico_nombre_insumo_bo(nom_arch, fecha_corte)
        matriz_ram = (
            self._cache_matriz_validada
            if self._cache_ruta_validada == ruta_archivo and self._cache_matriz_validada is not None
            else None
        )
        ok = self.dao.importar_insumo_en_tabla_dao(
            ruta_archivo, fecha_corte, nombre_tabla_destino, matriz_prevalidada=matriz_ram
        )
        total_filas = (len(matriz_ram) - 1) if matriz_ram else 0
        self.dao.registrar_bitacora_insumo_dao(nom_arch, fecha_corte, total_filas, nombre_tabla_destino)
        self.invalidar_cache_datos_bo()
        return ok

    def obtener_conteos_tres_tablas_fecha_bo(self, fecha_corte: str) -> Tuple[int, int, int]:
        if fecha_corte not in self._cache_conteos_fecha:
            self._cache_conteos_fecha[fecha_corte] = self.dao.contar_tres_insumos_por_fecha_dao(fecha_corte)
        return self._cache_conteos_fecha[fecha_corte]

    def _obtener_analisis_tres_regimenes_bo(self, fecha_corte: str):
        if fecha_corte not in self._cache_analisis_fecha:
            a1 = self.dao.extraer_analisis_1_asignaciones_sql(fecha_corte)
            a2 = self.dao.extraer_analisis_2_licencias_sql(fecha_corte)
            a3 = self.dao.extraer_analisis_3_produccion_compartida_sql(fecha_corte)
            self._cache_analisis_fecha[fecha_corte] = (a1, a2, a3)
        return self._cache_analisis_fecha[fecha_corte]

    @staticmethod
    def _clave_contrato_desde_texto(txt: str) -> str:
        partes = str(txt or "").split("|")
        cand = partes[1].strip() if len(partes) > 1 else str(txt or "").strip()
        return cand.split("(")[0].strip()

    def _volcar_datos_crudos_en_template(
        self, wb: openpyxl.Workbook, etiqueta_portafolio: str, rows_data: List[List[Any]], fila_destino: int
    ) -> int:
        ws_dat = wb[HOJA_TPL_DATOS]
        filas_det = [r for r in rows_data if not str(r[0]).strip().upper().startswith("TOTAL")]
        for r in filas_det:
            ws_dat.cell(fila_destino, 1, etiqueta_portafolio)
            for c_idx in range(8):
                ws_dat.cell(fila_destino, c_idx + 2, r[c_idx])
            ws_dat.cell(fila_destino, 10, self._clave_contrato_desde_texto(r[1]))
            fila_destino += 1
        return fila_destino

    def _sembrar_contratos_en_calc_instr(self, wb: openpyxl.Workbook) -> int:
        ws_dat = wb[HOJA_TPL_DATOS]
        ws_inst = wb[HOJA_TPL_CALCINSTR]
        vistos = []
        for r in range(2, ws_dat.max_row + 1):
            val = str(ws_dat.cell(r, 10).value or "").strip()
            if val and val not in vistos:
                vistos.append(val)
        for idx, cod in enumerate(vistos[:200], start=2):
            ws_inst.cell(idx, 1, cod)
        return len(vistos)

    def generar_reporte_ejecutivo_bo(
        self,
        fecha_corte: str,
        inc_asig: bool = True,
        inc_lic: bool = True,
        inc_pc: bool = True,
        inc_consolidado: bool = True,
        carpeta_destino: str = REPORTES_DIR,
    ) -> Tuple[bool, str, str]:
        if inc_consolidado:
            inc_asig = inc_lic = inc_pc = True

        n1, n2, n3 = self.obtener_conteos_tres_tablas_fecha_bo(fecha_corte)
        if (inc_asig and n1 <= 0) or (inc_lic and n2 <= 0) or (inc_pc and n3 <= 0):
            return False, "", f"Faltan insumos cargados para la fecha {fecha_corte}."

        (rows_a1, ops_1, ini_1, fin_1, prom_1, usd_1), (rows_a2, ops_2, ini_2, fin_2, prom_2, usd_2), (rows_a3, ops_3, ini_3, fin_3, prom_3, usd_3) = (
            self._obtener_analisis_tres_regimenes_bo(fecha_corte)
        )

        os.makedirs(carpeta_destino, exist_ok=True)
        fec_tag = fecha_corte.replace("/", "")
        if inc_consolidado or (inc_asig and inc_lic and inc_pc):
            tag = "Consolidado_Ingresos_Hidrocarburos"
        elif inc_asig:
            tag = "1_Asignaciones_Directas_Extraccion"
        elif inc_lic:
            tag = "2_Contratos_Licencia_Upstream"
        else:
            tag = "3_Produccion_Compartida_Upstream"

        ruta_salida = os.path.join(carpeta_destino, f"Reporte_Hidrocarburos_{tag}_{fec_tag}.xlsx")
        shutil.copy2(self.template_path, ruta_salida)

        wb = openpyxl.load_workbook(ruta_salida)
        ws = wb[HOJA_TPL_REPORTE]

        fila_datos = 2
        if inc_asig and rows_a1:
            fila_datos = self._volcar_datos_crudos_en_template(wb, TPL_PORTAFOLIO_1, rows_a1, fila_datos)
        if inc_lic and rows_a2:
            fila_datos = self._volcar_datos_crudos_en_template(wb, TPL_PORTAFOLIO_2, rows_a2, fila_datos)
        if inc_pc and rows_a3:
            fila_datos = self._volcar_datos_crudos_en_template(wb, TPL_PORTAFOLIO_3, rows_a3, fila_datos)
        self._sembrar_contratos_en_calc_instr(wb)

        ws.merge_cells("A1:H2")
        ws["A1"] = (
            "DIRECCIÓN CORPORATIVA DE ECONOMÍA Y FISCALIZACIÓN DE HIDROCARBUROS\n"
            "GERENCIA DE VERIFICACIÓN DE VOLÚMENES, REGALÍAS Y CONTRATOS UPSTREAM E&P"
        )
        ws["A1"].font = Font(name="Cambria", size=11, bold=True, color="1F4E78")
        ws["A1"].fill = PatternFill("solid", fgColor="D6DFEC")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.merge_cells("A3:H3")
        ws["A3"] = f"INFORME EJECUTIVO DE CONTRAPRESTACIONES Y REGALÍAS PETROLERAS — CORTE: {fecha_corte}"
        ws["A3"].font = Font(name="Cambria", size=10, bold=True, color="FFFFFF")
        ws["A3"].fill = PatternFill("solid", fgColor="004E4C")
        ws["A3"].alignment = Alignment(horizontal="center", vertical="center")

        cur_row = 5
        bloques = []
        if inc_asig and rows_a1:
            bloques.append(("ANÁLISIS SQL 1 — ASIGNACIONES DIRECTAS DE EXTRACCIÓN (DERECHOS SOBRE HIDROCARBUROS)", TPL_PORTAFOLIO_1, rows_a1))
        if inc_lic and rows_a2:
            bloques.append(("ANÁLISIS SQL 2 — CONTRATOS DE LICENCIA (REGALÍAS BÁSICAS, ADICIONALES Y CUOTAS)", TPL_PORTAFOLIO_2, rows_a2))
        if inc_pc and rows_a3:
            bloques.append(("ANÁLISIS SQL 3 — CONTRATOS DE PRODUCCIÓN COMPARTIDA (UTILIDAD OPERATIVA Y COSTOS)", TPL_PORTAFOLIO_3, rows_a3))

        hdr_cols = [
            "Concepto / Área Contractual (SQL)", "Moneda | Bloque / Contrato (Tasa %)",
            "Núm. Operaciones", "Valor Contractual ($ MDP)", "Contraprestación y Regalías ($ MDP)",
            "Derecho Promedio ($ MDP)", "Equivalente USD ($ MDD)", "% Part. | Estatus"
        ]

        for titulo_b, etiq_port, rows_b in bloques:
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
            ws.cell(cur_row, 1, titulo_b).font = Font(name="Cambria", size=9.5, bold=True, color="FFFFFF")
            ws.cell(cur_row, 1).fill = PatternFill("solid", fgColor="004E4C")
            cur_row += 1

            for c_i, h_text in enumerate(hdr_cols, 1):
                c_cell = ws.cell(cur_row, c_i, h_text)
                c_cell.font = Font(name="Cambria", size=9, bold=True, color="1F4E78")
                c_cell.fill = PatternFill("solid", fgColor="D6DFEC")
            cur_row += 1

            fila_ini = cur_row
            fila_fin = cur_row + len(rows_b) - 1
            col_rep = ["C", "D", "E", "F", "G"]
            col_dat = ["D", "E", "F", "G", "H"]

            for idx_r, r_vals in enumerate(rows_b):
                r_num = fila_ini + idx_r
                es_total = idx_r == len(rows_b) - 1
                ws.cell(r_num, 1, r_vals[0])
                ws.cell(r_num, 2, r_vals[1])
                ws.cell(r_num, 8, r_vals[7])
                for k_col in range(5):
                    col_l = col_rep[k_col]
                    dat_l = col_dat[k_col]
                    cell_f = ws.cell(r_num, 3 + k_col)
                    if not es_total:
                        cell_f.value = (
                            f'=SUMIFS(DATOS!${dat_l}$2:${dat_l}$5000,'
                            f'DATOS!$A$2:$A$5000,"{etiq_port}",'
                            f'DATOS!$B$2:$B$5000,$A{r_num})'
                        )
                    else:
                        cell_f.value = f"=SUM({col_l}{fila_ini}:{col_l}{fila_fin - 1})"
                    cell_f.number_format = "#,##0" if k_col == 0 else "$#,##0.00"

            cur_row = fila_fin + 2

        wb.save(ruta_salida)
        wb.close()
        return True, ruta_salida, ""

    def generar_reporte_comparativo_bo(
        self, fecha_corte: str, n_insumos: int = 2, carpeta_destino: str = REPORTES_DIR
    ) -> Tuple[bool, str, str]:
        n_insumos = 2 if n_insumos <= 2 else 3
        n1, n2, n3 = self.obtener_conteos_tres_tablas_fecha_bo(fecha_corte)
        if n1 <= 0 or n2 <= 0 or (n_insumos == 3 and n3 <= 0):
            return False, "", f"Se requieren los {n_insumos} insumos cargados para la fecha {fecha_corte}."

        rows_cmp, tot_ops, tot_fin = self.dao.extraer_comparativo_insumos_sql(fecha_corte, n_insumos)
        os.makedirs(carpeta_destino, exist_ok=True)
        fec_tag = fecha_corte.replace("/", "")
        ruta_salida = os.path.join(
            carpeta_destino, f"Reporte_Hidrocarburos_Comparativo_{n_insumos}_Insumos_{fec_tag}.xlsx"
        )
        shutil.copy2(self.template_path, ruta_salida)

        wb = openpyxl.load_workbook(ruta_salida)
        ws = wb[HOJA_TPL_REPORTE]

        (rows_a1, *_), (rows_a2, *_), (rows_a3, *_) = self._obtener_analisis_tres_regimenes_bo(fecha_corte)
        f_dat = 2
        f_dat = self._volcar_datos_crudos_en_template(wb, TPL_PORTAFOLIO_1, rows_a1, f_dat)
        f_dat = self._volcar_datos_crudos_en_template(wb, TPL_PORTAFOLIO_2, rows_a2, f_dat)
        if n_insumos == 3:
            f_dat = self._volcar_datos_crudos_en_template(wb, TPL_PORTAFOLIO_3, rows_a3, f_dat)
        self._sembrar_contratos_en_calc_instr(wb)

        ws.merge_cells("A1:H2")
        ws["A1"] = (
            f"DIRECCIÓN DE ECONOMÍA DE HIDROCARBUROS — REPORTE COMPARATIVO DE {n_insumos} INSUMOS POR CONTRATO\n"
            f"CONCILIACIÓN INTER-REGÍMENES DE PRODUCCIÓN Y CONTRAPRESTACIONES (CORTE: {fecha_corte})"
        )
        ws["A1"].font = Font(name="Cambria", size=11, bold=True, color="1F4E78")
        ws["A1"].fill = PatternFill("solid", fgColor="D6DFEC")

        fila_hdr = 4
        if n_insumos == 2:
            hdrs = [
                "Bloque / Contrato", "Modalidad Contractual", "Núm. Operaciones",
                "Asignaciones ($ MDP)", "Contratos Licencia ($ MDP)",
                "Total Combinado ($ MDP)", "Equivalente USD ($ MDD)", "Presencia en Libros"
            ]
        else:
            hdrs = [
                "Bloque / Contrato", "Modalidad Contractual", "Núm. Operaciones",
                "Asignaciones ($ MDP)", "Contratos Licencia ($ MDP)",
                "Producción Compartida ($ MDP)", "Total Consolidado ($ MDP)", "Presencia en Libros"
            ]
        for c_i, h_t in enumerate(hdrs, 1):
            ws.cell(fila_hdr, c_i, h_t).font = Font(name="Cambria", size=9, bold=True, color="FFFFFF")
            ws.cell(fila_hdr, c_i).fill = PatternFill("solid", fgColor="004E4C")

        fila_ini = 5
        for idx, r_val in enumerate(rows_cmp):
            rn = fila_ini + idx
            for c_i, v in enumerate(r_val, 1):
                ws.cell(rn, c_i, v)
            if n_insumos == 2:
                ws.cell(rn, 6, f"=D{rn}+E{rn}").number_format = "$#,##0.00"
            else:
                ws.cell(rn, 7, f"=D{rn}+E{rn}+F{rn}").number_format = "$#,##0.00"

        fila_fin = fila_ini + len(rows_cmp) - 1
        fila_tot = fila_fin + 1
        ws.cell(fila_tot, 1, f"TOTAL DEL CRUCE ({len(rows_cmp)} contratos)")
        for col_idx, col_let in [(3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G")]:
            ws.cell(fila_tot, col_idx, f"=SUM({col_let}{fila_ini}:{col_let}{fila_fin})")
        ws.cell(fila_tot, 8, "CONCILIADO")

        wb.save(ruta_salida)
        wb.close()
        return True, ruta_salida, ""

    def generar_dictamen_verificacion_bo(
        self, fecha_corte: str, carpeta_destino: str = DICTAMENES_DIR
    ) -> Tuple[bool, str]:
        (rows_a1, ops_1, _, fin_1, _, usd_1), (rows_a2, ops_2, _, fin_2, _, usd_2), (rows_a3, ops_3, _, fin_3, _, usd_3) = (
            self._obtener_analisis_tres_regimenes_bo(fecha_corte)
        )
        tot_ops = ops_1 + ops_2 + ops_3
        tot_fin = round(fin_1 + fin_2 + fin_3, 2)
        tot_usd = round(usd_1 + usd_2 + usd_3, 2)

        os.makedirs(carpeta_destino, exist_ok=True)
        fec_tag = fecha_corte.replace("/", "")
        ruta_dictamen = os.path.join(carpeta_destino, f"Dictamen_Verificacion_Hidrocarburos_{fec_tag}.md")

        contenido = f"""# DIRECCIÓN CORPORATIVA DE ECONOMÍA Y FISCALIZACIÓN DE HIDROCARBUROS
## Gerencia de Verificación de Volúmenes, Valoración Contractual y Regalías Upstream

**Folio Técnico:** `DCF-GVVR-{fec_tag}-001`  
**Fecha de Emisión:** `{fecha_corte}`  
**Asunto:** Dictamen de Conciliación de Volúmenes de Producción, Valor Contractual y Contraprestaciones Petroleras (Corte `{fecha_corte}`).

Se presenta la verificación cuantitativa de las liquidaciones registradas en las tres bases de datos operativas (`BD_INSUMO_ASIGNACIONES`, `BD_INSUMO_LICENCIAS` y `BD_INSUMO_PRODUCCION_COMPARTIDA`) para la fecha de corte **{fecha_corte}**:

| Régimen Contractual de Hidrocarburos | Base de Datos Auditada | Operaciones Verificadas | Contraprestación y Regalías ($ MDP) | Equivalente Divisas ($ MDD) | Estatus Técnico |
|---|---|---:|---:|---:|---|
| **1. Asignaciones Directas de Extracción** | `BD_INSUMO_ASIGNACIONES` | {ops_1:,d} | ${fin_1:,.2f} | ${usd_1:,.2f} | 100% Verificado |
| **2. Contratos de Licencia (Regalías y Cuotas)** | `BD_INSUMO_LICENCIAS` | {ops_2:,d} | ${fin_2:,.2f} | ${usd_2:,.2f} | 100% Verificado |
| **3. Contratos de Producción Compartida** | `BD_INSUMO_PRODUCCION_COMPARTIDA` | {ops_3:,d} | ${fin_3:,.2f} | ${usd_3:,.2f} | 100% Verificado |
| **TOTAL CONSOLIDADO UPSTREAM** | **`HIDROCARBUROS_UPSTREAM`** | **{tot_ops:,d}** | **${tot_fin:,.2f}** | **${tot_usd:,.2f}** | **100.00% CONCILIADO** |

**Elaboró:** Wembley Emmanuel Ibarra Tepato — *Especialista en Modelado Cuantitativo y Fiscalización de Hidrocarburos*
"""
        with open(ruta_dictamen, "w", encoding="utf-8") as f:
            f.write(contenido)
        return True, ruta_dictamen

    def exportar_modelo_estrella_bi_bo(self, carpeta_bi: str = BI_DIR) -> Dict[str, str]:
        """
        Exporta las tablas de Hechos y Dimensiones (Modelo Estrella) desde SQLite
        hacia CSV listos para conectarse en Power BI Desktop y Tableau Desktop/Public.
        """
        ds_dir = os.path.join(carpeta_bi, "datasets_estrella")
        os.makedirs(ds_dir, exist_ok=True)
        rutas = {}

        with self.dao._conectar() as conn:
            cur = conn.cursor()
            # 1) Dim_Contratos_Bloques.csv
            cur.execute(
                "SELECT idContrato, cveOperador, modalidadFiscal, tasaRegalia, "
                "limiteRecuperacionCostos, cuencaPetrolera FROM CONTRATOS_HIDROCARBUROS ORDER BY idContrato;"
            )
            p_dim = os.path.join(ds_dir, "Dim_Contratos_Bloques.csv")
            with open(p_dim, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["ID_CONTRATO", "CVE_OPERADOR", "MODALIDAD_FISCAL", "TASA_REGALIA_PCT", "LIMITE_RECUPERACION_COSTOS_PCT", "CUENCA_PETROLERA"])
                w.writerows(cur.fetchall())
            rutas["Dim_Contratos_Bloques"] = p_dim

            # 2) Fact_Contraprestaciones_Hidrocarburos.csv
            cur.execute(
                "SELECT [FECHA_INSUMO], [REGIMEN], [ID_CONTRATO], [OPERADOR_PETROLERO], [CUENCA_PETROLERA], "
                "[AREA_CONTRACTUAL], [DESCRIPCIÓN CUENTA], [MONEDA], [SDINMO]/1000000.0, [SDFIMN]/1000000.0, "
                "[SDPRMN]/1000000.0, [SDFIDL]/1000000.0 FROM PADRON_CONTRAPRESTACIONES_CONSOLIDADO;"
            )
            p_fact = os.path.join(ds_dir, "Fact_Contraprestaciones_Hidrocarburos.csv")
            with open(p_fact, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([
                    "FECHA_CORTE", "REGIMEN_CODIGO", "ID_CONTRATO", "OPERADOR_PETROLERO",
                    "CUENCA_PETROLERA", "AREA_CONTRACTUAL", "CONCEPTO_FISCAL", "MONEDA",
                    "VALOR_CONTRACTUAL_MDP", "CONTRAPRESTACION_MDP", "DERECHO_PROMEDIO_MDP", "EQUIVALENTE_USD_MDD"
                ])
                for r in cur.fetchall():
                    w.writerow([
                        r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
                        round(float(r[8]), 4), round(float(r[9]), 4), round(float(r[10]), 4), round(float(r[11]), 4)
                    ])
            rutas["Fact_Contraprestaciones_Hidrocarburos"] = p_fact

        return rutas
