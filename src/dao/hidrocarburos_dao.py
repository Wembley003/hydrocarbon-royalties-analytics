"""
CAPA 3 (PERSISTENCIA - DATA ACCESS OBJECT): hidrocarburos_dao.py
=============================================================================
Única clase del sistema autorizada para abrir conexiones hacia la base de datos
externa BD/HIDROCARBUROS_UPSTREAM.sqlite (fuera del módulo principal).
"""
import os
import sqlite3
from typing import List, Tuple, Optional, Any
import openpyxl

from src.constantes import (
    DB_PATH, COLUMNAS_53_OFICIALES,
    TABLA_INSUMO_ASIGNACIONES, TABLA_INSUMO_LICENCIAS,
    TABLA_INSUMO_PRODUCCION_COMPARTIDA, TABLA_CONSOLIDADA_CONTRAPRESTACIONES
)
from src.dto.modelos_dto import ContratoDTO
from src.dao import consultas_sql


CATALOGO_CONTRATOS_SEED = [
    ("BLK-ASIG-0142-SURESTE", "OPERADORA_ESTATAL_EP", "Asignación Directa Extracción Terrestre", 65.00, 11.00, "Cuenca del Sureste"),
    ("BLK-ASIG-0151-MARINA", "OPERADORA_ESTATAL_EP", "Asignación Directa Aguas Someras", 65.00, 11.00, "Sonda Marina"),
    ("BLK-ASIG-0155-MADURO", "OPERADORA_ESTATAL_EP", "Asignación Directa Extracción Madura", 65.00, 11.00, "Sonda Marina"),
    ("BLK-ASIG-0162-GASCON", "OPERADORA_ESTATAL_EP", "Asignación Gas y Condensados Terrestre", 65.00, 12.50, "Cuenca Golfo Centro"),
    ("BLK-ASIG-0170-ALTPRE", "OPERADORA_ESTATAL_EP", "Asignación Condensados Alta Presión", 65.00, 12.50, "Cuenca del Sureste"),
    ("BLK-ASIG-0185-LIGERO", "OPERADORA_ESTATAL_EP", "Asignación Crudo Ligero Terrestre", 65.00, 11.00, "Cuenca del Sureste"),
    ("BLK-PSC-R01-A02", "CONSORCIO_TALOS_EP", "Contrato Producción Compartida Somero", 68.40, 60.00, "Cuenca del Sureste"),
    ("BLK-PSC-R01-A07", "CONSORCIO_TALOS_EP", "Contrato Producción Compartida Unificado", 68.99, 60.00, "Cuenca del Sureste"),
    ("BLK-PSC-R01-A01", "ENERGIA_COSTERA_SA", "Contrato Producción Compartida Marino", 83.75, 60.00, "Cuenca del Sureste"),
    ("BLK-PSC-R01-A03", "HOKCHI_OFFSHORE_SA", "Contrato Producción Compartida Plataforma", 70.00, 60.00, "Cuenca del Sureste"),
    ("BLK-PSC-R01-A04", "FIELDWOOD_MARINA_SA", "Contrato Producción Compartida Desarrollo", 74.00, 60.00, "Sonda Marina"),
    ("BLK-LIC-R01-A01", "DIAVAZ_TERRESTRE_SA", "Contrato Licencia Terrestre Campos Maduros", 22.50, 100.00, "Cuenca Norte"),
    ("BLK-LIC-R01-A06", "PERSEUS_ENERGIA_SA", "Contrato Licencia Terrestre Crudo Pesado", 24.80, 100.00, "Cuenca del Sureste"),
    ("BLK-LIC-R01-P01", "WOODSIDE_DEEP_SA", "Contrato Licencia Aguas Profundas Norte", 17.50, 100.00, "Cinturón Plegado Profundo"),
    ("BLK-LIC-R01-P03", "OCEANIA_DEEP_SA", "Contrato Licencia Aguas Profundas Sur", 18.20, 100.00, "Cinturón Plegado Profundo"),
    ("BLK-LIC-R01-P06", "CARIGALI_OFFSHORE_SA", "Contrato Licencia Aguas Profundas Salina", 16.90, 100.00, "Cuenca Salina"),
    ("BLK-PSC-R02-A02", "WINTERSHALL_EP_SA", "Contrato Producción Compartida Somero", 66.50, 60.00, "Cuenca Tampico"),
    ("BLK-PSC-R02-A06", "ENERGIA_COSTERA_SA", "Contrato Producción Compartida Somero II", 75.10, 60.00, "Cuenca del Sureste"),
    ("BLK-LIC-R02-A01", "JAGUAR_GAS_SA", "Contrato Licencia Terrestre Gas Húmedo", 21.30, 100.00, "Cuenca Golfo Centro"),
    ("BLK-LIC-R02-A04", "IBEROAMERICANA_GAS", "Contrato Licencia Terrestre Gas Seco", 19.80, 100.00, "Cuenca Norte"),
    ("BLK-LIC-R02-P02", "OCEANIA_DEEP_SA", "Contrato Licencia Aguas Ultra Profundas", 18.50, 100.00, "Cordilleras Marinas"),
    ("BLK-PSC-R03-A08", "IBERIA_EXPLORACION", "Contrato Producción Compartida Somero III", 62.40, 60.00, "Cuenca Tampico"),
    ("BLK-MIG-001-MARINO", "OPERADORA_ESTATAL_EP", "Contrato Migrado Producción Compartida", 71.50, 60.00, "Sonda Marina"),
    ("BLK-ASO-001-TERRESTRE", "CHEIRON_PETROLEO_SA", "Contrato Asociación Terrestre", 64.20, 60.00, "Cuenca del Sureste"),
    ("BLK-ASO-002-TERRESTRE", "WINTERSHALL_EP_SA", "Contrato Asociación Campo Maduro", 63.80, 60.00, "Cuenca del Sureste"),
]


class HidrocarburosDAO:
    _backend_inicializado: bool = False

    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self.verificar_y_crear_backend_dao()

    def _conectar(self) -> sqlite3.Connection:
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA journal_mode = MEMORY;")
        return conn

    def verificar_y_crear_backend_dao(self) -> bool:
        if HidrocarburosDAO._backend_inicializado and os.path.exists(self.db_path):
            return True

        with self._conectar() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS USUARIOS (
                    cveUsuario TEXT PRIMARY KEY,
                    password TEXT NOT NULL,
                    nombreUsuario TEXT NOT NULL,
                    aPaterno TEXT NOT NULL,
                    aMaterno TEXT,
                    cvePerfil TEXT NOT NULL,
                    emailUsuario TEXT,
                    cveEstatus TEXT NOT NULL,
                    cvePuesto INTEGER
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS CONTRATOS_HIDROCARBUROS (
                    idContrato TEXT PRIMARY KEY,
                    cveOperador TEXT,
                    modalidadFiscal TEXT,
                    tasaRegalia REAL,
                    limiteRecuperacionCostos REAL,
                    cuencaPetrolera TEXT
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS BITACORA_CARGA_INSUMOS (
                    idCarga INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombreArchivo TEXT,
                    fechaCorte TEXT,
                    fechaHoraCarga TEXT,
                    totalRegistros INTEGER,
                    regimenesActualizados TEXT
                );
            """)

            cols_ddl = []
            for idx_c, col_name in enumerate(COLUMNAS_53_OFICIALES, 1):
                col_type = "REAL" if 30 <= idx_c <= 36 else "TEXT"
                cols_ddl.append(f"[{col_name}] {col_type}")
            ddl_53 = ", ".join(cols_ddl)

            for tbl_name in [
                TABLA_INSUMO_ASIGNACIONES,
                TABLA_INSUMO_LICENCIAS,
                TABLA_INSUMO_PRODUCCION_COMPARTIDA,
                TABLA_CONSOLIDADA_CONTRAPRESTACIONES,
            ]:
                cur.execute(f"CREATE TABLE IF NOT EXISTS [{tbl_name}] ({ddl_53});")
                cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{tbl_name}_fecha ON [{tbl_name}] ([FECHA_INSUMO]);")
                cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{tbl_name}_contrato ON [{tbl_name}] ([ID_CONTRATO]);")

            cur.execute(
                f"CREATE INDEX IF NOT EXISTS idx_cons_reg_fec ON [{TABLA_CONSOLIDADA_CONTRAPRESTACIONES}] ([FECHA_INSUMO], [REGIMEN]);"
            )
            cur.execute(
                "CREATE INDEX IF NOT EXISTS idx_bit_fec ON BITACORA_CARGA_INSUMOS (fechaCorte, nombreArchivo);"
            )

            cur.execute("SELECT COUNT(*) FROM USUARIOS;")
            if cur.fetchone()[0] == 0:
                cur.execute(
                    "INSERT INTO USUARIOS VALUES ('1000', 'ENC$E67D64692A29C0BB', 'Wembley Emmanuel', "
                    "'Ibarra', 'Tepato', 'A', 'wembley.ibarra@energia-upstream.org', 'ACTIVO', 1);"
                )
            cur.execute("SELECT COUNT(*) FROM CONTRATOS_HIDROCARBUROS;")
            if cur.fetchone()[0] == 0:
                cur.executemany(
                    "INSERT INTO CONTRATOS_HIDROCARBUROS VALUES (?, ?, ?, ?, ?, ?);",
                    CATALOGO_CONTRATOS_SEED,
                )
            conn.commit()

        HidrocarburosDAO._backend_inicializado = True
        return True

    def obtener_mapa_contratos_dao(self) -> set:
        with self._conectar() as conn:
            cur = conn.cursor()
            cur.execute(consultas_sql.sql_select_contratos())
            return {str(r[0]).strip().upper() for r in cur.fetchall() if r[0]}

    def contar_tres_insumos_por_fecha_dao(self, fecha_corte: str) -> Tuple[int, int, int]:
        n1 = n2 = n3 = 0
        with self._conectar() as conn:
            cur = conn.cursor()
            cur.execute(
                consultas_sql.sql_contar_tres_insumos_por_fecha(),
                (fecha_corte, fecha_corte, fecha_corte),
            )
            for idx_reg, cnt in cur.fetchall():
                if idx_reg == 1:
                    n1 = int(cnt)
                elif idx_reg == 2:
                    n2 = int(cnt)
                elif idx_reg == 3:
                    n3 = int(cnt)
        return n1, n2, n3

    def importar_insumo_en_tabla_dao(
        self,
        ruta_archivo: str,
        fecha_corte: str,
        nombre_tabla_destino: str,
        matriz_prevalidada: Optional[List[List[Any]]] = None,
    ) -> bool:
        tbl_dest = nombre_tabla_destino.strip() or TABLA_INSUMO_ASIGNACIONES
        if tbl_dest == TABLA_INSUMO_ASIGNACIONES:
            regimen_filtro = "PRIMERO"
        elif tbl_dest == TABLA_INSUMO_LICENCIAS:
            regimen_filtro = "TESORERIA"
        else:
            regimen_filtro = "SEGUNDO"

        if matriz_prevalidada is None:
            wb = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
            ws = wb.active
            matriz_prevalidada = [list(row) for row in ws.iter_rows(values_only=True)]
            wb.close()

        filas_datos = []
        for row in matriz_prevalidada[1:]:
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue
            fila_limpia = []
            for c_idx in range(53):
                val = row[c_idx] if c_idx < len(row) else ""
                if 29 <= c_idx <= 35:
                    try:
                        fila_limpia.append(float(str(val).replace(",", "")) if val is not None else 0.0)
                    except ValueError:
                        fila_limpia.append(0.0)
                elif c_idx == 51 and fecha_corte:
                    fila_limpia.append(fecha_corte)
                else:
                    fila_limpia.append("" if val is None else str(val).strip())
            filas_datos.append(tuple(fila_limpia))

        placeholders = ", ".join(["?"] * 53)
        with self._conectar() as conn:
            cur = conn.cursor()
            cur.execute(consultas_sql.sql_delete_insumo_por_fecha(tbl_dest), (fecha_corte,))
            cur.execute(
                consultas_sql.sql_delete_consolidada_por_fecha_y_regimen(),
                (fecha_corte, regimen_filtro),
            )
            cur.executemany(f"INSERT INTO [{tbl_dest}] VALUES ({placeholders});", filas_datos)
            cur.execute(consultas_sql.sql_copiar_insumo_a_consolidada(tbl_dest), (fecha_corte,))
            conn.commit()
        return True

    def registrar_bitacora_insumo_dao(
        self, nombre_archivo: str, fecha_corte: str, total_reg: int, regimenes_str: str
    ) -> bool:
        with self._conectar() as conn:
            conn.execute(
                consultas_sql.sql_insert_bitacora(),
                (nombre_archivo, fecha_corte, int(total_reg), regimenes_str),
            )
            conn.commit()
        return True

    def _ejecutar_analisis_agregado(
        self, sql_query: str, fecha_corte: str, etiqueta_otros: str, etiqueta_total: str, sufijo_estatus: str
    ) -> Tuple[List[List[Any]], int, float, float, float, float]:
        with self._conectar() as conn:
            cur = conn.cursor()
            cur.execute(sql_query, (fecha_corte,))
            raw_rows = cur.fetchall()

        if not raw_rows:
            return [], 0, 0.0, 0.0, 0.0, 0.0

        tot_ops = sum(int(r[2]) for r in raw_rows)
        tot_ini = sum(float(r[3]) for r in raw_rows)
        tot_fin = sum(float(r[4]) for r in raw_rows)
        tot_prom = sum(float(r[5]) for r in raw_rows)
        tot_usd = sum(float(r[6]) for r in raw_rows)

        top_items = []
        o_ops = 0
        o_ini = o_fin = o_prom = o_usd = 0.0
        for idx, r in enumerate(raw_rows, 1):
            if idx <= 4:
                top_items.append([f"{idx}. {r[0]}", r[1], int(r[2]), float(r[3]), float(r[4]), float(r[5]), float(r[6])])
            else:
                o_ops += int(r[2])
                o_ini += float(r[3])
                o_fin += float(r[4])
                o_prom += float(r[5])
                o_usd += float(r[6])

        if o_ops > 0:
            top_items.append([
                f"{len(top_items) + 1}. {etiqueta_otros}",
                "M.N. / USD | Varios Contratos",
                o_ops, o_ini, o_fin, o_prom, o_usd
            ])

        rows_out = []
        for item in top_items:
            pct = (item[4] / tot_fin * 100.0) if tot_fin > 0 else 0.0
            rows_out.append([
                item[0], item[1], item[2],
                round(item[3], 2), round(item[4], 2), round(item[5], 2), round(item[6], 2),
                f"{pct:.2f}% | {sufijo_estatus}"
            ])

        rows_out.append([
            etiqueta_total, f"Insumo {fecha_corte}", tot_ops,
            round(tot_ini, 2), round(tot_fin, 2), round(tot_prom, 2), round(tot_usd, 2),
            "100.00% CONCILIADO"
        ])
        return rows_out, tot_ops, tot_ini, tot_fin, tot_prom, tot_usd

    def extraer_analisis_1_asignaciones_sql(self, fecha_corte: str):
        return self._ejecutar_analisis_agregado(
            consultas_sql.sql_reporte_analisis_1_asignaciones(),
            fecha_corte,
            "OTROS DERECHOS DE EXTRACCION",
            "TOTAL 1. ASIGNACIONES DIRECTAS (SQL)",
            "VERIFICADO E1",
        )

    def extraer_analisis_2_licencias_sql(self, fecha_corte: str):
        return self._ejecutar_analisis_agregado(
            consultas_sql.sql_reporte_analisis_2_licencias(),
            fecha_corte,
            "OTRAS AREAS DE LICENCIA",
            "TOTAL 2. CONTRATOS DE LICENCIA (SQL)",
            "EN CUMPLIMIENTO",
        )

    def extraer_analisis_3_produccion_compartida_sql(self, fecha_corte: str):
        return self._ejecutar_analisis_agregado(
            consultas_sql.sql_reporte_analisis_3_produccion_compartida(),
            fecha_corte,
            "OTROS BLOQUES PRODUCCION COMPARTIDA",
            "TOTAL 3. PRODUCCIÓN COMPARTIDA (SQL)",
            "CONCILIADO E1",
        )

    def extraer_comparativo_insumos_sql(self, fecha_corte: str, n_insumos: int = 2):
        with self._conectar() as conn:
            cur = conn.cursor()
            if n_insumos == 2:
                cur.execute(consultas_sql.sql_comparativo_dos_insumos(), (fecha_corte, fecha_corte))
            else:
                cur.execute(
                    consultas_sql.sql_comparativo_tres_insumos(),
                    (fecha_corte, fecha_corte, fecha_corte),
                )
            raw = cur.fetchall()

        rows_out = []
        tot_ops = 0
        tot_fin_mdp = 0.0
        for r in raw:
            clave = str(r[0])
            modalidad = str(r[1])
            if n_insumos == 2:
                ops_asig = int(r[2])
                ops_lic = int(r[3])
                ops = ops_asig + ops_lic
                f1 = round(float(r[4]), 2)
                f2 = round(float(r[5]), 2)
                f_comb = round(f1 + f2, 2)
                usd = round(float(r[7]), 2)
                libros = (1 if ops_asig > 0 else 0) + (1 if ops_lic > 0 else 0)
                pres = "AMBOS REGÍMENES" if libros == 2 else ("SOLO ASIGNACIONES" if f1 > 0 else "SOLO LICENCIA")
                rows_out.append([clave, modalidad, ops, f1, f2, f_comb, usd, pres])
                tot_ops += ops
                tot_fin_mdp += f_comb
            else:
                ops = int(r[2])
                f1 = round(float(r[3]), 2)
                f2 = round(float(r[4]), 2)
                f3 = round(float(r[5]), 2)
                f_cons = round(f1 + f2 + f3, 2)
                libros = int(r[7])
                pres = f"{libros} de 3 regímenes"
                rows_out.append([clave, modalidad, ops, f1, f2, f3, f_cons, pres])
                tot_ops += ops
                tot_fin_mdp += f_cons

        return rows_out, tot_ops, round(tot_fin_mdp, 2)
