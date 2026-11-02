# Hydrocarbon Royalties & Upstream Production Fiscalization Engine (`hydrocarbon-royalties-analytics`)

![Architecture](https://img.shields.io/badge/Architecture-3--Tier_(Presentation_|_BO_|_DAO_+_SQL)-004E4C)
![Stack](https://img.shields.io/badge/Stack-Python_3.11_|_SQLite_|_OpenPyXL_|_Power_BI_|_Tableau_|_VBA-1F4E78)
![Tests](https://img.shields.io/badge/Tests-100%25_Passing-brightgreen)

**Autor:** Wembley Emmanuel Ibarra Tepato ([@Wembley003](https://github.com/Wembley003))

---

## 1. Descripción General

Motor cuantitativo y transaccional en **Arquitectura de 3 Capas** para la **auditoría de volúmenes de producción de hidrocarburos (petróleo crudo, gas natural y condensados), valuación contractual, control de límites de recuperación de costos y liquidación de regalías y contraprestaciones petroleras**.

Incluye tres módulos integrados:
1. **Motor Transaccional y Analítico en Python + SQLite + OpenPyXL (`src/`)** con base de datos física externa (`BD/HIDROCARBUROS_UPSTREAM.sqlite`).
2. **Módulo de Inteligencia de Negocios en Power BI y Tableau (`bi_powerbi_tableau/`)** con exportación automática de Modelo Estrella (`Fact_Contraprestaciones_Hidrocarburos.csv` y `Dim_Contratos_Bloques.csv`), catálogo de medidas DAX (`Medidas_DAX_Hidrocarburos.dax`), tema JSON y libro nativo de Tableau (`Tablero_Ejecutivo_Hidrocarburos.twb`).
3. **Edición de Escritorio en Microsoft Access / VBA Nativo (`vba_access_3capas/`)** con constructor automático en 1 clic, base de datos externa en `BD/` y separación estricta `Presentación -> FondoBO -> FondoDAO -> ConsultasSQL + 10 DTOs`.

---

## 2. Regímenes Contractuales Modelados y Validación de 53 Columnas

1. **Asignaciones Directas de Extracción (`BD_INSUMO_ASIGNACIONES`):** Derechos sobre utilidad compartida, derechos de extracción y derechos de exploración por cuenca y campo.
2. **Contratos de Licencia (`BD_INSUMO_LICENCIAS`):** Regalías básicas progresivas en función del precio e hidrocarburo, regalías adicionales y cuotas contractuales para fase exploratoria.
3. **Contratos de Producción Compartida (`BD_INSUMO_PRODUCCION_COMPARTIDA`):** Límite de recuperación de costos, utilidad operativa compartida y mecanismo de ajuste de rentabilidad ($R	ext{-Factor}$).

### Reglas de Integridad y Plantilla `Templates/Generico.xlsx`
* **Validador Matricial en RAM (`53 columnas`):** Verifica en una sola lectura el formato oficial `INSUMO_<REGIMEN>_DD_MM_AAAA.xlsx`, orden exacto de las 53 columnas, 0 celdas vacías, 20 columnas de códigos enteros, RFC del operador, 7 columnas de valor contractual y contraprestaciones $\ge 0$, consistencia de `FECHA_INSUMO` y existencia de cada bloque/contrato (`ID_CONTRATO`) en el catálogo maestro `CONTRATOS_HIDROCARBUROS`.
* **Inmutabilidad de `Templates/Generico.xlsx`:** El archivo plantilla jamás se abre ni modifica directamente; cada uno de los **6 reportes** (3 individuales por régimen, 1 Consolidado Integral, 1 Comparativo de 2 insumos y 1 Comparativo de 3 insumos) nace de una copia sobre la que se vuelcan los resultados SQL en la hoja `DATOS` y se calculan totales mediante fórmulas vivas `=SUMIFS(...)` y `=SUM(...)`.

---

## 3. Ejecución y Pruebas Automatizadas

```bash
python3 src/ui/consola_ejecutiva.py
python3 -m unittest discover -s tests -p "test_*.py" -v
```
