# Módulo Ejecutivo de Business Intelligence (`bi_powerbi_tableau/`)

Arquitectura analítica en modelo estrella para **Power BI** y **Tableau Desktop** aplicada a la fiscalización de producción *Upstream* y liquidación de contraprestaciones petroleras:

* **`datasets_estrella/Fact_Contraprestaciones_Hidrocarburos.csv` y `Dim_Contratos_Bloques.csv`**: Tablas de hechos y dimensiones por contrato/asignación, cuenca, fluido, precio marcador, costo recuperable y contraprestaciones a favor del Estado.
* **`powerbi/Medidas_DAX_Hidrocarburos.dax`**: Librería de medidas DAX (*Government Take %*, *Valor Contractual Total*, *Recuperación de Costos Auditada*, *Regalías Base* y *Contraprestación sobre Utilidad Operativa*).
* **`powerbi/Tema_Ejecutivo_Hidrocarburos.json`**: Tema institucional para Power BI Desktop.
* **`tableau/Tablero_Ejecutivo_Hidrocarburos.twb`**: Libro nativo de Tableau Desktop con cascada fiscal y comparativo por modalidad contractual.
