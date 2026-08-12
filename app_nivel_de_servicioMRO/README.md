Dx Compradores - Streamlit
Réplica en Python del pbix `Nivel_de_servicio_BI.pbix`, página "Dx Compradores".
Estructura
`src/config.py` — rutas y parámetros de negocio (SLA, fecha de corte)
`src/loaders.py` — carga de datos, uno por query M del pbix
`src/transform.py` — merges y cálculos (equivalente a columnas/medidas DAX)
`src/app.py` — la app Streamlit
Pendiente de confirmar (ver TODO en el código)
Fórmula real de `% Cumplimiento`, `Nivel de Servicio`, `Estado Solped`, `Aplica?`
Clave de cruce real de `Responsable de MRP`
Nombres de columna exactos en los 3 archivos de SharePoint (hoy son placeholders)
Correr local
```
pip install -r requirements.txt
# coloca ME5A_con_Ariba.xlsx y los 3 excels de responsables en /data
streamlit run src/app.py
```
Migración a Azure (cuando el ticket de TI esté aprobado)
Solo se tocan las funciones de `loaders.py`:
`cargar_data_pr` → leer desde Blob Storage
`cargar_responsable_*` → leer desde SharePoint vía Microsoft Graph API
El resto del código (transform.py, app.py) no cambia.
