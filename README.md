# AutoPrice DGEHM

Recolector diario del reporte oficial de precios de combustibles de la DGEHM de El Salvador.

## Funcionamiento

1. GitHub Actions abre el reporte oficial con Chromium.
2. Exporta el reporte en Excel.
3. Detecta dinámicamente los encabezados y transforma los datos.
4. Rechaza cargas incompletas con menos de 100 estaciones.
5. Guarda `data/latest.json`, `data/latest.xlsx` y una copia diaria en `data/history/`.

La ejecución automática está programada para las 6:30 a. m. de El Salvador. También puede iniciarse manualmente desde **Actions → Descargar precios DGEHM → Run workflow**.

Fuente oficial: https://sinapp.dgehm.gob.sv/drhm/estadisticas.aspx?uid=2

## Acceso mediante proxy salvadoreño

Si DGEHM bloquea las direcciones de GitHub, configure estos secretos en
**Settings → Secrets and variables → Actions**:

- `PROXY_SERVER`: servidor y puerto del proxy, por ejemplo `gate.example.com:7000`.
- `PROXY_USERNAME`: usuario generado por el proveedor con país El Salvador.
- `PROXY_PASSWORD`: contraseña del proxy.

El recolector utiliza el proxy solamente cuando `PROXY_SERVER` está definido.
