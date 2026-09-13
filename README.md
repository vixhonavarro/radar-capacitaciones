# Radar de Licitaciones de Capacitación

Informe automático de las licitaciones públicas **abiertas** en [Mercado Público](https://www.mercadopublico.cl) relacionadas con capacitación, cursos, formación y relatorías, con filtros por región, comuna, modalidad y plazo de cierre, y enlace a la ficha oficial de cada licitación.

**No hace scraping**: consulta la [API oficial de Mercado Público](https://api.mercadopublico.cl) (ChileCompra).

## Cómo funciona

- `buscar_licitaciones.ps1` descarga las licitaciones activas, filtra por palabras clave de capacitación (con límites de palabra para evitar falsos positivos), consulta el detalle de cada una (organismo, comuna, montos, fechas) y genera `salida/informe_licitaciones_capacitacion.html` + `salida/licitaciones_capacitacion.csv`.
- Los detalles se cachean en `data/cache_detalles.json`, así cada ejecución solo consulta las licitaciones nuevas.
- El workflow `.github/workflows/actualizar.yml` lo ejecuta de lunes a viernes en la mañana y publica el informe en **GitHub Pages**; también puede lanzarse a mano desde la pestaña *Actions* con "Run workflow".

## Ejecución local (Windows)

Doble clic en `Actualizar_Licitaciones.cmd`, o bien:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File buscar_licitaciones.ps1
```

Compatible con Windows PowerShell 5.1 y PowerShell 7+.

## Ticket de la API

Por defecto se usa el ticket público de prueba de ChileCompra (compartido y con límite de consultas). Para uso productivo, solicite un ticket propio gratuito en [api.mercadopublico.cl](https://api.mercadopublico.cl) y configúrelo:

- **Local**: variable de entorno `MP_TICKET`, o parámetro `-Ticket`.
- **GitHub Actions**: secret del repositorio llamado `MP_TICKET` (Settings → Secrets and variables → Actions).

## Notas

- La comuna/región mostrada corresponde a la unidad compradora; la modalidad (Online/Presencial/Mixta) se detecta automáticamente del texto de cada licitación. El lugar exacto de ejecución se confirma en las bases.
- Los montos y plazos son referenciales; verifique siempre las bases oficiales.
