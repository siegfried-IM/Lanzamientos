# Dashboard Seguimiento Lanzamientos - Siegfried

Dashboard interactivo de seguimiento de lanzamientos: market share IQVIA y venta interna,
con cortes por familia, producto y presentación.

## Ver el dashboard online

Una vez activado GitHub Pages, la URL pública será:

**https://siegfried-im.github.io/Lanzamientos/**

(Sirve `dashboard.html` directamente — el archivo es autocontenido, embebe sus propios datos.)

## Cómo activar GitHub Pages (3 clics)

1. En GitHub, ir a este repo → **Settings** → **Pages**.
2. En **Build and deployment** → **Source**, elegir **Deploy from a branch**.
3. En **Branch**, elegir `main` y `/ (root)`. Apretar **Save**.
4. Esperar 1–2 minutos. La URL queda activa.

## Funcionalidad del dashboard

- **IQVIA**: Mercado total + Pack Siegfried + MS% por molécula o ATC.
- **Venta Interna (QLICK)**: packs por familia y presentación.
- **Filtros y vistas**:
  - Selector de Familia (con renombrado: ALIDIAL → Alidial L, ISIS → Isis Nat).
  - Toggle Por Molécula / Por ATC (cambia el denominador del MS%).
  - Toggle Sólo Familia / Con Presentaciones (colapsado vs detallado).
  - Toggle Mostrar / Ocultar evolución mes a mes (filas Δ%).
- **Selector de meses**:
  - Últimos 12 meses como chips siempre visibles.
  - Botones colapsables por año para meses anteriores.
  - Atajos: Últimos 3 / 6 / 12 / Todos / Ninguno.
  - El año más reciente se resalta con borde azul más fuerte.
- **Cabecera con metadata**: última actualización, mes IQVIA más reciente,
  mes Venta Interna más reciente.
- **Headers de tabla fijos (sticky)** al hacer scroll para no perder de vista
  los meses cuando se desliza por contenido largo.

## Actualizar con datos nuevos

`dashboard.html` es estático y autocontenido. Para regenerarlo con datos nuevos:

1. Tener los Excel originales en la misma carpeta:
   `IQUVIA_PM.xlsx`, `QLICK_VTA_INTERNA.xlsx`, `MAESTRO.xlsx`.
2. Correr `build_dashboard.py` (requiere Python + `openpyxl`).
3. Hacer commit y push del nuevo `dashboard.html`.

> Los Excel **no** se versionan en este repo (están en `.gitignore`)
> porque contienen datos de mercado y ventas internas.

## Tecnología

- HTML estático con JavaScript vanilla (sin frameworks).
- Datos embebidos como JSON dentro del HTML.
- Cero dependencias en tiempo de ejecución — abrible directo en cualquier
  navegador moderno, también sin internet una vez descargado.
