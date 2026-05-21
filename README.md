# Dashboard Seguimiento Lanzamientos - Siegfried

Dashboard interactivo de seguimiento de lanzamientos: market share IQVIA y venta interna,
con cortes por familia, producto y presentación.

> **Repo privado.** El acceso está limitado a miembros de la organización
> `siegfried-IM`. No hay URL pública del dashboard (GitHub Pages no está
> disponible en el plan actual para repos privados).

## Cómo ver el dashboard desde otra PC (miembros del equipo)

### Opción A — Descargar y abrir (más simple)

1. Entrá al repo en GitHub: https://github.com/siegfried-IM/Lanzamientos
2. Hacé clic en **`dashboard.html`** dentro de la lista de archivos.
3. Arriba a la derecha, hacé clic en el icono **Download raw file** (↓).
4. Abrí el archivo descargado con tu navegador (doble clic).

El dashboard funciona offline — todos los datos están embebidos en el HTML.
Es 100% interactivo (filtros, toggles, selector de meses, etc.).

### Opción B — Clonar el repo (para los que actualizan)

```bash
git clone https://github.com/siegfried-IM/Lanzamientos.git
cd Lanzamientos
# Abrir dashboard.html con doble clic
```

## Funcionalidad del dashboard

- **IQVIA**: Mercado total + Pack Siegfried + MS% por molécula o ATC.
- **Venta Interna (QLICK)**: packs por familia y presentación.
- **Filtros y vistas**:
  - Selector de Familia (renombrado visual: ALIDIAL → Alidial L, ISIS → Isis Nat).
  - Toggle Por Molécula / Por ATC (cambia el denominador del MS%).
  - Toggle Sólo Familia / Con Presentaciones.
  - Toggle Mostrar / Ocultar evolución mes a mes (filas Δ%).
- **Selector de meses**:
  - Últimos 12 meses como chips siempre visibles.
  - Botones colapsables por año para meses anteriores (sólo años con ≥3 productos activos).
  - Atajos: Últimos 3 / 6 / 12 / Todos / Ninguno.
  - El año más reciente se resalta con borde azul más fuerte.
- **Cabecera con metadata**: última actualización, mes IQVIA más reciente,
  mes Venta Interna más reciente.
- **Headers de tabla fijos (sticky)** al hacer scroll.

## Actualizar con datos nuevos

`dashboard.html` es autocontenido — embebe sus propios datos. Para regenerarlo
con datos nuevos (sólo lo hace quien tiene los Excel locales):

1. Actualizá los Excel en la misma carpeta:
   `IQUVIA_PM.xlsx`, `QLICK_VTA_INTERNA.xlsx`, `MAESTRO.xlsx`.
2. Doble clic en **`actualizar_dashboard.bat`** → regenera `dashboard.html`.
3. Doble clic en **`subir_a_github.bat`** → commit + push del nuevo HTML.

> Los Excel **no** se versionan en este repo (están en `.gitignore`)
> porque contienen datos de mercado y ventas internas.

## Tecnología

- HTML estático con JavaScript vanilla (sin frameworks).
- Datos embebidos como JSON dentro del HTML.
- Cero dependencias en tiempo de ejecución — abrible directo en cualquier
  navegador moderno, también offline.
- Generador: Python + openpyxl (`build_dashboard.py`).

## Archivos

| Archivo | Para qué |
|---|---|
| `dashboard.html` | El dashboard interactivo (lo que se ve) |
| `index.html` | Redirect a `dashboard.html` (por si clonan el repo) |
| `build_dashboard.py` | Generador desde los Excel |
| `actualizar_dashboard.bat` | Doble clic para regenerar localmente |
| `subir_a_github.bat` | Doble clic para subir cambios al repo |
| `COMO_ACTUALIZAR.txt` | Guía paso a paso para actualizaciones |
| `.gitignore` | Excluye `.xlsx` (datos sensibles no se versionan) |
