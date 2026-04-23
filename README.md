# Presupuestos Ferretería

Aplicación web para crear presupuestos de ferretería desde una lista de precios en Excel.

La versión publicada en GitHub Pages funciona en el navegador: carga el Excel localmente, permite buscar productos, armar el presupuesto y exportar el resultado como `.xlsx`.

## Archivos principales
- `index.html`: estructura de la aplicación web.
- `programa.css`: estilos visuales y responsive.
- `programa.js`: lectura de Excel, búsqueda, cálculo y exportación.
- `manifest.json`, `service-worker.js`, `icon.svg`: soporte básico para instalarla en celular como PWA.

## Publicación
Activar GitHub Pages desde:

`Settings > Pages > Build and deployment > Deploy from a branch > main > /(root)`

La URL quedará con este formato:

`https://FacuGod.github.io/presupuesto-ferreteria/`

## Privacidad
No subas listas de precios ni presupuestos reales al repositorio. El archivo `.gitignore` evita subir archivos Excel por defecto.

Más detalles en `README_presupuestos_ferreteria.md`.
