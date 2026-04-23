# Programa de presupuestos para ferretería

## Versión web para GitHub Pages
El proyecto ya incluye una versión web estática lista para publicar en GitHub Pages:

- `index.html`
- `programa.css`
- `programa.js`
- `manifest.json`
- `service-worker.js`
- `icon.svg`

Esta versión funciona sin Python en el servidor. Lee la lista de precios Excel desde el navegador y exporta el presupuesto también desde el navegador.

Importante: no subas tu lista de precios ni presupuestos reales al repositorio si tienen información privada. El archivo `.gitignore` ya evita subir archivos Excel como `*.xlsx`, `*.xlsm`, `*.xltx` y `*.xltm`.

### Probar la web en tu computadora
```bash
python3 -m http.server 8765
```

Después abrí:
```text
http://127.0.0.1:8765/
```

### Publicar en GitHub Pages desde la web de GitHub
1. Entrá en https://github.com e iniciá sesión.
2. Creá un repositorio nuevo. Puede llamarse, por ejemplo, `presupuestos-ferreteria`.
3. Dejalo como `Public` si usás GitHub Free y querés publicarlo sin configurar planes pagos.
4. Subí estos archivos a la raíz del repositorio:
   - `index.html`
   - `programa.css`
   - `programa.js`
   - `manifest.json`
   - `service-worker.js`
   - `icon.svg`
   - `.nojekyll`
5. No subas archivos Excel con precios o presupuestos reales.
6. En el repositorio, andá a **Settings**.
7. En el menú lateral, entrá a **Pages**.
8. En **Build and deployment**, elegí **Deploy from a branch**.
9. En **Branch**, elegí `main` y carpeta `/(root)`.
10. Guardá con **Save**.
11. Esperá unos minutos. La página quedará en una dirección como:
    ```text
    https://TU_USUARIO.github.io/presupuestos-ferreteria/
    ```

### Publicar usando terminal
Primero creá el repositorio vacío en GitHub. Después, desde esta carpeta:
```bash
git init
git add index.html programa.css programa.js manifest.json service-worker.js icon.svg .nojekyll .gitignore README_presupuestos_ferreteria.md
git commit -m "Crear version web para GitHub Pages"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/presupuestos-ferreteria.git
git push -u origin main
```

Luego activá GitHub Pages en **Settings > Pages > Deploy from a branch > main > /(root)**.

### Instalarla en el celular
Cuando la página esté publicada:
1. Abrí la URL desde el navegador del celular.
2. En Android/Chrome: menú de tres puntos > **Agregar a pantalla principal**.
3. En iPhone/Safari: botón compartir > **Agregar a pantalla de inicio**.

## Qué hace
- Carga tu lista de precios en Excel.
- Detecta automáticamente la fila de encabezados.
- Busca productos por código, nombre, descripción, marca o índice.
- Agrega productos al presupuesto con cantidad.
- Permite usar tres modos de precio:
  - Lista unitario
  - Neto unitario
  - Neto 100 unidades prorrateado por unidad
- Calcula descuento, IVA y recargo/flete.
- Guarda el presupuesto final en un archivo Excel.

## Cómo usarlo
1. Instalá Python 3.10 o superior.
2. Instalá la dependencia:
   ```bash
   pip install -r requirements_presupuestos_ferreteria.txt
   ```
3. Ejecutá el programa:
   ```bash
   python presupuestos_ferreteria.py
   ```
4. Dentro del programa, hacé clic en **Cargar lista de precios** y elegí tu Excel.
5. Buscá productos, agregalos al presupuesto y luego guardalo en Excel.

## Ejecutable
Ya hay un ejecutable creado en:
```bash
dist/PresupuestosFerreteria
```

Podés abrir la carpeta `dist` y hacer doble click en `PresupuestosFerreteria`.
El programa abre una interfaz web local en tu navegador, pero sigue funcionando desde tu computadora.

Si modificás el programa y querés volver a generar el ejecutable, ejecutá:
```bash
./crear_ejecutable.sh
```

La versión con interfaz moderna está en:
```bash
presupuestos_ferreteria_web.py
```

La versión Tkinter original queda disponible en:
```bash
presupuestos_ferreteria.py
```

## También podés abrir el programa ya con una lista cargada
```bash
python presupuestos_ferreteria.py "Lista de Precios CLIENTES 01.11.2025.xlsx"
```

## Estructura esperada del Excel
El programa detecta columnas como estas:
- CODIGO INTERNO
- PRODUCTO
- DESCRIPCION ADICIONAL
- MARCA
- INDICE
- UNIDAD CAJA GRANEL
- UNIDAD CAJA FRACCION
- PRECIO DE LISTA UNITARIO
- PRECIO DE LISTA NETO UNITARIO
- PRECIO DE LISTA NETO 100 UNID

## Siguiente mejora posible
Se puede extender para:
- guardar clientes frecuentes,
- exportar PDF,
- usar base de datos,
- separar vendedores,
- manejar stock,
- generar logo y formato comercial.
