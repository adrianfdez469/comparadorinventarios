# InventoryComparer

Compara cantidades entre dos inventarios definidos en un mismo libro Excel y genera un informe en HTML (`comparacion_inventario.html`). El Excel **no se modifica**.

---

## Archivo de entrada

El libro debe llamarse **`data.xlsx`** y estar en la **carpeta raíz** del proyecto (junto a los scripts), salvo que uses la opción `--xlsx` del script Python.

El script **solo lee** el Excel (no lo modifica ni lo borra) y escribe únicamente el HTML de informe. Si indicas `--out` con la misma ruta que el Excel, el programa se detiene para no sobrescribir el fichero.

Debe tener **al menos tres hojas**. El programa usa el **orden de las pestañas** (primera, segunda y tercera hoja), **no** el nombre que tenga cada hoja.

---

## Fila de encabezados

En las **tres hojas**, la **fila 1** se trata como cabecera y **no** se usa como dato. Los productos y correlaciones empiezan en la **fila 2**.

---

## Hoja 1 (primera pestaña)

| Columna A | Columna B |
|-----------|-----------|
| Nombre del producto | Cantidad |

- Los nombres deben coincidir **exactamente** (incluidos espacios y mayúsculas) con lo que pongas en la columna A de la **hoja de correlación**.
- La cantidad puede ser un **número** o un **texto** que contenga un número; el programa toma el **primer número** que encuentre. Ejemplos válidos: `5`, `5 pzs.`, `350 pzs`, `12,5 kg`.
- Si en la hoja 1 hay **varias filas con el mismo nombre**, solo cuenta la **última** fila leída.

---

## Hoja 2 (segunda pestaña)

| Columna A | Columna B |
|-----------|-----------|
| Nombre del producto | Cantidad |

- Misma disposición que la hoja 1: nombre en A, cantidad en B.
- Lo habitual es que la cantidad sea un **número**; si es texto con número, también se interpreta.
- Para la hoja 2 se usa formato numérico tipo **punto miles y coma decimales** (ejemplo: `11.835,00` = `11835`).
- Misma regla de **nombres duplicados**: gana la **última** fila.

---

## Hoja 3 (tercera pestaña) — correlación de nombres

| Columna A | Columna B |
|-----------|-----------|
| Nombre tal como está en la **hoja 1** | Nombre tal como está en la **hoja 2** |

- Cada **fila de datos** (desde la 2) define **un par** a comparar: se busca la cantidad del nombre de A en la hoja 1 y la del nombre de B en la hoja 2, y se comparan.
- Si la columna A está **vacía**, esa fila se **ignora**.
- Si la columna A tiene texto pero la B está vacía, la comparación saldrá en error en el informe (falta el nombre en hoja 2).
- El informe lista las filas de correlación y además agrega productos que existan solo en una hoja (sin pareja en la otra), comparando contra `0`.

---

## Comparación

- Se comparan los valores numéricos; se considera coincidencia si la diferencia es **muy pequeña** (tolerancia numérica, por decimales de Excel).
- Si falta un producto, falta cantidad interpretable o los números no coinciden, el resultado es **Fail**; en el HTML puedes **hacer clic** en Fail para ver el **motivo** y la diferencia.
- Si un producto existe solo en una hoja, se compara contra `0`: si ese valor es `0` se marca **OK**, en caso contrario **Fail**.

---

## Cómo ejecutarlo

- **Python:** `python compare_inventory.py` (opcional: `--xlsx ruta.xlsx --out informe.html`). Requisito: `pip install -r requirements.txt`.
- **macOS:** doble clic en `run_comparacion.command` (genera el HTML y lo abre).
- **Windows:** doble clic en `run_comparacion.bat`.
- **Web (sin backend):** abre `index.html`, sube el Excel y pulsa **Comparar**.
