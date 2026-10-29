# Historial Manager — Contexto del Proyecto

## Qué hace la app

Aplicación de escritorio para automatizar la descarga del historial de facturas de UTE (empresa eléctrica uruguaya) para múltiples clientes empresariales. Descarga PDFs del portal de UTE, los parsea y genera reportes Excel consolidados por subcuenta.

## Stack tecnológico

- **Lenguaje:** Python 3.12+
- **GUI:** PyWebView 6.1 (ventana desktop con HTML/CSS/JS adentro)
- **Scraping:** Playwright 1.58.0 (Chromium headless)
- **PDF parsing:** pdfplumber 0.11.9, pdfminer.six, pypdfium2
- **Excel:** openpyxl 3.1.5
- **Data:** pandas 3.0.1
- **DB:** SQLite 3 (sin ORM externo, wrapper propio)
- **Frontend:** HTML5 + CSS3 + JS vanilla (sin frameworks)
- **Servidor interno:** bottle 0.13.4

## Estructura de archivos

```
Hisotrial_manager/
├── src/
│   ├── historial_manager.py       # Entry point: PyWebView + bridge Python↔JS (195 líneas)
│   ├── historial_facturas_ute.py  # Motor de scraping Playwright del portal UTE (877 líneas)
│   ├── parsing_facturas.py        # Parseo de PDFs de factura (TH/GS) → dicts (892 líneas)
│   ├── excel_export.py            # Generación del Excel consolidado con openpyxl (128 líneas)
│   └── database.py                # Wrapper SQLite (205 líneas)
├── ui/
│   └── index.html                 # UI completa (594 líneas)
├── data/
│   └── seg_historial.db           # Base de datos SQLite
├── setup.bat                      # Setup inicial (instala Python, venv, deps, Chromium)
├── Iniciar.bat                    # Lanzador de la app
└── requirements.txt               # Dependencias pinneadas
```

### Separación de responsabilidades del motor

El motor está dividido en 3 módulos con dependencias en una sola dirección
(sin ciclos): `historial_facturas_ute.py → excel_export.py → parsing_facturas.py`.

- **`parsing_facturas.py`** — funciones **puras** de parseo de PDF (reciben
  `pdf_path`/texto, devuelven dicts). No tocan Playwright ni la DB. Es lo más
  testeable de forma aislada (`parse_factura_historial`, `extract_*` TH y GS).
- **`historial_facturas_ute.py`** — solo scraping con Playwright (iframes,
  reintentos, reenganche de frames). Código delicado construido a prueba y error:
  **no mezclar parsing ni Excel acá de nuevo**. Importa lo que necesita de los
  otros dos módulos.
- **`excel_export.py`** — generación del Excel con openpyxl (`guardar_excel_consolidado`).

## Cómo se lanza

```
Iniciar.bat → .venv\Scripts\python.exe src\historial_manager.py
```

Para setup inicial: `setup.bat` (instala Python 3.12, crea venv, instala deps, descarga Chromium para Playwright).

## Arquitectura

```
historial_manager.py (PyWebView)
    └── HistorialApi (clase expuesta a JS via window.pywebview.api)
            ├── cargar_estado()          → lee clientes de DB
            ├── crear_cliente()          → CRUD clientes
            ├── actualizar_cliente()
            ├── eliminar_cliente()
            ├── crear_subcuenta()        → CRUD subcuentas
            ├── eliminar_subcuenta()
            ├── iniciar_descarga()       → lanza thread en background
            │       └── historial_facturas_ute.run_historial()
            ├── obtener_historial()      → registros de descarga
            └── abrir_carpeta()         → abre carpeta de output en Explorer
```

Comunicación JS → Python: `window.pywebview.api.metodo(args)`  
Comunicación Python → JS (logs en tiempo real): `window.recibirLog(mensaje)`

## Base de datos (SQLite)

**Tablas:**
- `clientes` — id, nombre, usuario, password (base64), carpeta
- `subcuentas` — id, cliente_id, numero
- `descargas` — id, cliente_id, usuario, directorio, fecha_inicio, fecha_fin, estado, notas

Las contraseñas se guardan en base64 (ofuscación, no encriptación real).

## Flujo de descarga

1. Usuario selecciona cliente + subcuentas en la UI
2. JS llama `iniciar_descarga(cliente_id)`
3. Python levanta un thread background y ejecuta `run_historial()`
4. Playwright abre Chromium headless, loguea en portal UTE
5. Navega iframes del portal, recolecta links de facturas de cada subcuenta
6. Por cada subcuenta: descarga PDFs (hasta 24 facturas), parsea datos
7. Datos extraídos de PDF: nombre cliente, dirección, importes, lecturas de medidor, tensión, zona
8. Al finalizar: genera Excel con formato y estilos (openpyxl)
9. Registra descarga en tabla `descargas`
10. Logs de cada paso se envían al frontend via `window.recibirLog()`

## Output generado

Directorio de salida: `C:/Users/{user}/Documents/SEG/Tarea UTE/{cliente_name}/{subcuenta}/`  
Archivos: `historial_facturas_{subcuenta}.xlsx` + carpeta `historial/` con los PDFs descargados

## Detalles técnicos importantes

**PDF Parsing (`parsing_facturas.py`):**
- Extrae texto por coordenadas de página (bloque cliente: 28–78% eje X, 0–24% eje Y)
- Convierte formato numérico uruguayo: `1.234,56` → `1234.56`
- Detecta nivel de tensión, zona, tramos de consumo desde texto de factura
- Maneja reconstrucción de texto con problemas de espaciado (estilo OCR)

**Playwright:**
- Timeout de 12 minutos por cuenta
- Manejo de iframes dinámicos (portal UTE usa JS pesado)
- Función `_goto_estable()` para navegación con reintentos
- Detección y re-engagement de frames desconectados

**UI (`index.html`):**
- Sidebar oscuro con lista de clientes
- Área principal: datos del cliente, subcuentas (checkbox individual o bulk), botón descarga
- Log viewer en tiempo real
- Modal para crear/editar clientes
- Toggle de visibilidad de contraseña
- Input bulk de subcuentas: acepta números separados por coma, punto y coma o salto de línea

## No hay tests

Sin suite de tests. Hay un bloque `if __name__ == "__main__"` en `historial_facturas_ute.py` para uso CLI manual.
