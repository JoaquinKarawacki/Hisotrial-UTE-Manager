# Historial Facturas UTE — Guía de Instalación

Este proyecto descarga automáticamente el historial de facturas de UTE para múltiples clientes empresariales y genera reportes Excel consolidados por subcuenta.

## Estructura del Proyecto

```
Hisotrial_manager/
├── src/
│   ├── historial_manager.py       # Entry point: PyWebView + bridge Python↔JS
│   ├── historial_facturas_ute.py  # Motor de scraping y parsing de PDFs
│   └── database.py                # Wrapper SQLite (clientes, subcuentas, descargas)
├── ui/
│   └── index.html                 # Frontend HTML/CSS/JS
├── img/
│   └── logo.png                   # Logo mostrado en la sidebar (no versionado)
├── data/
│   └── seg_historial.db           # Base de datos SQLite (no versionada)
├── requirements.txt                # Dependencias Python
├── setup.bat                       # Instalador (ejecuta una única vez)
├── Iniciar.bat                     # Inicia la aplicación
└── README.md                       # Este archivo
```

## Instalación (Primera Vez)

### 1. Descargar el Proyecto

Descarga esta carpeta completa en tu PC.

### 2. Ejecutar setup.bat

Haz doble clic en **`setup.bat`**. Esto:
- Verifica si tienes Python instalado (si no, lo descarga automáticamente)
- Crea un entorno virtual `.venv`
- Instala todas las dependencias de `requirements.txt`
- Descarga el navegador Chromium para Playwright

⏱️ Esto puede tomar **5-10 minutos** la primera vez.

```bash
setup.bat
```

### 3. ¡Listo!

Una vez completado, verás el mensaje:
```
Listo! Ahora podes iniciar la app con:
Iniciar.bat
```

## Uso

### Iniciar la Aplicación

Haz doble clic en **`Iniciar.bat`** para abrir la ventana de la aplicación.

```bash
Iniciar.bat
```

Esto abrirá una ventana con:
1. **Sidebar de clientes**: lista de clientes cargados, con opción de crear uno nuevo
2. **Subcuentas**: alta/baja de subcuentas por cliente (individual o en bloque)
3. **Botón de descarga**: inicia la descarga del historial para las subcuentas seleccionadas (o todas, si no se selecciona ninguna)
4. **Log en tiempo real**: progreso de la descarga
5. **Abrir carpeta**: acceso directo a la carpeta de resultados del cliente

### Credenciales

Las credenciales de UTE (usuario y contraseña) se cargan una vez por cliente, desde el modal "Nuevo Cliente" / "Editar Cliente", y quedan guardadas en la base de datos local. Al editar un cliente, el campo de contraseña puede dejarse en blanco para conservar la contraseña actual.

### Resultados

Después de la descarga, encontrarás en `Documents/SEG/Tarea UTE/{cliente}/`:

```
{cliente}/
├── historial_0123456789/       # PDFs descargados de la subcuenta 1
├── 0123456789.xlsx             # Excel consolidado de la subcuenta 1
├── historial_0234567890/       # PDFs descargados de la subcuenta 2
├── 0234567890.xlsx
└── ...
```

Cada Excel contiene:
- **Historial**: datos parseados de todas las facturas (24 máximo por subcuenta)
- Columnas: número de cuenta, factura, período, consumo, lecturas de medidor, etc.

## Historial de Descargas

La aplicación guarda un registro de todas las descargas en la tabla `descargas` de `data/seg_historial.db`:
- Fecha y hora de inicio/fin
- Usuario utilizado
- Directorio de salida
- Estado (exitosa/error)

## Troubleshooting

### Python no se instala automáticamente

Si el setup falla al descargar Python:
1. Descarga Python 3.12+ desde https://www.python.org/downloads/
2. Instálalo y marca la opción "Add Python to PATH"
3. Ejecuta `setup.bat` nuevamente

### Chromium no descarga

Si falla la descarga de Chromium:
```bash
.venv\Scripts\python -m playwright install chromium
```

### Error de credenciales

- Verifica que tu usuario y contraseña sean correctos
- UTE puede requerir autenticación de dos factores (revisar el navegador que se abre)

### La descarga se queda colgada

El script tiene un corte automático a 12 minutos por subcuenta. Si se queda pegado más tiempo, cierra la ventana y reinicia.

## Archivos Importantes

| Archivo | Descripción |
|---------|-----------|
| `src/historial_facturas_ute.py` | Lógica de scraping y parseo de PDFs |
| `src/historial_manager.py` | Entry point y bridge Python↔JS (PyWebView) |
| `src/database.py` | Base de datos SQLite (clientes, subcuentas, historial) |
| `ui/index.html` | Interfaz visual (HTML/CSS/JavaScript) |
| `requirements.txt` | Lista de librerías Python necesarias |

## Dependencias Principales

- **Playwright**: Automatización del navegador (descarga de facturas)
- **PyWebView**: Framework para interfaz gráfica
- **Pandas**: Procesamiento de datos y generación de Excel
- **pdfplumber**: Extracción de texto de PDFs
- **openpyxl**: Escritura de archivos Excel

## Notas de seguridad

- Las contraseñas de UTE se guardan en `data/seg_historial.db` **en base64** (ofuscación, no encriptación real). El archivo `.db` está excluido de git vía `.gitignore` — nunca fuerces su versionado (`git add -f`).
- No hay tests automatizados. Hay un bloque `if __name__ == "__main__"` en `historial_facturas_ute.py` para uso CLI manual.

## Contacto / Soporte

Si tienes problemas, envía un mensaje con:
1. El último log de la aplicación (captura pantalla o copia-pega)
2. El error exacto que ves
3. Tu sistema operativo y versión de Python

---

**SEG Ingeniería** © 2026
