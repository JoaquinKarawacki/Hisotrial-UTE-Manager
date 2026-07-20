"""
historial_manager.py — Frontend para descargas de historial de facturas (PyWebView).

Estructura:
    Tarea UTE/
    ├── historial_facturas_ute.py
    ├── historial_manager.py
    ├── database.py
    ├── requirements.txt
    ├── setup.bat
    ├── Iniciar.bat
    └── ui/
        └── index.html
"""

import os
import sys
import threading
import time
import json

import webview
import historial_facturas_ute
import database as db


BASE_OUTPUT = os.path.join(os.path.expanduser("~"), "Documents", "SEG", "Tarea UTE")


class HistorialApi:
    """Métodos expuestos al JavaScript via window.pywebview.api.<metodo>()"""

    def __init__(self):
        self._descarga_en_curso = False
        self._logs_buffer = []
        self._descarga_id = None

    # ── GESTIÓN DE CLIENTES ───────────────────────────────────────────────────

    def cargar_estado(self) -> dict:
        """Retorna todos los clientes y sus estados recientes."""
        try:
            clientes_data = db.get_estado_completo()
            clientes = []
            for cliente_dict in clientes_data:
                clientes.append({
                    "id": cliente_dict["id"],
                    "nombre": cliente_dict["nombre"],
                    "usuario": cliente_dict["usuario"],
                    "password_decoded": cliente_dict["password_decoded"],
                    "subcuentas": cliente_dict.get("subcuentas", [])
                })
            return {"ok": True, "clientes": clientes}
        except Exception as e:
            return {"ok": False, "error": str(e), "clientes": []}

    def crear_cliente(self, nombre: str, usuario: str, password: str) -> dict:
        """Crea un nuevo cliente."""
        try:
            if not nombre or not usuario or not password:
                return {"ok": False, "error": "Todos los campos son requeridos"}
            
            cliente_id = db.crear_cliente(nombre, usuario, password)
            return {"ok": True, "id": cliente_id}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def actualizar_cliente(self, cliente_id: int, nombre: str, usuario: str, password: str = "") -> dict:
        """Actualiza datos de un cliente. Si password viene vacío, se conserva la actual."""
        try:
            if not nombre or not usuario:
                return {"ok": False, "error": "Nombre y usuario son requeridos"}

            db.actualizar_cliente(cliente_id, nombre, usuario, password)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def eliminar_cliente(self, cliente_id: int) -> dict:
        """Elimina un cliente."""
        try:
            db.eliminar_cliente(cliente_id)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def crear_subcuenta(self, cliente_id: int, numero: str) -> dict:
        """Crea una subcuenta para un cliente."""
        try:
            subcuenta_id = db.crear_subcuenta(cliente_id, numero)
            return {"ok": True, "id": subcuenta_id}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def eliminar_subcuenta(self, subcuenta_id: int) -> dict:
        """Elimina una subcuenta por ID."""
        try:
            db.eliminar_subcuenta(subcuenta_id)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── DESCARGAS ─────────────────────────────────────────────────────────────

    def iniciar_descarga(self, cliente_id: int, subcuentas_seleccionadas: list[str] = None) -> dict:
        """
        Lanza la descarga de historial para un cliente específico.
        Obtiene credenciales de la BD e inicia descarga en hilo separado.
        """
        if self._descarga_en_curso:
            return {"ok": False, "error": "Ya hay una descarga en curso."}

        try:
            # Obtener cliente y credenciales
            cliente = db.get_cliente(cliente_id)
            if not cliente:
                return {"ok": False, "error": "Cliente no encontrado"}
            
            usuario = cliente["usuario"]
            password = cliente["password_decoded"]
            output_dir = os.path.join(BASE_OUTPUT, cliente["nombre"])
            cuentas = [str(c).strip() for c in (subcuentas_seleccionadas or []) if str(c).strip()]

            self._descarga_en_curso = True
            self._descarga_id = db.registrar_descarga(cliente_id, usuario, output_dir)
            self._logs_buffer = []

            self._log(f"🚀 Iniciando descarga de historial...")
            self._log(f"   Cliente : {cliente['nombre']}")
            self._log(f"   Usuario : {usuario}")
            self._log(f"   Subcuentas seleccionadas: {len(cuentas)}")
            self._log(f"   Salida  : {output_dir}")

            def tarea():
                try:
                    historial_facturas_ute.run_historial(
                        ute_user=usuario,
                        ute_pass=password,
                        cuentas_seleccionadas=cuentas,
                        output_dir=output_dir,
                        log_fn=self._log,
                    )
                    db.finalizar_descarga(self._descarga_id, "exitosa")
                    self._log("✓ Descarga completada y registrada en historial")
                except Exception as e:
                    self._log(f"❌ Error inesperado: {e}")
                    db.finalizar_descarga(self._descarga_id, "error", str(e))
                finally:
                    self._descarga_en_curso = False
                    self._log("__DESCARGA_FINALIZADA__")

            threading.Thread(target=tarea, daemon=True).start()
            return {"ok": True}

        except Exception as e:
            self._descarga_en_curso = False
            return {"ok": False, "error": str(e)}

    def descarga_en_curso(self) -> bool:
        """Devuelve True si hay una descarga activa."""
        return self._descarga_en_curso

    def obtener_historial(self, cliente_id: int = None) -> list:
        """Devuelve el historial de descargas (opcional: filtrado por cliente)."""
        try:
            historial = db.obtener_historial(cliente_id=cliente_id, limite=20)
            return [dict(h) for h in historial]
        except Exception as e:
            return []

    def abrir_carpeta_cliente(self, cliente_id: int) -> dict:
        """Abre la carpeta de salida de un cliente en el explorador."""
        try:
            cliente = db.get_cliente(cliente_id)
            if not cliente:
                return {"ok": False, "error": "Cliente no encontrado"}
            ruta = os.path.join(BASE_OUTPUT, cliente["nombre"])
            return self.abrir_carpeta(ruta)
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def abrir_carpeta(self, ruta: str) -> dict:
        """Abre la carpeta en el explorador."""
        try:
            os.makedirs(ruta, exist_ok=True)
            if sys.platform == "win32":
                os.startfile(ruta)
            elif sys.platform == "darwin":
                os.system(f'open "{ruta}"')
            else:
                os.system(f'xdg-open "{ruta}"')
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _log(self, mensaje: str):
        """Envía mensaje de log al frontend."""
        def _enviar_js(msg: str):
            try:
                msg_safe = (msg
                            .replace("\\", "\\\\")
                            .replace("'", "\\'")
                            .replace("\n", " "))
                if "window" in globals() and window is not None:
                    window.evaluate_js(f"window.recibirLog('{msg_safe}')")
            except Exception:
                pass

        threading.Thread(target=_enviar_js, args=(mensaje,), daemon=True).start()
        print(mensaje)


# ── Inicialización ────────────────────────────────────────────────────────────

UI_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ui", "index.html")

api = HistorialApi()

window = webview.create_window(
    title="SEG Ingeniería — Historial Facturas UTE",
    url=f"file://{UI_PATH}",
    js_api=api,
    width=1200,
    height=760,
    min_size=(900, 600),
    resizable=True,
    text_select=False,
)

if __name__ == "__main__":
    webview.start(debug=False)
