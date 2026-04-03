# ============================================================
# SISTEMA ANALIZADOR ACADÉMICO UCSUR
# APLICACIÓN PRINCIPAL — VENTANA ÚNICA
# ============================================================

import customtkinter as ctk
from datetime import datetime
from tkinter import simpledialog, messagebox

# ============================================================
# IMPORTACIÓN DE VISTAS (NO SE MODIFICA)
# ============================================================

from vista_query import FrameQuery
from vista_global import FrameGlobal
from vista_carrera import FrameCarrera
from vista_secciones import FrameSecciones
from vista_registro import FrameRegistroNotas
from vista_informe_final_curso import FrameInformeFinalCurso   

# ✅ NUEVO: Riesgo Académico
from vista_riesgo_academico import FrameRiesgoAcademico      
from vista_analisis_riesgo import FrameAnalisisRiesgo



# ============================================================
# CONFIGURACIÓN GENERAL (ORIGINAL)
# ============================================================

APP_TITLE = "Sistema Analizador Académico – UCSUR"

APP_SIZE = (1050, 720)
MIN_SIZE = (980, 680)

COL_UCSUR_AZUL = "#003B70"
COL_BG = "#F5F5F5"


# ============================================================
# 🔐 CONTROL OCULTO POR FECHA (NUEVO, DISCRETO)
# ============================================================

def verificar_fecha_y_clave():
    """
    Control institucional de acceso:
    - Antes del 01/05/2026: acceso normal
    - Desde el 01/05/2026: solicita contraseña
    """
    fecha_limite = datetime(2026, 5, 1)
    hoy = datetime.now()

    if hoy >= fecha_limite:
        root = ctk.CTk()
        root.withdraw()

        clave = simpledialog.askstring(
            "Acceso restringido",
            "Ingrese la contraseña para continuar:",
            show="*"
        )

        root.destroy()

        if clave != "CiEnCiAs2025":
            messagebox.showerror(
                "Acceso denegado",
                "Contraseña incorrecta.\n\nEl sistema se cerrará."
            )
            return False

    return True


# ============================================================
# APLICACIÓN PRINCIPAL (ORIGINAL + EXTENSIÓN)
# ============================================================

class AppUCSUR(ctk.CTk):

    def __init__(self):
        super().__init__()

        # -------------------------------
        # Configuración básica
        # -------------------------------
        self.title(APP_TITLE)
        self._centrar_ventana(*APP_SIZE)
        self.minsize(*MIN_SIZE)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # -------------------------------
        # Layout base
        # -------------------------------
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self._build_sidebar()
        self._build_container()
        self._build_frames()

        # Vista inicial
        self.mostrar("query")

    # ========================================================
    # CENTRAR VENTANA (ORIGINAL)
    # ========================================================

    def _centrar_ventana(self, width, height):
        self.update_idletasks()

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        x = (screen_w - width) // 2
        y = (screen_h - height) // 2

        self.geometry(f"{width}x{height}+{x}+{y}")

    # ========================================================
    # SIDEBAR (ORIGINAL + NUEVO BOTÓN)
    # ========================================================

    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(
            self,
            width=240,
            fg_color="white",
            corner_radius=0
        )
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self.sidebar.grid_propagate(False)

        # Título
        ctk.CTkLabel(
            self.sidebar,
            text="UCSUR",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).pack(pady=(22, 4))

        ctk.CTkLabel(
            self.sidebar,
            text="Sistema de\nAnálisis Académico",
            justify="center",
            font=ctk.CTkFont(size=13),
            text_color="#555"
        ).pack(pady=(0, 20))

        # Botones de navegación
        self._btn("📥 Extraer Query", "query")
        self._btn("📊 Análisis Global", "global")
        self._btn("🎓 Análisis por Carrera", "carrera")
        self._btn("📚 Análisis por Curso", "secciones")

        # 🔹 NUEVO BOTÓN
        self._btn("📌 INFORME FINAL DE CURSO", "informe_final_curso")

        # ✅ NUEVO BOTÓN: Riesgo Académico
        self._btn("⚠️ Riesgo Académico", "riesgo_academico")

        self._btn("📝 Registro de Notas", "registro")

        ctk.CTkFrame(
            self.sidebar,
            height=2,
            fg_color="#DDDDDD"
        ).pack(fill="x", padx=12, pady=18)

        self._btn("ℹ️ Acerca de", "acerca")
        self._btn("❌ Salir", "salir", danger=True)

    def _btn(self, texto, vista, danger=False):
        fg = "#AA0000" if danger else COL_UCSUR_AZUL
        hover = "#CC0000" if danger else "#004B8D"

        ctk.CTkButton(
            self.sidebar,
            text=texto,
            height=44,
            fg_color=fg,
            hover_color=hover,
            corner_radius=10,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=lambda: self._accion(vista)
        ).pack(fill="x", padx=15, pady=6)

    def _accion(self, vista):
        if vista == "salir":
            self.destroy()
        elif vista == "acerca":
            self.mostrar_acerca_de()
        else:
            self.mostrar(vista)

    # ========================================================
    # CONTENEDOR CENTRAL (ORIGINAL)
    # ========================================================

    def _build_container(self):
        self.container = ctk.CTkFrame(
            self,
            fg_color=COL_BG,
            corner_radius=0
        )
        self.container.grid(row=0, column=1, sticky="nsew")
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

    # ========================================================
    # FRAMES (ORIGINAL + NUEVO)
    # ========================================================

    def _build_frames(self):
        self.frames = {
            "query": FrameQuery(self.container, self),
            "global": FrameGlobal(self.container, self),
            "carrera": FrameCarrera(self.container, self),
            "secciones": FrameSecciones(self.container, self),
            "informe_final_curso": FrameInformeFinalCurso(self.container, self),  

            # ✅ NUEVO: Riesgo Académico
            "riesgo_academico": FrameRiesgoAcademico(self.container, self),       
            "analisis_riesgo": FrameAnalisisRiesgo(self.container, self),

            "registro": FrameRegistroNotas(self.container, self),
        }

        for frame in self.frames.values():
            frame.grid(row=0, column=0, sticky="nsew")

    # ========================================================
    # NAVEGACIÓN (ORIGINAL)
    # ========================================================

    def mostrar(self, nombre):
        frame = self.frames.get(nombre)
        if not frame:
            return

        if hasattr(frame, "on_show") and callable(getattr(frame, "on_show")):
            try:
                frame.on_show()
            except Exception as e:
                print(f"⚠️ Error al refrescar la vista '{nombre}': {e}")

        frame.tkraise()

    def mostrar_vista(self, nombre):
        self.mostrar(nombre)

    # ========================================================
    # ℹ️ ACERCA DE (ORIGINAL)
    # ========================================================

    def mostrar_acerca_de(self):
        win = ctk.CTkToplevel(self)
        win.title("Acerca de")
        win.geometry("520x300")
        win.resizable(False, False)
        win.grab_set()

        texto = (
            "Sistema Analizador Académico – UCSUR\n\n"
            "Desarrollado en Cursos Básicos – Ciencias\n"
            "Equipo de Matemática\n\n"
            "Desarrollado por Richard Cubas.\n"
            "Windows puede mostrar advertencias por tratarse de un ejecutable no firmado.\n"
            "El archivo es seguro y ha sido validado por el equipo académico."
        )

        ctk.CTkLabel(
            win,
            text=texto,
            justify="center",
            font=ctk.CTkFont(size=13)
        ).pack(expand=True, padx=20, pady=20)

        ctk.CTkButton(
            win,
            text="Cerrar",
            command=win.destroy
        ).pack(pady=10)


# ============================================================
# MAIN (ORIGINAL + CONTROL)
# ============================================================

if __name__ == "__main__":

    if not verificar_fecha_y_clave():
        exit(0)

    app = AppUCSUR()
    app.mainloop()
