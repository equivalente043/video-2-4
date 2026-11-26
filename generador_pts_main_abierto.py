# -*- coding: utf-8 -*-
"""
GENERADOR PTS FWD & ETF - MAIN ABIERTO   v4 - 2026-11-26
=========================================================
El script NUNCA abre, cierra ni reinicia Excel a la fuerza: trabaja sobre tu Excel, con el main
abierto y Bloomberg funcionando. Para cada día hábil del rango:

    1. Escribe la fecha en MAIN!D4 del main abierto y espera 30 segundos a que cargue Bloomberg.
    2. Revisa: si hay #¿NOMBRE? (Bloomberg no cargado) no guarda nada; activa Bloomberg y recalcula.
    3. Guarda una COPIA del main (SaveCopyAs, el main no se toca), la abre sin recalcular y sin macros,
       y hace los mismos pasos de genera_report sobre la copia: guardar como aaaammdd.xlsx,
       pegar valores en todas las hojas y guardar.
    4. Reabre el reporte guardado y confirma que no tenga #¿NOMBRE?.

Al terminar deja MAIN!D4 como estaba. Si el main es tuyo (abierto para editar), le apaga el
Autoguardado mientras corre para que los cambios de fecha no se graben en tu archivo.

Antes de correrlo: abre Excel y el main como siempre y confirma que Bloomberg cargó datos.
Requisitos: pip install pywin32 openpyxl
"""

import os
import csv
import glob
import time
import ctypes
import tempfile
import threading
import subprocess
import datetime as dt
from collections import Counter, deque

try:
    import pywintypes
    COM_ERROR = pywintypes.com_error
except ImportError:  # solo para poder probar el script fuera de Windows
    pywintypes = None

    class COM_ERROR(Exception):
        hresult = None


# =============================================================================
# CONFIGURACIÓN
# =============================================================================
CARPETA_MAIN = r"C:\Users\p669213\OneDrive - Principal Financial Group\Maria"
NOMBRE_MAIN = ""      # vacío = busca el único .xlsm con "fwd" y "etf" en el nombre

# Carpeta local sincronizada de los reportes: CARPETA_SALIDA\<año>\<aaaammdd>.xlsx
CARPETA_SALIDA = (
    r"C:\Users\p669213\Principal Financial Group"
    r"\Gerencia Finanzas AFP - Control y Cumplimiento Inversiones"
    r"\002 Cumplimiento\Conflictos de Interés\Control de precios de transacciones"
    r"\Pts FWD & ETF - Bloomberg"
)

FECHA_DESDE_SUGERIDA = "20260904"
FECHA_HASTA_SUGERIDA = "20261117"

SEG_ESPERA_BLOOMBERG = 30   # espera después de escribir la fecha
MAX_REQUESTING_OK = 350     # celdas "Requesting Data" aceptables al guardar (lo normal es ~276-290)
SEG_ESPERA_EXTRA = 60       # si hay más que eso, espera hasta este tiempo adicional
SEG_ESPERA_ADDIN = 90       # si el script tiene que abrir Excel, cuánto esperar a que cargue Bloomberg

REVISAR_EXISTENTES = True   # revisa los reportes que ya existen; los malos se regeneran

EXCLUIR_FERIADOS_CHILE = True
FERIADOS_CHILE = {
    "20260101", "20260403", "20260404", "20260501", "20260521", "20260621",
    "20260629", "20260716", "20260815", "20260918", "20260919", "20261012",
    "20261031", "20261101", "20261208", "20261225",
}

MANTENER_PANTALLA_ENCENDIDA = True
MAX_INTENTOS_POR_FECHA = 5
MAX_HORAS = 12
SEG_EXCEL_OCUPADO = 600     # Excel ocupado más que esto: se detiene y te avisa (no se cierra Excel)
SEG_REINTENTO_COM = 0.5
# =============================================================================

VERSION = "main abierto v4 (2026-11-26)"
XL_CALC_AUTOMATIC = -4105
XL_CALC_MANUAL = -4135
XL_PASTE_VALUES = -4163
XL_OPENXML_WORKBOOK = 51
XL_ERR_NAME = -2146826259          # #¿NOMBRE? / #NAME?
MACROS_DESHABILITADAS = 3          # msoAutomationSecurityForceDisable
EXCEL_OCUPADO = {-2147418111, -2147417846, -2146777998}
EXCEL_CAIDO = {-2147417848, -2147023174, -2147023170, -2147023169, -2147221251, -2147417851}


class ErrorFatal(Exception):
    """Necesita a una persona: se detiene."""


class ErrorExcel(Exception):
    """Excel se cerró o no responde."""


class ErrorBloombergNoCargado(Exception):
    """Excel no tiene cargado el complemento de Bloomberg (#¿NOMBRE?): no se guarda nada."""


class ErrorBloombergConexion(Exception):
    """Bloomberg sin conexión o con límite de datos: se espera y se reintenta."""


def log(msg=""):
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def esperar(seg, texto):
    fin = time.time() + seg
    while True:
        falta = fin - time.time()
        if falta <= 0:
            break
        print(f"      {texto} {int(falta):3d}s   ", end="\r", flush=True)
        time.sleep(min(1, falta))
    print(" " * 60, end="\r")


def _hresults(e):
    codigos = {getattr(e, "hresult", None)}
    info = getattr(e, "excepinfo", None)
    if info and len(info) > 5:
        codigos.add(info[5])
    return codigos


def com(fn, seg_max=None):
    """Llamada a Excel: si está ocupado espera y reintenta; si se cayó o se colgó, ErrorExcel."""
    t0 = time.time()
    seg_max = SEG_EXCEL_OCUPADO if seg_max is None else seg_max
    while True:
        try:
            return fn()
        except COM_ERROR as e:
            codigos = _hresults(e)
            if codigos & EXCEL_CAIDO:
                raise ErrorExcel("Excel se cerró o dejó de responder") from e
            if not codigos & EXCEL_OCUPADO:
                raise
            if time.time() - t0 > seg_max:
                raise ErrorExcel(f"Excel siguió ocupado más de {seg_max:.0f}s (colgado)") from e
            time.sleep(SEG_REINTENTO_COM)


def ruta_reporte(fecha):
    return os.path.join(CARPETA_SALIDA, f"{fecha:%Y}", f"{fecha:%Y%m%d}.xlsx")


class PantallaEncendida:
    """PC despierto y pantalla prendida. Si nadie toca el PC por ~50 s, mueve el mouse 1 pixel
    ida y vuelta para que no se bloquee. No hace clics ni teclea."""

    ES_CONTINUOUS, ES_SYSTEM_REQUIRED, ES_DISPLAY_REQUIRED = 0x80000000, 0x00000001, 0x00000002

    def __init__(self):
        self._parar = threading.Event()

    def iniciar(self):
        if not MANTENER_PANTALLA_ENCENDIDA:
            return
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(
                self.ES_CONTINUOUS | self.ES_SYSTEM_REQUIRED | self.ES_DISPLAY_REQUIRED)
        except Exception:
            return
        threading.Thread(target=self._mantener, daemon=True).start()
        log("Pantalla y PC se mantienen encendidos mientras corre el script.")

    def _mantener(self):
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        while not self._parar.wait(20):
            try:
                ctypes.windll.kernel32.SetThreadExecutionState(
                    self.ES_CONTINUOUS | self.ES_SYSTEM_REQUIRED | self.ES_DISPLAY_REQUIRED)
                ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii))
                if ((ctypes.windll.kernel32.GetTickCount() - lii.dwTime) & 0xFFFFFFFF) > 50_000:
                    ctypes.windll.user32.mouse_event(0x0001, 1, 0, 0, 0)
                    ctypes.windll.user32.mouse_event(0x0001, -1, 0, 0, 0)
            except Exception:
                pass

    def detener(self):
        self._parar.set()
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(self.ES_CONTINUOUS)
        except Exception:
            pass


def buscar_main():
    if NOMBRE_MAIN:
        ruta = os.path.join(CARPETA_MAIN, NOMBRE_MAIN)
        if not os.path.isfile(ruta):
            raise ErrorFatal(f"No existe el main: {ruta}")
        return ruta
    candidatos = [p for p in glob.glob(os.path.join(CARPETA_MAIN, "*.xlsm"))
                  if not os.path.basename(p).startswith("~$")
                  and "fwd" in os.path.basename(p).lower() and "etf" in os.path.basename(p).lower()]
    if len(candidatos) != 1:
        raise ErrorFatal("No pude identificar un único main. Escribe su nombre en NOMBRE_MAIN. "
                         f"Encontrados: {[os.path.basename(c) for c in candidatos]}")
    return candidatos[0]


def pedir_fecha(texto, sugerida):
    while True:
        s = input(f"{texto} (aaaammdd) [Enter = {sugerida}]: ").strip() or sugerida
        try:
            return dt.datetime.strptime(s, "%Y%m%d").date()
        except ValueError:
            print("   Formato inválido, ejemplo: 20260904")


def dias_habiles(desde, hasta):
    feriados = {dt.datetime.strptime(f, "%Y%m%d").date() for f in FERIADOS_CHILE}
    habiles, saltados = [], []
    d = desde
    while d <= hasta:
        if d.weekday() < 5:
            (saltados if EXCLUIR_FERIADOS_CHILE and d in feriados else habiles).append(d)
        d += dt.timedelta(days=1)
    return habiles, saltados


# -----------------------------------------------------------------------------
# Revisión de un reporte ya guardado (sin Excel)
# -----------------------------------------------------------------------------
def revisar_reporte(ruta, fecha):
    """None si el reporte está bien; si no, el motivo. Si no se puede abrir (en uso), None."""
    import warnings
    import openpyxl

    warnings.filterwarnings("ignore")
    try:
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    except PermissionError:
        return None
    except Exception as e:
        return f"archivo dañado ({type(e).__name__})"
    try:
        if "MAIN" not in wb.sheetnames:
            return "no tiene la hoja MAIN"
        d4 = next(wb["MAIN"].iter_rows(min_row=4, max_row=4, min_col=4, max_col=4, values_only=True))[0]
        if not (isinstance(d4, dt.datetime) and d4.date() == fecha):
            return f"MAIN!D4 dice {d4}"
        nombre = requesting = 0
        for ws in wb.worksheets:
            for fila in ws.iter_rows(values_only=True):
                for v in fila:
                    if isinstance(v, str):
                        if v.startswith("#NAME"):
                            nombre += 1
                        elif "Requesting Data" in v:
                            requesting += 1
        if nombre:
            return f"{nombre} celdas #¿NOMBRE? (Bloomberg no estaba cargado)"
        if requesting > MAX_REQUESTING_OK:
            return f"{requesting} celdas 'Requesting Data' (Bloomberg no alcanzó a cargar)"
        return None
    finally:
        wb.close()


# -----------------------------------------------------------------------------
# Tu Excel y el main
# -----------------------------------------------------------------------------
def conectar_excel():
    """Se conecta a tu Excel abierto. Si no hay Excel abierto, lo abre normalmente (con Bloomberg)."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Excel.Application"), False
    except Exception:
        pass
    log("   No hay Excel abierto: lo abro normalmente (con Bloomberg) y espero...")
    subprocess.Popen('start "" excel.exe', shell=True)
    for _ in range(60):
        time.sleep(2)
        try:
            return win32com.client.GetActiveObject("Excel.Application"), True
        except Exception:
            pass
    raise ErrorFatal("No pude conectarme a Excel. Ábrelo a mano con el main y vuelve a ejecutar.")


def complementos_bloomberg(xl):
    """Lista (descripción, conectado) de los complementos COM de Bloomberg de este Excel."""
    salida = []
    try:
        for ca in com(lambda: list(xl.COMAddIns), seg_max=30):
            try:
                texto = f"{ca.Description} {ca.ProgId}"
                if "bloomberg" in texto.lower():
                    salida.append((ca.Description, bool(ca.Connect)))
            except Exception:
                pass
    except Exception:
        pass
    return salida


def activar_bloomberg(xl):
    """Activa los complementos de Bloomberg que estén apagados en este Excel."""
    try:
        for ca in com(lambda: list(xl.COMAddIns), seg_max=30):
            try:
                if "bloomberg" in f"{ca.Description} {ca.ProgId}".lower() and not ca.Connect:
                    ca.Connect = True
                    log(f"   Activé el complemento de Excel: {ca.Description}")
            except Exception:
                pass
    except Exception:
        pass
    try:
        for ai in com(lambda: list(xl.AddIns), seg_max=30):
            try:
                texto = f"{ai.Name} {ai.FullName}".lower()
                if ("bloomberg" in texto or "bofaddin" in texto) and not ai.Installed:
                    ai.Installed = True
                    log(f"   Activé el complemento de Excel: {ai.Name}")
            except Exception:
                pass
    except Exception:
        pass


def esperar_bloomberg_en_excel(xl):
    """Después de abrir Excel: espera a que el complemento de Bloomberg esté conectado
    (si está apagado lo activa) y deja un margen para que registre sus fórmulas."""
    t0 = time.time()
    while time.time() - t0 < SEG_ESPERA_ADDIN:
        estado = complementos_bloomberg(xl)
        if estado and all(c for _, c in estado):
            break
        if estado:
            activar_bloomberg(xl)
        print(f"      esperando que cargue Bloomberg en Excel {int(time.time() - t0):3d}s   ", end="\r", flush=True)
        time.sleep(3)
    print(" " * 60, end="\r")
    esperar(15, "margen para que Bloomberg registre sus fórmulas")


def obtener_main(xl, ruta_main):
    """Devuelve (wb, lo_abrio_el_script). Usa el main si ya está abierto; si no, lo abre en solo lectura."""
    nombre = os.path.basename(ruta_main).lower()
    for w in com(lambda: list(xl.Workbooks)):
        if com(lambda: w.Name).lower() == nombre:
            return w, False
    log("   El main no estaba abierto: lo abro en solo lectura (queda abierto todo el proceso).")
    com(lambda: setattr(xl, "DisplayAlerts", False))
    try:
        wb = com(lambda: xl.Workbooks.Open(Filename=ruta_main, UpdateLinks=0, ReadOnly=True))
    finally:
        com(lambda: setattr(xl, "DisplayAlerts", True))
    return wb, True


def estado_libro(wb):
    """Cuenta en todo el libro: celdas 'Requesting Data', celdas #¿NOMBRE? y mensajes de Bloomberg
    de conexión o límite."""
    requesting = nombre = 0
    problema = None
    for ws in com(lambda: list(wb.Worksheets)):
        vals = com(lambda: ws.UsedRange.Value2)
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        for fila in vals:
            for v in fila:
                if isinstance(v, str):
                    if v.startswith("#N/A"):
                        up = v.upper()
                        if "REQUESTING DATA" in up:
                            requesting += 1
                        elif problema is None and any(t in up for t in ("LIMIT", "CAPACITY", "CONNECTION", "LOGIN", "LOGGED")):
                            problema = v
                elif type(v) is int and v == XL_ERR_NAME:
                    nombre += 1
    return requesting, nombre, problema


# -----------------------------------------------------------------------------
# Generación desde una copia del main
# -----------------------------------------------------------------------------
def generar_desde_copia(xl, main, ruta_salida):
    """Guarda una copia del main tal como está (con la fecha y los datos ya cargados), la abre sin
    recalcular y sin ejecutar macros, y hace sobre ella los pasos de genera_report. El main no se toca."""
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    copia = os.path.join(tempfile.gettempdir(), f"pts_copia_{int(time.time() * 1000)}.xlsm")
    com(lambda: main.SaveCopyAs(copia))

    calc_antes = com(lambda: xl.Calculation)
    eventos_antes = com(lambda: xl.EnableEvents)
    try:
        seguridad_antes = xl.AutomationSecurity
    except Exception:
        seguridad_antes = None
    wb = None
    try:
        com(lambda: setattr(xl, "Calculation", XL_CALC_MANUAL))     # la copia no se recalcula
        com(lambda: setattr(xl, "EnableEvents", False))              # no corre Workbook_Open
        try:
            xl.AutomationSecurity = MACROS_DESHABILITADAS
        except Exception:
            pass
        com(lambda: setattr(xl, "ScreenUpdating", False))
        com(lambda: setattr(xl, "DisplayAlerts", False))
        com(lambda: setattr(xl, "CutCopyMode", False))
        wb = com(lambda: xl.Workbooks.Open(Filename=copia, UpdateLinks=0, ReadOnly=False))

        # mismos pasos que genera_report
        hoja_activa = com(lambda: wb.ActiveSheet.Name)
        com(lambda: wb.SaveAs(Filename=ruta_salida, FileFormat=XL_OPENXML_WORKBOOK))
        activo_cambiado = False
        for ws in com(lambda: list(wb.Worksheets)):
            com(lambda: ws.Cells.Copy())
            try:
                com(lambda: ws.Cells.PasteSpecial(Paste=XL_PASTE_VALUES))
            except COM_ERROR:
                com(lambda: ws.Activate())
                activo_cambiado = True
                com(lambda: ws.Cells.PasteSpecial(Paste=XL_PASTE_VALUES))
        com(lambda: setattr(xl, "CutCopyMode", False))
        if activo_cambiado:
            com(lambda: wb.Worksheets(hoja_activa).Activate())
        for ws in com(lambda: list(wb.Worksheets)):
            if com(lambda: ws.UsedRange.HasFormula) is not False:
                raise ValueError(f"La hoja {ws.Name} quedó con fórmulas después de pegar valores")
        com(lambda: wb.Save())
        com(lambda: wb.Close(SaveChanges=False))
        wb = None
    finally:
        if wb is not None:
            try:
                com(lambda: wb.Close(SaveChanges=False), seg_max=30)
            except Exception:
                pass
        for prop, val in (("DisplayAlerts", True), ("ScreenUpdating", True),
                          ("EnableEvents", eventos_antes), ("Calculation", calc_antes)):
            try:
                com(lambda: setattr(xl, prop, val), seg_max=60)
            except Exception:
                pass
        if seguridad_antes is not None:
            try:
                xl.AutomationSecurity = seguridad_antes
            except Exception:
                pass
        try:
            os.remove(copia)
        except OSError:
            pass


def generar_fecha(xl, main, fecha):
    """Pone la fecha en el main abierto, espera, revisa y genera el reporte. Devuelve el detalle."""
    ruta = ruta_reporte(fecha)
    mtime_antes = os.path.getmtime(ruta) if os.path.isfile(ruta) else None
    tocado = False
    listo = False
    try:
        com(lambda: setattr(xl, "Calculation", XL_CALC_AUTOMATIC))
        com(lambda: setattr(main.Worksheets("MAIN").Range("D4"), "Value2",
                            float((fecha - dt.date(1899, 12, 30)).days)))
        esperar(SEG_ESPERA_BLOOMBERG, "esperando Bloomberg")

        # --- Revisión antes de guardar ---
        requesting, nombre, problema = estado_libro(main)
        if problema:
            raise ErrorBloombergConexion(f"Bloomberg responde '{problema}'")
        if nombre:
            log(f"   {nombre} celdas #¿NOMBRE?: Bloomberg no está cargado en Excel. No guardo nada; "
                f"lo activo y recalculo.")
            for _ in range(2):
                activar_bloomberg(xl)
                esperar(15, "esperando que cargue Bloomberg")
                com(lambda: xl.CalculateFull())
                esperar(SEG_ESPERA_BLOOMBERG, "esperando Bloomberg")
                requesting, nombre, problema = estado_libro(main)
                if not nombre:
                    log("   Bloomberg quedó cargado; sigo.")
                    break
            if nombre:
                raise ErrorBloombergNoCargado(f"{nombre} celdas siguen en #¿NOMBRE?")
        extra = 0
        while requesting > MAX_REQUESTING_OK and extra < SEG_ESPERA_EXTRA:
            esperar(10, f"Bloomberg aún cargando ({requesting} pendientes)")
            extra += 10
            requesting, nombre, problema = estado_libro(main)
        if nombre:
            raise ErrorBloombergNoCargado(f"{nombre} celdas en #¿NOMBRE?")
        if requesting > MAX_REQUESTING_OK:
            raise ValueError(f"Bloomberg no alcanzó a cargar ({requesting} celdas 'Requesting Data')")

        # --- Generar desde una copia ---
        tocado = True
        generar_desde_copia(xl, main, ruta)

        # --- Revisar el archivo guardado ---
        if not os.path.isfile(ruta):
            raise ValueError("Excel no dejó el archivo en la carpeta de salida")
        motivo = revisar_reporte(ruta, fecha)
        if motivo:
            raise ValueError(f"el archivo guardado quedó malo: {motivo}")
        listo = True
        return f"archivo verificado, 'Requesting Data' al guardar: {requesting}"
    except BaseException:
        if tocado and not listo:
            try:
                if os.path.isfile(ruta) and (mtime_antes is None or os.path.getmtime(ruta) != mtime_antes):
                    os.remove(ruta)
                    log(f"   Borré {os.path.basename(ruta)} a medio hacer de este intento.")
            except OSError:
                pass
        raise


def registrar(ruta_log, fecha, estado, seg, detalle):
    try:
        with open(ruta_log, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=";").writerow(
                [dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), f"{fecha:%Y%m%d}", estado, f"{seg:.0f}", detalle])
    except OSError:
        pass


# -----------------------------------------------------------------------------
# Programa principal
# -----------------------------------------------------------------------------
def main():
    print("=" * 78)
    print(f"  GENERADOR PTS FWD & ETF {VERSION} - no abre ni cierra Excel")
    print("=" * 78)
    ruta_main = buscar_main()
    log(f"Main: {ruta_main}")
    desde = pedir_fecha("Fecha DESDE", FECHA_DESDE_SUGERIDA)
    hasta = pedir_fecha("Fecha HASTA", FECHA_HASTA_SUGERIDA)
    if hasta < desde:
        raise ErrorFatal("La fecha HASTA es anterior a la fecha DESDE")
    habiles, feriados = dias_habiles(desde, hasta)

    existen = [d for d in habiles if os.path.isfile(ruta_reporte(d))]
    malos = []
    if REVISAR_EXISTENTES and existen:
        log(f"Revisando los {len(existen)} reportes que ya existen en el rango...")
        for i, d in enumerate(existen, 1):
            print(f"      {i}/{len(existen)} {d:%Y%m%d}   ", end="\r", flush=True)
            motivo = revisar_reporte(ruta_reporte(d), d)
            if motivo:
                malos.append(d)
                log(f"   {d:%Y%m%d} está malo: {motivo} -> se regenera")
    pendientes = sorted([d for d in habiles if d not in existen] + malos)

    print()
    print(f"  Días hábiles en el rango   : {len(habiles)}")
    if feriados:
        print(f"  Feriados excluidos         : {', '.join(f'{d:%Y%m%d}' for d in feriados)}")
    print(f"  Ya existen y están bien    : {len(existen) - len(malos)} (se saltan)")
    if malos:
        print(f"  Existen pero están malos   : {len(malos)} (se regeneran)")
    print(f"  A generar                  : {len(pendientes)}   (~{len(pendientes) * 0.8:.0f} min)")
    print("  Cierra los reportes aaaammdd.xlsx que tengas abiertos (no se pueden reemplazar si están abiertos).")
    print("  Mientras corre no uses ese Excel (el resto del PC sí).")
    print()
    if not pendientes:
        log("No hay nada que generar: todas las fechas del rango tienen su reporte bueno.")
        return
    if input("¿Partimos? (s/n): ").strip().lower() not in ("s", "si", "sí", "y"):
        log("Cancelado.")
        return

    xl, lo_abri = conectar_excel()
    com(lambda: setattr(xl, "Visible", True))
    if lo_abri:
        esperar_bloomberg_en_excel(xl)
    main_wb, main_del_script = obtener_main(xl, ruta_main)
    hoja_main = com(lambda: main_wb.Worksheets("MAIN"))
    d4_original = com(lambda: hoja_main.Range("D4").Formula)
    autosave_antes = None
    if not main_del_script:
        try:
            autosave_antes = main_wb.AutoSaveOn
            if autosave_antes:
                main_wb.AutoSaveOn = False
                log("   Apagué el Autoguardado del main mientras corre (se vuelve a prender al final).")
        except Exception:
            autosave_antes = None

    ruta_log = os.path.join(CARPETA_MAIN, "log_generador_main_abierto.csv")
    pantalla = PantallaEncendida()
    pantalla.iniciar()
    cola = deque(pendientes)
    intentos, ultimo_error = Counter(), {}
    hechos, abandonadas = [], []
    fallas_addin = 0
    total = len(pendientes)
    inicio = time.time()
    motivo_fin = ""
    try:
        while cola:
            if time.time() - inicio > MAX_HORAS * 3600:
                motivo_fin = f"se cumplió el tope de {MAX_HORAS} horas"
                break
            fecha = cola[0]
            extra = f"  (intento {intentos[fecha] + 1})" if intentos[fecha] else ""
            log(f"({len(hechos) + 1}/{total}) {fecha:%Y%m%d}{extra}")
            t = time.time()
            try:
                detalle = generar_fecha(xl, main_wb, fecha)
            except (KeyboardInterrupt, ErrorFatal):
                raise
            except ErrorExcel as e:
                raise ErrorFatal(f"{e}. No cierro Excel a la fuerza: revísalo y vuelve a correr el script "
                                 "(sigue desde donde quedó).")
            except ErrorBloombergNoCargado as e:
                fallas_addin += 1
                registrar(ruta_log, fecha, "ESPERA", time.time() - t, f"Bloomberg no cargado: {e}")
                if fallas_addin >= 2:
                    raise ErrorFatal(
                        "Bloomberg no está cargado en tu Excel (#¿NOMBRE?). No se guardó ningún reporte malo.\n"
                        "   Revisa que la pestaña Bloomberg funcione (Archivo > Opciones > Complementos >\n"
                        "   Complementos COM > Ir: marca Bloomberg) y vuelve a correr el script.")
                esperar(60, "reintento en")
                continue
            except ErrorBloombergConexion as e:
                registrar(ruta_log, fecha, "ESPERA", time.time() - t, str(e))
                log(f"   {e}. Espero 3 min (¿Terminal logueado?) y reintento.")
                esperar(180, "reintento en")
                continue
            except Exception as e:
                detalle = f"{type(e).__name__}: {e}"
                intentos[fecha] += 1
                ultimo_error[fecha] = detalle
                registrar(ruta_log, fecha, "ERROR", time.time() - t, detalle)
                log(f"   falló -> {detalle}")
                if len(cola) > 1:
                    cola.rotate(-1)
                if fecha in cola and intentos[fecha] >= MAX_INTENTOS_POR_FECHA:
                    cola.remove(fecha)
                    abandonadas.append(fecha)
                    log(f"   {fecha:%Y%m%d}: {intentos[fecha]} intentos sin éxito, la dejo para revisar.")
                continue
            seg = time.time() - t
            cola.popleft()
            hechos.append(fecha)
            fallas_addin = 0
            registrar(ruta_log, fecha, "OK", seg, detalle)
            log(f"      OK en {seg:.0f}s ({detalle}). Quedan {len(cola)}")
    except ErrorFatal as e:
        motivo_fin = str(e)
    except KeyboardInterrupt:
        motivo_fin = "detenido por el usuario (Ctrl+C)"
    finally:
        pantalla.detener()
        # deja el main como estaba
        try:
            com(lambda: setattr(xl, "Calculation", XL_CALC_AUTOMATIC), seg_max=60)
            if main_del_script:
                com(lambda: main_wb.Close(SaveChanges=False), seg_max=60)
            else:
                com(lambda: setattr(hoja_main.Range("D4"), "Formula", d4_original), seg_max=60)
                if autosave_antes:
                    try:
                        main_wb.AutoSaveOn = True
                    except Exception:
                        pass
        except Exception:
            pass

    print()
    print("=" * 78)
    log(f"RESUMEN: {len(hechos)} de {total} generados en {(time.time() - inicio) / 60:.0f} min.")
    if motivo_fin:
        log(f"Terminó antes: {motivo_fin}")
    for d in abandonadas:
        print(f"   SIN LOGRAR {d:%Y%m%d}: {ultimo_error.get(d, '')}")
    faltan = [d for d in pendientes if d not in hechos]
    if faltan:
        print(f"   Faltan: {', '.join(f'{d:%Y%m%d}' for d in faltan)}  -> vuelve a correrlo con el mismo rango.")
    else:
        print("   LISTO: todas las fechas pedidas tienen su reporte, sin #¿NOMBRE?.")
    if not main_del_script:
        print("   El main quedó con MAIN!D4 como estaba. Ciérralo sin guardar si Excel te pregunta.")
    print(f"   Log: {ruta_log}")
    print("=" * 78)


if __name__ == "__main__":
    try:
        main()
    except ErrorFatal as e:
        log(f"DETENIDO: {e}")
