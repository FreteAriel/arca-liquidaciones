"""
App principal - Flask + WebSocket-like polling
Interfaz web local para automatización ARCA + ARBA → Excel
+ Módulo de Liquidaciones: IIBB Local, Convenio Multilateral, IVA
Compatible con Railway (HEADLESS automático, PORT desde env)
"""

import asyncio
import calendar
import json
import os
import threading
import time
import uuid
from datetime import datetime, timedelta

# En Railway (o cualquier servidor sin display) usamos headless=True
# Detectamos por la variable de entorno que Railway inyecta automáticamente
HEADLESS = os.getenv("RAILWAY_ENVIRONMENT") is not None or os.getenv("HEADLESS", "0") == "1"

from flask import Flask, jsonify, render_template, request, send_file, abort

from arca_scraper import ARCAScraper
from arba_scraper import ARBAScraper
from excel_gen import generar_excel
from liquidacion_scraper import AutonomoLiquidador, IIBBLocalLiquidador, COMLiquidador, IVALiquidador

app = Flask(__name__)

# Estado de las sesiones activas
sessions: dict[str, dict] = {}
_sessions_lock = threading.Lock()

# Semáforo: máximo 2 browsers Playwright simultáneos (BUG-02)
browser_sem = threading.Semaphore(2)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# -----------------------------------------------------------------------
# Cleanup thread: elimina sesiones con más de 30 minutos (BUG-01)
# -----------------------------------------------------------------------

SESSION_TTL_MINUTES = 30

def _cleanup_sessions():
    """Elimina sesiones terminadas con más de SESSION_TTL_MINUTES de antigüedad."""
    while True:
        time.sleep(300)  # revisar cada 5 minutos
        cutoff = datetime.now() - timedelta(minutes=SESSION_TTL_MINUTES)
        with _sessions_lock:
            to_del = [
                sid for sid, s in sessions.items()
                if s.get("estado") in ("completado", "error")
                and s.get("_created_at", datetime.now()) < cutoff
            ]
            for sid in to_del:
                sessions.pop(sid, None)

_cleaner = threading.Thread(target=_cleanup_sessions, daemon=True)
_cleaner.start()


# -----------------------------------------------------------------------
# Rutas web
# -----------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/liquidaciones")
def liquidaciones():
    return render_template("liquidaciones.html")


@app.route("/iniciar", methods=["POST"])
def iniciar():
    """Lanza el proceso de automatización en un hilo separado."""
    data = request.get_json()
    cuit          = data.get("cuit", "").strip()
    password      = data.get("password", "").strip()
    password_arba = data.get("password_arba", "").strip()
    mes           = int(data.get("mes", 1))
    anio          = int(data.get("anio", datetime.now().year))
    saldo_favor_1p      = float(data.get("saldo_favor_1p", 0) or 0)
    saldo_favor_2p      = float(data.get("saldo_favor_2p", 0) or 0)
    alicuota_iibb       = float(data.get("alicuota_iibb", 0) or 0)
    saldo_anterior_iibb = float(data.get("saldo_anterior_iibb", 0) or 0)

    if not cuit or not password or not password_arba:
        return jsonify({"error": "Faltan datos requeridos (CUIT, clave ARCA y clave ARBA)"}), 400

    # Calcular primer y último día del mes seleccionado
    ultimo_dia = calendar.monthrange(anio, mes)[1]
    fecha_desde = f"01/{mes:02d}/{anio}"
    fecha_hasta = f"{ultimo_dia:02d}/{mes:02d}/{anio}"

    session_id = str(uuid.uuid4())
    with _sessions_lock:
        sessions[session_id] = {
            "estado": "iniciando",
            "logs": [],
            "progreso": 0,
            "archivo": None,
            "error": None,
            "_created_at": datetime.now(),
        }

    t = threading.Thread(
        target=_run_automation,
        args=(session_id, cuit, password, password_arba, fecha_desde, fecha_hasta, mes, anio,
              saldo_favor_1p, saldo_favor_2p, alicuota_iibb, saldo_anterior_iibb),
        daemon=True,
    )
    t.start()

    return jsonify({"session_id": session_id})


# ─────────────────────────────────────────────────────────────────────────────
# RUTAS DE LIQUIDACIÓN
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/liquidar-autonomo", methods=["POST"])
def liquidar_autonomo():
    """Genera VEP de pago mensual de autónomos en ARCA SETI."""
    data = request.get_json()
    cuit      = data.get("cuit", "").strip()
    password  = data.get("password", "").strip()
    mes       = int(data.get("mes", 1))
    anio      = int(data.get("anio", datetime.now().year))
    categoria = data.get("categoria", "").strip()
    importe   = data.get("importe")   # None = usa el pre-cargado
    medio_pago = data.get("medio_pago", "qr").strip()

    if importe is not None:
        try:
            importe = float(str(importe).replace(".", "").replace(",", "."))
        except Exception:
            importe = None

    if not cuit or not password:
        return jsonify({"error": "Faltan CUIT y clave fiscal ARCA"}), 400
    if not categoria:
        return jsonify({"error": "Falta la categoría de autónomo"}), 400

    session_id = _crear_sesion()
    t = threading.Thread(
        target=_run_liquidacion_autonomo,
        args=(session_id, cuit, password, mes, anio, categoria, importe, medio_pago),
        daemon=True,
    )
    t.start()
    return jsonify({"session_id": session_id})


@app.route("/liquidar-iibb", methods=["POST"])
def liquidar_iibb():
    """Liquida IIBB Local en ARBA."""
    data = request.get_json()
    cuit          = data.get("cuit", "").strip()
    password_arba = data.get("password_arba", "").strip()
    mes           = int(data.get("mes", 1))
    anio          = int(data.get("anio", datetime.now().year))
    actividades   = data.get("actividades", [{"monto": 0}])
    saldo_ant     = float(data.get("saldo_favor_anterior", 0) or 0)
    medio_pago    = data.get("medio_pago", "qr").strip()

    if not cuit or not password_arba:
        return jsonify({"error": "Faltan CUIT y clave ARBA"}), 400

    session_id = _crear_sesion()
    t = threading.Thread(
        target=_run_liquidacion_iibb,
        args=(session_id, cuit, password_arba, mes, anio, actividades, saldo_ant, medio_pago),
        daemon=True,
    )
    t.start()
    return jsonify({"session_id": session_id})


@app.route("/liquidar-com", methods=["POST"])
def liquidar_com():
    """Liquida Convenio Multilateral en SIFERE y genera VEP."""
    data = request.get_json()
    cuit      = data.get("cuit", "").strip()
    password  = data.get("password", "").strip()
    mes       = int(data.get("mes", 1))
    anio      = int(data.get("anio", datetime.now().year))
    base_caba = float(data.get("base_caba", 0) or 0)
    base_bsas = float(data.get("base_bsas", 0) or 0)

    if not cuit or not password:
        return jsonify({"error": "Faltan CUIT y clave fiscal ARCA"}), 400

    session_id = _crear_sesion()
    t = threading.Thread(
        target=_run_liquidacion_com,
        args=(session_id, cuit, password, mes, anio, base_caba, base_bsas),
        daemon=True,
    )
    t.start()
    return jsonify({"session_id": session_id})


@app.route("/liquidar-iva", methods=["POST"])
def liquidar_iva():
    """Liquida posición IVA en ARCA y genera VEP."""
    data = request.get_json()
    cuit     = data.get("cuit", "").strip()
    password = data.get("password", "").strip()
    mes      = int(data.get("mes", 1))
    anio     = int(data.get("anio", datetime.now().year))

    if not cuit or not password:
        return jsonify({"error": "Faltan CUIT y clave fiscal ARCA"}), 400

    campos_iva = {k: float(data.get(k, 0) or 0) for k in [
        "vta_cf_neto_21", "vta_cf_iva_21", "vta_cf_neto_105", "vta_cf_iva_105",
        "vta_ri_neto_21", "vta_ri_iva_21", "vta_ri_neto_105", "vta_ri_iva_105",
        "vta_exentas",
        "cmp_neto_21", "cmp_iva_21", "cmp_neto_105", "cmp_iva_105",
        "retenciones", "saldo_favor_1p", "saldo_favor_2p",
    ]}
    medio_pago = data.get("medio_pago", "qr").strip()

    session_id = _crear_sesion()
    t = threading.Thread(
        target=_run_liquidacion_iva,
        args=(session_id, cuit, password, mes, anio, campos_iva, medio_pago),
        daemon=True,
    )
    t.start()
    return jsonify({"session_id": session_id})


# ─────────────────────────────────────────────────────────────────────────────
# Estado y descarga (compartidos)
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/estado/<session_id>")
def estado(session_id):
    """Polling de estado - el frontend consulta cada 2 segundos."""
    sess = sessions.get(session_id)
    if not sess:
        return jsonify({"error": "Sesión no encontrada"}), 404
    return jsonify(sess)


@app.route("/descargar/<session_id>")
def descargar(session_id):
    """Descarga el archivo generado (Excel o PDF/VEP)."""
    sess = sessions.get(session_id)
    if not sess or not sess.get("archivo"):
        abort(404)
    ruta = sess["archivo"]
    if not os.path.exists(ruta):
        abort(404)
    ext = os.path.splitext(ruta)[1].lower()
    mime = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pdf":  "application/pdf",
    }.get(ext, "application/octet-stream")
    return send_file(
        ruta,
        as_attachment=True,
        download_name=os.path.basename(ruta),
        mimetype=mime,
    )


# -----------------------------------------------------------------------
# Automatización principal (Libros IVA — existente)
# -----------------------------------------------------------------------

def _crear_sesion() -> str:
    session_id = str(uuid.uuid4())
    with _sessions_lock:
        sessions[session_id] = {
            "estado": "iniciando",
            "logs": [],
            "progreso": 0,
            "archivo": None,
            "error": None,
            "_created_at": datetime.now(),
        }
    return session_id


def _run_automation(session_id, cuit, password, password_arba, fecha_desde, fecha_hasta, mes, anio,
                    saldo_favor_1p=0.0, saldo_favor_2p=0.0, alicuota_iibb=0.0, saldo_anterior_iibb=0.0):
    """Ejecuta el flujo completo en un hilo con su propio event loop."""
    loop = asyncio.new_event_loop()
    try:
        with browser_sem:
            loop.run_until_complete(
                _automation(session_id, cuit, password, password_arba, fecha_desde, fecha_hasta, mes, anio,
                            saldo_favor_1p, saldo_favor_2p, alicuota_iibb, saldo_anterior_iibb)
            )
    except Exception as e:
        sessions[session_id]["estado"] = "error"
        sessions[session_id]["error"] = str(e)
        _log(session_id, f"❌ Error fatal: {e}")
    finally:
        loop.close()


async def _automation(session_id, cuit, password, password_arba, fecha_desde, fecha_hasta, mes, anio,
                      saldo_favor_1p=0.0, saldo_favor_2p=0.0, alicuota_iibb=0.0, saldo_anterior_iibb=0.0):
    sess = sessions[session_id]
    log = lambda msg: _log(session_id, msg)

    arca = ARCAScraper(cuit, password, log_fn=log)

    try:
        # ================================================================
        # PASO 1: ARCA — Comprobantes Emitidos (IVA Ventas)
        # ================================================================
        _set_progreso(session_id, 5)
        log("🚀 Iniciando ARCA...")
        await arca.start(headless=HEADLESS)

        _set_progreso(session_id, 10)
        await arca.login()

        _set_progreso(session_id, 25)
        log(f"📄 Obteniendo Comprobantes Emitidos ({fecha_desde} → {fecha_hasta})...")
        emitidos = await arca.get_comprobantes_emitidos(fecha_desde, fecha_hasta)
        log(f"   ✅ {len(emitidos)} comprobantes emitidos")

        # ================================================================
        # PASO 2: ARCA — Comprobantes Recibidos (IVA Compras)
        # ================================================================
        _set_progreso(session_id, 50)
        log(f"📄 Obteniendo Comprobantes Recibidos ({fecha_desde} → {fecha_hasta})...")
        recibidos = await arca.get_comprobantes_recibidos(fecha_desde, fecha_hasta)
        log(f"   ✅ {len(recibidos)} comprobantes recibidos")

        # ================================================================
        # PASO 3: ARCA — Retenciones SICORE 767 (IVA Compras)
        # ================================================================
        _set_progreso(session_id, 70)
        log(f"🔒 Obteniendo Retenciones SICORE 767 ({fecha_desde} → {fecha_hasta})...")
        sicore = []
        try:
            sicore = await arca.get_retenciones_sicore(fecha_desde, fecha_hasta)
            log(f"   ✅ {len(sicore)} retenciones SICORE obtenidas")
        except Exception as e_sicore:
            log(f"   ⚠️ No se pudieron obtener retenciones SICORE: {e_sicore}")
            sicore = []

        await arca.close()

        # ================================================================
        # PASO 4: ARBA — Deducciones Ingresos Brutos
        # ================================================================
        _set_progreso(session_id, 78)
        log("🏛️ Obteniendo deducciones ARBA (Ingresos Brutos)...")

        arba = ARBAScraper(cuit, password_arba, log_fn=log)
        deducciones = {"percepciones": [], "sitrac": [], "retenciones_bancarias": []}
        try:
            await arba.start(headless=HEADLESS)
            await arba.login()
            deducciones = await arba.get_deducciones(mes, anio)
            total_arba = sum(len(v) for v in deducciones.values())
            log(f"   ✅ ARBA: {total_arba} deducciones IIBB obtenidas")
        except Exception as e_arba:
            log(f"   ⚠️ ARBA no disponible: {e_arba} — se continúa sin deducciones IIBB")
        finally:
            try:
                await arba.close()
            except Exception:
                pass

        # ================================================================
        # PASO 5: Generar Excel con los libros IVA
        # ================================================================
        _set_progreso(session_id, 88)
        log("📊 Generando Excel — Libro IVA Ventas + Compras...")

        cuit_limpio    = cuit.replace("-", "").replace(".", "")
        nombre_archivo = f"Libros_IVA_{cuit_limpio}_{mes:02d}_{anio}.xlsx"
        ruta_excel     = os.path.join(OUTPUT_DIR, nombre_archivo)

        generar_excel(emitidos, recibidos, sicore, deducciones, ruta_excel,
                      saldo_favor_1p=saldo_favor_1p,
                      saldo_favor_2p=saldo_favor_2p,
                      alicuota_iibb=alicuota_iibb,
                      saldo_anterior_iibb=saldo_anterior_iibb,
                      log_fn=log)

        _set_progreso(session_id, 100)
        sess["archivo"] = ruta_excel
        sess["estado"]  = "completado"
        log(f"✅ Excel generado: {nombre_archivo}")
        log(f"   IVA Ventas: {len(emitidos)} filas | IVA Compras: {len(recibidos)} filas")
        log("📁 Descargá el archivo con el botón de abajo.")

    except Exception as e:
        sess["estado"] = "error"
        sess["error"]  = str(e)
        log(f"❌ Error: {e}")
        try:
            await arca.close()
        except Exception:
            pass
        raise


# -----------------------------------------------------------------------
# Runners de Liquidación
# -----------------------------------------------------------------------

def _run_liquidacion_autonomo(session_id, cuit, password, mes, anio,
                               categoria, importe, medio_pago):
    loop = asyncio.new_event_loop()
    log = lambda msg: _log(session_id, msg)
    try:
        _set_progreso(session_id, 5)
        log(f"👤 Iniciando pago Autónomo — {mes:02d}/{anio}")
        log(f"   Categoría: {categoria}")
        if importe:
            log(f"   Importe: $ {importe:,.2f}")
        else:
            log("   Importe: usa el máximo pre-cargado")
        log(f"   Medio de pago: {medio_pago}")

        liq = AutonomoLiquidador(cuit, password, log_fn=log, download_dir=OUTPUT_DIR)
        _set_progreso(session_id, 10)

        with browser_sem:
            resultado = loop.run_until_complete(
                liq.pagar_autonomo(
                    mes=mes,
                    anio=anio,
                    categoria=categoria,
                    importe=importe,
                    medio_pago=medio_pago,
                )
            )

        _set_progreso(session_id, 100)
        if resultado["resultado"] == "ok":
            sessions[session_id]["estado"]  = "completado"
            sessions[session_id]["archivo"] = resultado.get("pdf_path")
            numero_vep = resultado.get("numero_vep", "—")
            importe_real = resultado.get("importe")
            if importe_real:
                log(f"💰 VEP generado — N°: {numero_vep} — Importe: $ {importe_real:,.2f}")
            else:
                log(f"💰 VEP generado — N°: {numero_vep}")
            log("📄 Descargá el VEP con el botón de abajo.")
        else:
            sessions[session_id]["estado"] = "error"
            sessions[session_id]["error"]  = resultado.get("error")
            log(f"❌ {resultado.get('error')}")

    except Exception as e:
        sessions[session_id]["estado"] = "error"
        sessions[session_id]["error"]  = str(e)
        log(f"❌ Error fatal Autónomo: {e}")
    finally:
        loop.close()


def _run_liquidacion_iibb(session_id, cuit, password_arba, mes, anio, actividades, saldo_ant,
                           medio_pago="qr"):
    loop = asyncio.new_event_loop()
    log = lambda msg: _log(session_id, msg)
    try:
        _set_progreso(session_id, 5)
        log(f"🏛️ Iniciando liquidación IIBB Local — {mes:02d}/{anio}")
        log(f"   Actividades: {len(actividades)} | Saldo anterior: ${saldo_ant:,.2f}")

        liq = IIBBLocalLiquidador(cuit, password_arba, log_fn=log, download_dir=OUTPUT_DIR)
        _set_progreso(session_id, 10)

        with browser_sem:
            resultado = loop.run_until_complete(
                liq.liquidar(
                    anio=anio,
                    mes=mes,
                    actividades=actividades,
                    saldo_favor_anterior=saldo_ant,
                    medio_pago=medio_pago,
                )
            )

        _set_progreso(session_id, 100)
        res_tipo = resultado.get("resultado", "ok")
        if res_tipo in ("ok", "saldo_a_favor", "saldo_a_pagar"):
            sessions[session_id]["estado"] = "completado"
            # Archivo principal: VEP (si hay) o comprobante R-606M
            vep = resultado.get("vep") or {}
            archivo = vep.get("pdf_path") or resultado.get("comprobante")
            sessions[session_id]["archivo"] = archivo
            if res_tipo == "saldo_a_pagar":
                a_pagar = resultado.get("a_pagar", 0)
                vep_nro = vep.get("numero_vep", "—")
                log(f"💰 Saldo a pagar: ${a_pagar:,.2f} — VEP N°: {vep_nro}")
                log("📄 Descargá el VEP con el botón de abajo.")
            else:
                a_favor = resultado.get("a_favor", 0)
                log(f"✅ Saldo a favor del contribuyente: ${a_favor:,.2f}")
                log("📄 Descargá el comprobante R-606M con el botón de abajo.")
        else:
            sessions[session_id]["estado"] = "error"
            sessions[session_id]["error"]  = resultado.get("error")
            log(f"❌ {resultado.get('error')}")

    except Exception as e:
        sessions[session_id]["estado"] = "error"
        sessions[session_id]["error"]  = str(e)
        log(f"❌ Error fatal IIBB: {e}")
    finally:
        loop.close()


def _run_liquidacion_com(session_id, cuit, password, mes, anio, base_caba, base_bsas):
    loop = asyncio.new_event_loop()
    log = lambda msg: _log(session_id, msg)
    try:
        _set_progreso(session_id, 5)
        log(f"🗺️ Iniciando liquidación Convenio Multilateral — {mes:02d}/{anio}")
        log(f"   CABA: ${base_caba:,.2f} | Bs.As: ${base_bsas:,.2f} | Total: ${base_caba+base_bsas:,.2f}")

        liq = COMLiquidador(cuit, password, log_fn=log, download_dir=OUTPUT_DIR)
        _set_progreso(session_id, 10)

        with browser_sem:
            resultado = loop.run_until_complete(
                liq.liquidar(anio=anio, mes=mes, base_caba=base_caba, base_bsas=base_bsas)
            )

        _set_progreso(session_id, 100)
        if resultado["resultado"] == "ok":
            sessions[session_id]["estado"]  = "completado"
            sessions[session_id]["archivo"] = resultado.get("vep")
            total = resultado.get("total_a_pagar", 0)
            if total > 0:
                log(f"💰 Total VEP a pagar: ${total:,.2f}")
            log("📄 Descargá el VEP con el botón de abajo.")
        else:
            sessions[session_id]["estado"] = "error"
            sessions[session_id]["error"]  = resultado.get("error")
            log(f"❌ {resultado.get('error')}")

    except Exception as e:
        sessions[session_id]["estado"] = "error"
        sessions[session_id]["error"]  = str(e)
        log(f"❌ Error fatal COM: {e}")
    finally:
        loop.close()


def _run_liquidacion_iva(session_id, cuit, password, mes, anio, campos, medio_pago="qr"):
    loop = asyncio.new_event_loop()
    log = lambda msg: _log(session_id, msg)
    try:
        _set_progreso(session_id, 5)
        debito  = campos["vta_cf_iva_21"] + campos["vta_cf_iva_105"] + \
                  campos["vta_ri_iva_21"] + campos["vta_ri_iva_105"]
        credito = campos["cmp_iva_21"] + campos["cmp_iva_105"]
        exentas = campos.get("vta_exentas", 0.0)
        log(f"🧾 Iniciando liquidación IVA — {mes:02d}/{anio}")
        log(f"   Débito fiscal: ${debito:,.2f} | Crédito fiscal: ${credito:,.2f}")
        if exentas > 0:
            log(f"   Ventas exentas / no gravadas: ${exentas:,.2f}")
        log(f"   Posición estimada: ${debito - credito:,.2f}")

        liq = IVALiquidador(cuit, password, log_fn=log, download_dir=OUTPUT_DIR)
        _set_progreso(session_id, 10)

        with browser_sem:
            resultado = loop.run_until_complete(
                liq.liquidar(anio=anio, mes=mes, **campos, medio_pago=medio_pago)
            )

        _set_progreso(session_id, 100)
        res_tipo = resultado.get("resultado", "ok")
        if res_tipo in ("ok", "saldo_a_favor", "saldo_a_pagar"):
            sessions[session_id]["estado"]  = "completado"
            # Archivo: VEP si hay saldo a pagar, F.2051 si hay saldo a favor
            vep = resultado.get("vep") or {}
            archivo = vep.get("pdf_path") or resultado.get("pdf_f2051") or resultado.get("vep_path")
            sessions[session_id]["archivo"] = archivo
            posicion = resultado.get("posicion", 0)
            if res_tipo == "saldo_a_pagar":
                vep_nro = vep.get("numero_vep", "—")
                log(f"💰 Saldo a pagar IVA: ${posicion:,.2f} — VEP N°: {vep_nro}")
                log("📄 Descargá el VEP con el botón de abajo.")
            else:
                log(f"✅ Saldo a favor IVA: ${abs(posicion):,.2f}")
                log("📄 Descargá el F.2051 con el botón de abajo.")
        else:
            sessions[session_id]["estado"] = "error"
            sessions[session_id]["error"]  = resultado.get("error")
            log(f"❌ {resultado.get('error')}")

    except Exception as e:
        sessions[session_id]["estado"] = "error"
        sessions[session_id]["error"]  = str(e)
        log(f"❌ Error fatal IVA: {e}")
    finally:
        loop.close()


# -----------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------

def _log(session_id: str, msg: str):
    if session_id in sessions:
        ts = datetime.now().strftime("%H:%M:%S")
        sessions[session_id]["logs"].append(f"[{ts}] {msg}")


def _set_progreso(session_id: str, valor: int):
    if session_id in sessions:
        sessions[session_id]["progreso"] = valor
        # No sobreescribir estado si ya fue marcado como completado/error (BUG-07)
        if sessions[session_id]["estado"] not in ("completado", "error"):
            sessions[session_id]["estado"] = "procesando"


# -----------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    port = int(os.getenv("PORT", 5000))
    host = "0.0.0.0" if os.getenv("RAILWAY_ENVIRONMENT") else "127.0.0.1"
    print("=" * 60)
    print("  ARCA + ARBA -> Libros IVA + Liquidaciones")
    if host == "127.0.0.1":
        print(f"  Abri tu navegador en: http://localhost:{port}")
        print(f"  Liquidaciones en:     http://localhost:{port}/liquidaciones")
    else:
        print(f"  Corriendo en Railway — puerto {port}")
        print(f"  Modo headless: {HEADLESS}")
    print("=" * 60)
    app.run(host=host, port=port, debug=False, threaded=True)
