"""
Scrapers de liquidación impositiva
===================================
- AutonomoLiquidador   : Genera VEP pago mensual de autónomos en ARCA SETI
- IIBBLocalLiquidador  : Presenta DJ Anticipo IIBB en ARBA (arba.gov.ar)
- COMLiquidador        : Presenta DJ CM03 en SIFERE/COMARB y genera VEP
- IVALiquidador        : Liquida posición IVA en ARCA y genera VEP
"""

import asyncio
import os
import re
import urllib.request
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# Detectar entorno Railway para modo headless automático
HEADLESS = os.getenv("RAILWAY_ENVIRONMENT") is not None or os.getenv("HEADLESS", "0") == "1"


# ───────────────────────────────────────────────────────────────────────────────
# Helpers compartidos
# ───────────────────────────────────────────────────────────────────────────────

def _get_system_proxy():
    try:
        proxies = urllib.request.getproxies()
        return proxies.get("https") or proxies.get("http") or None
    except Exception:
        return None


def _find_browser():
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def _fmt_monto(valor: float) -> str:
    """Formatea un número como string sin separador de miles, con coma decimal."""
    return f"{valor:.2f}".replace(".", ",")


# ───────────────────────────────────────────────────────────────────────────────
# Base scraper con lifecycle común
# ───────────────────────────────────────────────────────────────────────────────

class _BaseScraper:
    def __init__(self, cuit: str, password: str, log_fn=print, download_dir: str = None):
        self.cuit = re.sub(r"[^0-9]", "", cuit)
        self.password = password
        self.log = log_fn
        self.download_dir = download_dir or os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(self.download_dir, exist_ok=True)
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    async def start(self, headless: bool = False):
        self.playwright = await async_playwright().start()
        proxy = _get_system_proxy()
        proxy_cfg = {"server": proxy} if proxy else None
        extra_args = ["--start-maximized", "--no-sandbox", "--disable-dev-shm-usage"]
        browser_path = _find_browser()
        launch_kwargs = dict(headless=headless, args=extra_args, proxy=proxy_cfg)
        if browser_path:
            self.log(f"🌐 Usando browser: {browser_path}")
            launch_kwargs["executable_path"] = browser_path
        try:
            self.browser = await self.playwright.chromium.launch(**launch_kwargs)
        except Exception:
            launch_kwargs.pop("executable_path", None)
            try:
                self.browser = await self.playwright.chromium.launch(channel="chrome", **launch_kwargs)
            except Exception:
                self.browser = await self.playwright.chromium.launch(**launch_kwargs)

        self.context = await self.browser.new_context(
            accept_downloads=True,
            viewport={"width": 1280, "height": 900},
        )
        self.page = await self.context.new_page()
        # Timeout por defecto para todos los selectores / acciones (ROB-02)
        self.page.set_default_timeout(60000)

    async def close(self):
        try:
            if self.context:
                await self.context.close()
            if self.browser:
                await self.browser.close()
            if self.playwright:
                await self.playwright.stop()
        except Exception:
            pass


# ───────────────────────────────────────────────────────────────────────────────
# Medios de pago ARCA SETI — textos reconocibles en los botones del portal
# ───────────────────────────────────────────────────────────────────────────────

MEDIOS_PAGO_TEXTOS = {
    "qr":             ["QR", "Pago con QR", "Cobro con QR"],
    "pagar":          ["Pagar", "Pagar.com", "Pagar ("],
    "pagomiscuentas": ["PagoMisCuentas", "Banelco", "PagoMis"],
    "interbanking":   ["Interbanking"],
    "xngroup":        ["XN Group", "XN Latin", "XN Group Latin"],
}


async def _seleccionar_medio_pago(page, medio: str, log_fn=print):
    """
    Hace click en el botón del medio de pago indicado.
    medio: "qr" | "pagar" | "pagomiscuentas" | "interbanking" | "xngroup"
    """
    textos = MEDIOS_PAGO_TEXTOS.get(medio.lower(), ["QR"])
    for texto in textos:
        try:
            btn = page.get_by_text(texto, exact=False)
            if await btn.count() > 0:
                await btn.first.click()
                log_fn(f"   ✅ Medio de pago seleccionado: {texto}")
                return True
        except Exception:
            continue
    # Fallback: click en el primero disponible
    log_fn(f"   ⚠️ No se encontró '{medio}' — se intenta con el primer medio disponible")
    try:
        btns = await page.query_selector_all(
            "button.medio-pago, div.medio-pago, li.medio-pago, "
            "img[alt*='pago'], button:has-text('Pagar'), div[role='button']"
        )
        if btns:
            await btns[0].click()
            return True
    except Exception:
        pass
    return False


# ═══════════════════════════════════════════════════════════════════════════════
# AUTÓNOMO — ARCA SETI
# ═══════════════════════════════════════════════════════════════════════════════

class AutonomoLiquidador(_BaseScraper):
    """
    Genera el VEP de pago mensual de autónomos en ARCA SETI.

    Flujo:
        ARCA login → SETI → Pagos → Gestión de VEP → Nuevo VEP
        → Grupo "Autonomo" → Tipo "Autonomo - Pago Mensual"
        → Período mes/año + Categoría + Importe (aportes seg. social 308)
        → VEPs a enviar → Seleccioná medio de pago → Aceptar
        → Descargar VEP PDF

    Parámetros de entrada:
        mes       : int   — Mes del período (1-12)
        anio      : int   — Año del período
        categoria : str   — Código + descripción (ej: "301: T3 Cat I Ingresos hasta $25.000")
                            Puede ser solo el código numérico (ej: "301")
        importe   : float — Importe a abonar. None = usa el pre-cargado por el sistema
        medio_pago: str   — "qr"|"pagar"|"pagomiscuentas"|"interbanking"|"xngroup"
    """

    ARCA_LOGIN_URL = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"
    SETI_URL       = "https://seti.afip.gob.ar/setiweb/"
    SETI_NUEVO_VEP = "https://seti.afip.gob.ar/setiweb/#/pago/nuevo-vep?op=1"
    SETI_VEPS      = "https://seti.afip.gob.ar/setiweb/#/pago/veps-a-enviar"

    # ------------------------------------------------------------------
    async def login_arca(self):
        """Login en ARCA con Clave Fiscal."""
        self.log("🔐 Iniciando sesión en ARCA...")
        await self.page.goto(self.ARCA_LOGIN_URL, wait_until="networkidle", timeout=30000)

        cuit_fmt = (
            f"{self.cuit[:2]}-{self.cuit[2:10]}-{self.cuit[10:]}"
            if len(self.cuit) == 11 else self.cuit
        )
        await self.page.wait_for_selector("#F1\\:username, input[name='username']", timeout=15000)
        try:
            await self.page.fill("#F1\\:username", cuit_fmt)
        except Exception:
            await self.page.fill("input[name='username']", cuit_fmt)

        await self.page.click("#F1\\:btnSiguiente, button:has-text('Siguiente')")
        await self.page.wait_for_load_state("networkidle", timeout=15000)

        await self.page.wait_for_selector("input[type='password']", timeout=15000)
        await self.page.fill("input[type='password']", self.password)
        await self.page.click("#F1\\:btnIngresar, button:has-text('Ingresar'), button[type='submit']")
        await self.page.wait_for_load_state("networkidle", timeout=30000)
        self.log("✅ Login ARCA exitoso")

    # ------------------------------------------------------------------
    async def navegar_seti(self):
        """Navega al portal SETI."""
        self.log("🌐 Navegando a SETI...")
        await self.page.goto(self.SETI_URL, wait_until="networkidle", timeout=30000)
        # Cerrar popup "Novedades" si aparece
        try:
            await self.page.click(
                "button:has-text('Entendido'), button:has-text('Cerrar'), "
                "button:has-text('Aceptar'), [aria-label='Cerrar']",
                timeout=5000
            )
        except Exception:
            pass
        self.log("✅ En SETI")

    # ------------------------------------------------------------------
    async def crear_vep_autonomo(self, mes: int, anio: int,
                                  categoria: str, importe: float = None) -> float:
        """
        Completa el formulario de Nuevo VEP para autónomos.
        Retorna el importe final del VEP (puede diferir si se usó el pre-cargado).
        """
        self.log(f"📝 Creando VEP Autónomo — {mes:02d}/{anio} — Categoría: {categoria}")

        # Navegar a Pagos → Gestión VEP → Nuevo VEP
        await self.page.goto(self.SETI_NUEVO_VEP, wait_until="networkidle", timeout=30000)
        await self.page.wait_for_load_state("networkidle", timeout=20000)

        # Esperar que cargue el formulario
        await self.page.wait_for_timeout(2000)

        # ── Paso 1: CUIT / Organismo / Grupo / Tipo ──────────────────────────
        self.log("   Paso 1: Seleccionando organismo y tipo de pago...")

        # CUIT — si hay dropdown de contribuyentes, elegir el propio CUIT
        try:
            cuit_fmt = f"{self.cuit[:2]}-{self.cuit[2:10]}-{self.cuit[10:]}"
            await self.page.select_option(
                "select[formcontrolname*='cuit'], select[id*='cuit'], select[name*='cuit']",
                label=cuit_fmt, timeout=5000
            )
        except Exception:
            pass

        # Organismo Recaudador → ARCA (suele venir prellenado)
        try:
            await self.page.select_option(
                "select[formcontrolname*='organismo'], select[id*='organismo']",
                label="ARCA", timeout=5000
            )
        except Exception:
            pass

        # Grupos de Tipos de Pagos → "Autonomo"
        try:
            await self.page.select_option(
                "select[formcontrolname*='grupo'], select[id*='grupo'], select[name*='grupo']",
                label="Autonomo", timeout=8000
            )
            await self.page.wait_for_timeout(1500)
        except Exception:
            # Intentar por texto visible
            try:
                await self.page.get_by_label("Grupos de Tipos de Pagos").select_option(
                    label="Autonomo"
                )
            except Exception:
                self.log("   ⚠️ No se pudo seleccionar grupo 'Autonomo'")

        # Tipo de Pago → "Autonomo - Pago Mensual"
        try:
            await self.page.select_option(
                "select[formcontrolname*='tipo'], select[id*='tipo'], select[name*='tipo']",
                label="Autonomo - Pago Mensual", timeout=8000
            )
        except Exception:
            try:
                await self.page.get_by_label("Tipo de Pago").select_option(
                    label="Autonomo - Pago Mensual"
                )
            except Exception:
                self.log("   ⚠️ No se pudo seleccionar tipo 'Autonomo - Pago Mensual'")

        # Click "Siguiente"
        await self.page.click(
            "button:has-text('Siguiente'), input[value='Siguiente'], "
            "button:has-text('SIGUIENTE')"
        )
        await self.page.wait_for_load_state("networkidle", timeout=20000)
        await self.page.wait_for_timeout(2000)

        # ── Paso 2: Período + Categoría + Importe ────────────────────────────
        self.log(f"   Paso 2: Ingresando período {mes:02d}/{anio} y categoría...")

        # Mes del período
        try:
            await self.page.select_option(
                "select[formcontrolname*='mes'], select[id*='mes'], select[name*='mes']",
                value=str(mes), timeout=5000
            )
        except Exception:
            try:
                await self.page.fill(
                    "input[formcontrolname*='mes'], input[id*='mes']", str(mes)
                )
            except Exception:
                pass

        # Año del período
        try:
            await self.page.select_option(
                "select[formcontrolname*='anio'], select[id*='anio'], select[name*='anio'], "
                "select[formcontrolname*='year'], select[id*='year']",
                value=str(anio), timeout=5000
            )
        except Exception:
            try:
                await self.page.fill(
                    "input[formcontrolname*='anio'], input[id*='anio'], "
                    "input[formcontrolname*='year']", str(anio)
                )
            except Exception:
                pass

        # Categoría / CRA
        try:
            # Intentar match exacto primero, luego por código numérico
            cod = categoria.split(":")[0].strip() if ":" in categoria else categoria.strip()
            for lbl in [categoria, cod]:
                try:
                    await self.page.select_option(
                        "select[formcontrolname*='categoria'], select[id*='categoria'], "
                        "select[formcontrolname*='cra'], select[id*='cra']",
                        label=lbl, timeout=5000
                    )
                    break
                except Exception:
                    continue
        except Exception:
            self.log(f"   ⚠️ No se pudo seleccionar categoría '{categoria}'")

        await self.page.wait_for_timeout(2000)  # esperar que se pre-cargue el importe

        # Importe — aportes seg. social autónomos (código 308)
        importe_final = importe
        if importe is not None:
            try:
                campo_imp = await self.page.query_selector(
                    "input[formcontrolname*='importe'], input[id*='importe'], "
                    "input[formcontrolname*='monto'], input[id*='monto'], "
                    "input[formcontrolname*='aportes']"
                )
                if campo_imp:
                    await campo_imp.triple_click()
                    await campo_imp.fill(str(importe).replace(".", ","))
                    self.log(f"   Importe ingresado: $ {importe:,.2f}")
            except Exception as e:
                self.log(f"   ⚠️ No se pudo modificar importe: {e}")
        else:
            # Leer el importe pre-cargado para devolverlo
            try:
                campo_imp = await self.page.query_selector(
                    "input[formcontrolname*='importe'], input[id*='importe'], "
                    "input[formcontrolname*='monto'], input[id*='monto'], "
                    "input[formcontrolname*='aportes']"
                )
                if campo_imp:
                    val = await campo_imp.input_value()
                    importe_final = float(val.replace(".", "").replace(",", ".")) if val else None
                    if importe_final:
                        self.log(f"   Importe pre-cargado: $ {importe_final:,.2f}")
            except Exception:
                pass

        # Click "Siguiente"
        await self.page.click(
            "button:has-text('Siguiente'), input[value='Siguiente'], "
            "button:has-text('SIGUIENTE')"
        )
        await self.page.wait_for_load_state("networkidle", timeout=20000)
        await self.page.wait_for_timeout(2000)
        self.log("   ✅ VEP creado — en lista de VEPs a enviar")
        return importe_final

    # ------------------------------------------------------------------
    async def seleccionar_y_enviar_vep(self, medio_pago: str = "qr") -> str | None:
        """
        En la pantalla 'VEPs a enviar', marca el VEP y selecciona el medio de pago.
        Retorna el número de VEP si lo puede leer.
        """
        self.log(f"💳 Seleccionando medio de pago: {medio_pago}...")

        # Marcar el checkbox del VEP
        try:
            checkbox = await self.page.query_selector(
                "input[type='checkbox'], mat-checkbox, "
                "td:first-child input[type='checkbox']"
            )
            if checkbox:
                await checkbox.check()
                await self.page.wait_for_timeout(500)
        except Exception:
            pass

        # Click en "Seleccioná medio de pago"
        await self.page.click(
            "button:has-text('Seleccioná medio de pago'), "
            "button:has-text('Selecciona medio de pago'), "
            "button:has-text('Seleccionar medio'), "
            "a:has-text('medio de pago')",
            timeout=10000
        )
        await self.page.wait_for_load_state("networkidle", timeout=15000)
        await self.page.wait_for_timeout(1500)

        # Seleccionar el medio de pago
        await _seleccionar_medio_pago(self.page, medio_pago, self.log)
        await self.page.wait_for_timeout(500)

        # Click "Aceptar"
        try:
            await self.page.click(
                "button:has-text('Aceptar'), button:has-text('ACEPTAR'), "
                "button:has-text('Continuar'), button:has-text('Confirmar')",
                timeout=8000
            )
            await self.page.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass

        await self.page.wait_for_timeout(2000)

        # Leer número de VEP de la pantalla final
        numero_vep = None
        try:
            texto = await self.page.inner_text("body")
            match = re.search(r"N[°º]\s*(?:de\s+)?VEP[:\s]*([\d]+)", texto, re.IGNORECASE)
            if not match:
                match = re.search(r"VEP[:\s#Nº]*\s*([\d]{6,})", texto, re.IGNORECASE)
            if match:
                numero_vep = match.group(1)
                self.log(f"   VEP N°: {numero_vep}")
        except Exception:
            pass

        return numero_vep

    # ------------------------------------------------------------------
    async def descargar_vep_pdf(self) -> str | None:
        """Descarga el PDF del VEP desde la pantalla de confirmación."""
        self.log("📄 Descargando PDF del VEP...")
        try:
            async with self.page.expect_download(timeout=25000) as dl_info:
                await self.page.click(
                    "button:has-text('Descargar VEP'), a:has-text('Descargar VEP'), "
                    "button:has-text('Descargar'), a:has-text('Descargar')",
                    timeout=10000
                )
            download = await dl_info.value
            fname = download.suggested_filename or "VEP_Autonomo.pdf"
            path = os.path.join(self.download_dir, fname)
            await download.save_as(path)
            self.log(f"✅ PDF descargado: {os.path.basename(path)}")
            return path
        except Exception as e:
            self.log(f"   ⚠️ No se pudo descargar PDF: {e}")
            # Intentar captura de pantalla como fallback
            try:
                path_png = os.path.join(self.download_dir, "VEP_Autonomo_screen.png")
                await self.page.screenshot(path=path_png, full_page=True)
                self.log(f"   📸 Captura de pantalla guardada: {os.path.basename(path_png)}")
                return path_png
            except Exception:
                return None

    # ------------------------------------------------------------------
    async def pagar_autonomo(self, mes: int, anio: int,
                              categoria: str,
                              importe: float = None,
                              medio_pago: str = "qr") -> dict:
        """
        Flujo completo de pago mensual de autónomos.
        Retorna dict con numero_vep, importe, pdf_path, resultado.
        """
        resultado = {
            "resultado": "ok",
            "numero_vep": None,
            "importe": importe,
            "pdf_path": None,
            "error": None,
        }
        try:
            await self.start(headless=HEADLESS)
            await self.login_arca()
            await self.navegar_seti()
            importe_real = await self.crear_vep_autonomo(mes, anio, categoria, importe)
            resultado["importe"] = importe_real
            numero_vep = await self.seleccionar_y_enviar_vep(medio_pago)
            resultado["numero_vep"] = numero_vep
            pdf = await self.descargar_vep_pdf()
            resultado["pdf_path"] = pdf
            self.log(f"✅ Pago autónomo completado — VEP N°: {numero_vep}")
        except Exception as e:
            resultado["resultado"] = "error"
            resultado["error"] = str(e)
            self.log(f"❌ Error en pago autónomo: {e}")
        finally:
            await self.close()
        return resultado


# ═══════════════════════════════════════════════════════════════════════════════
# IIBB LOCAL — ARBA
# ═══════════════════════════════════════════════════════════════════════════════

class IIBBLocalLiquidador(_BaseScraper):
    """
    Presenta la DJ Anticipo de Ingresos Brutos Local en ARBA.

    Parámetros de entrada:
        anio                  : int  — Año del período (ej: 2026)
        mes                   : int  — Mes del período (1-12)
        actividades           : list[dict] con claves:
                                  'codigo'  (str, opcional para referencia)
                                  'monto'   (float, base imponible gravada)
        saldo_favor_anterior  : float — Saldo acumulado a favor de períodos anteriores
    """

    ARBA_LOGIN_URL = "https://sso.arba.gov.ar/Login/login?service=https%3A%2F%2Fwww.arba.gov.ar%2FGestionar%2FPanelAutogestion.asp"
    IIBB_PRES_URL  = "https://app.arba.gov.ar/IBPresentaciones/"

    # ------------------------------------------------------------------
    async def login_arba(self, cuit: str, password_arba: str):
        """Login SSO de ARBA con CUIT y clave ARBA."""
        self.log("🔐 Navegando al login de ARBA...")
        await self.page.goto(self.ARBA_LOGIN_URL, wait_until="networkidle", timeout=30000)

        # Campo CUIT / usuario
        await self.page.wait_for_selector("input[name='username'], input[id='username'], input[type='text']", timeout=15000)
        await self.page.fill("input[name='username'], input[id='username'], input[type='text']", cuit)

        # Campo contraseña
        await self.page.fill("input[type='password']", password_arba)

        # Botón ingresar
        await self.page.click("button[type='submit'], input[type='submit'], button:has-text('Ingresar')")
        await self.page.wait_for_load_state("networkidle", timeout=30000)
        self.log("✅ Login ARBA exitoso")

    # ------------------------------------------------------------------
    async def navegar_a_iibb(self):
        """Navega al módulo de Presentaciones IIBB."""
        self.log("📋 Navegando a Ingresos Brutos — Presentaciones de DJ...")

        # Desde el panel de autogestión, buscar el link de IIBB
        try:
            link = self.page.get_by_text("Presentá tu declaración jurada de Ingresos Brutos")
            await link.click(timeout=10000)
            await self.page.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            # Navegar directo al módulo
            await self.page.goto(self.IIBB_PRES_URL, wait_until="networkidle", timeout=20000)

        self.log("✅ En módulo IIBB Presentaciones")

    # ------------------------------------------------------------------
    async def iniciar_dj(self, anio: int, mes: int):
        """Inicia una nueva DJ Anticipo para el período indicado."""
        self.log(f"📝 Iniciando DJ Anticipo {mes:02d}/{anio}...")

        # Menú Presentación → DJ Anticipo → Inicio
        await self.page.click("text=Presentación")
        await self.page.wait_for_selector("text=Dj Anticipo", timeout=10000)
        await self.page.click("text=Dj Anticipo")
        await self.page.wait_for_selector("text=Inicio", timeout=5000)
        await self.page.click("text=Inicio")
        await self.page.wait_for_load_state("networkidle", timeout=20000)

        # Seleccionar Régimen Mensual
        try:
            await self.page.select_option("select[name*='egimen'], select[id*='egimen']", label="Mensual")
        except Exception:
            pass

        # Seleccionar Año
        try:
            await self.page.select_option("select[name*='nio'], select[id*='nio']", str(anio))
        except Exception:
            await self.page.fill("input[name*='nio'], input[id*='nio']", str(anio))

        # Seleccionar Mes
        try:
            await self.page.select_option("select[name*='es'], select[id*='es']", str(mes))
        except Exception:
            await self.page.fill("input[name*='es'], input[id*='es']", str(mes))

        # Hacer click en "Iniciar DJ"
        await self.page.click("button:has-text('Iniciar DJ'), input[value*='Iniciar']")
        await self.page.wait_for_load_state("networkidle", timeout=30000)
        self.log(f"✅ DJ iniciada para {mes:02d}/{anio}")

    # ------------------------------------------------------------------
    async def cargar_actividades(self, actividades: list):
        """
        Carga el monto imponible en cada actividad de la DJ.
        actividades: list[dict] con 'monto' (float)
        """
        self.log(f"📊 Cargando {len(actividades)} actividad(es)...")

        # Navegar a Carga de DJ desde el menú
        try:
            await self.page.click("text=Presentación")
            await self.page.wait_for_selector("text=Dj Anticipo", timeout=5000)
            await self.page.click("text=Dj Anticipo")
            await self.page.wait_for_selector("text=Carga de DJ", timeout=5000)
            await self.page.click("text=Carga de DJ")
            await self.page.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass

        for i, act in enumerate(actividades):
            monto = float(act.get("monto", 0))
            codigo = act.get("codigo", f"actividad {i+1}")
            self.log(f"   Actividad {codigo}: $ {monto:,.2f}")

            try:
                # Buscar botón "Carga de la DJ" o el ícono de editar de cada fila
                carga_links = await self.page.query_selector_all(
                    "a:has-text('Carga de la DJ'), a:has-text('Modificar'), "
                    "img[title*='Editar'], button:has-text('Carga')"
                )
                if i < len(carga_links):
                    await carga_links[i].click()
                    await self.page.wait_for_load_state("networkidle", timeout=15000)

                    # Ingresar monto imponible
                    campo_monto = await self.page.query_selector(
                        "input[name*='monto'], input[name*='imponible'], "
                        "input[id*='monto'], input[id*='imponible'], "
                        "input[class*='monto']"
                    )
                    if campo_monto:
                        await campo_monto.triple_click()
                        await campo_monto.fill(str(monto).replace(".", ","))

                    # Confirmar
                    await self.page.click(
                        "button:has-text('Modificar'), button:has-text('Guardar'), "
                        "input[value*='Modificar'], input[value*='Guardar']"
                    )
                    await self.page.wait_for_load_state("networkidle", timeout=15000)
                    self.log(f"   ✅ Actividad {i+1} cargada")
            except Exception as e:
                self.log(f"   ⚠️ No se pudo cargar actividad {i+1}: {e}")

    # ------------------------------------------------------------------
    async def cerrar_dj(self, saldo_favor_anterior: float = 0.0):
        """Cierra la DJ y envía — opcionalmente ingresa saldo a favor anterior."""
        self.log("📤 Cerrando y enviando DJ...")

        # Navegar a Consulta DDJJ pendientes o Cierre desde el menú
        try:
            await self.page.click("text=Presentación")
            await self.page.wait_for_selector("text=Dj Anticipo", timeout=5000)
            await self.page.click("text=Dj Anticipo")
            try:
                await self.page.wait_for_selector("text=Consulta DDJJ pendientes", timeout=3000)
                await self.page.click("text=Consulta DDJJ pendientes")
                await self.page.wait_for_load_state("networkidle", timeout=15000)
                # Abrir la DJ pendiente
                await self.page.click("a:has-text('Ver'), a:has-text('Abrir'), td a", timeout=5000)
                await self.page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
        except Exception:
            pass

        # Ingresar saldo a favor anterior si > 0
        if saldo_favor_anterior > 0:
            try:
                campo_saldo = await self.page.query_selector(
                    "input[name*='saldo'], input[id*='saldo'], "
                    "input[placeholder*='saldo'], input[placeholder*='Saldo']"
                )
                if campo_saldo:
                    await campo_saldo.triple_click()
                    await campo_saldo.fill(str(saldo_favor_anterior).replace(".", ","))
                    self.log(f"   Saldo a favor anterior: $ {saldo_favor_anterior:,.2f}")
                    # Buscar botón "Recalcular" o "Actualizar"
                    try:
                        await self.page.click(
                            "button:has-text('Recalcular'), button:has-text('Actualizar'), "
                            "a:has-text('Ingrese saldo y recalcule')"
                        )
                        await self.page.wait_for_load_state("networkidle", timeout=10000)
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"   ⚠️ No se pudo ingresar saldo anterior: {e}")

        # Click en "Enviar"
        try:
            await self.page.click("button:has-text('Enviar'), input[value*='Enviar']")
            await self.page.wait_for_load_state("networkidle", timeout=30000)
            self.log("✅ DJ enviada correctamente")
        except Exception as e:
            self.log(f"⚠️ Error al enviar DJ: {e}")

    # ------------------------------------------------------------------
    async def leer_saldo_cierre(self) -> tuple[float, float]:
        """
        Lee el resultado del cierre de DJ.
        Retorna (saldo_a_pagar, saldo_a_favor):
          - saldo_a_pagar > 0 → hay deuda
          - saldo_a_favor > 0 → hay saldo a favor del contribuyente
        """
        texto = await self.page.inner_text("body")

        saldo_a_pagar = 0.0
        saldo_a_favor = 0.0

        # Buscar "Saldo a Pagar" o "A Pagar"
        match_pagar = re.search(
            r"(?:Saldo\s+a\s+pagar|A\s+pagar|Impuesto\s+a\s+pagar)[^\d]*\$?\s*([\d.,]+)",
            texto, re.IGNORECASE
        )
        if match_pagar:
            saldo_a_pagar = float(match_pagar.group(1).replace(".", "").replace(",", "."))

        # Buscar "Saldo a Favor" negativo en la tabla (puede venir como importe negativo)
        match_favor = re.search(
            r"(?:Saldo\s+a\s+favor|Saldo\s+acumulado|Favor\s+del\s+contribuyente)[^\d\-]*"
            r"(-?\s*[\d.,]+)",
            texto, re.IGNORECASE
        )
        if match_favor:
            val = float(match_favor.group(1).replace(" ", "").replace(".", "").replace(",", "."))
            saldo_a_favor = abs(val)

        self.log(f"   📊 Cierre DJ: A pagar ${saldo_a_pagar:,.2f} | A favor ${saldo_a_favor:,.2f}")
        return saldo_a_pagar, saldo_a_favor

    # ------------------------------------------------------------------
    async def generar_vep_arba(self, importe: float, medio_pago: str = "qr") -> dict:
        """
        Genera el VEP de IIBB ARBA cuando hay saldo a pagar.
        Intenta desde el portal ARBA → Liquidaciones → VEP.
        Retorna dict con numero_vep y pdf_path.
        """
        self.log(f"💳 Generando VEP ARBA — Importe: $ {importe:,.2f}...")
        vep_info = {"numero_vep": None, "pdf_path": None}
        try:
            # Desde la pantalla actual de ARBA (post-cierre DJ), buscar link de pago/VEP
            try:
                await self.page.click(
                    "a:has-text('Pagar'), button:has-text('Pagar'), "
                    "a:has-text('VEP'), button:has-text('VEP'), "
                    "a:has-text('Generar VEP'), a:has-text('Volante')",
                    timeout=8000
                )
                await self.page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                # Navegar a sección de liquidaciones ARBA
                try:
                    await self.page.click("text=Liquidaciones", timeout=5000)
                    await self.page.wait_for_load_state("networkidle", timeout=10000)
                    await self.page.click(
                        "a:has-text('VEP'), a:has-text('Pago'), button:has-text('Generar')",
                        timeout=5000
                    )
                    await self.page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    pass

            # Seleccionar medio de pago
            await _seleccionar_medio_pago(self.page, medio_pago, self.log)
            await self.page.wait_for_timeout(500)

            # Confirmar / Enviar VEP
            try:
                await self.page.click(
                    "button:has-text('Aceptar'), button:has-text('Enviar'), "
                    "button:has-text('Confirmar'), input[value='Enviar']",
                    timeout=8000
                )
                await self.page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass

            # Leer número de VEP
            texto = await self.page.inner_text("body")
            match = re.search(r"N[°º]\s*(?:de\s+)?VEP[:\s]*([\d]+)", texto, re.IGNORECASE)
            if not match:
                match = re.search(r"VEP[:\s#Nº]*\s*([\d]{6,})", texto, re.IGNORECASE)
            if match:
                vep_info["numero_vep"] = match.group(1)
                self.log(f"   VEP ARBA N°: {vep_info['numero_vep']}")

            # Descargar PDF
            try:
                async with self.page.expect_download(timeout=20000) as dl_info:
                    await self.page.click(
                        "button:has-text('Descargar'), a:has-text('Descargar'), "
                        "a:has-text('PDF'), button:has-text('PDF')",
                        timeout=10000
                    )
                download = await dl_info.value
                fname = download.suggested_filename or "VEP_IIBB_ARBA.pdf"
                path = os.path.join(self.download_dir, fname)
                await download.save_as(path)
                vep_info["pdf_path"] = path
                self.log(f"✅ VEP ARBA descargado: {os.path.basename(path)}")
            except Exception as e:
                self.log(f"   ⚠️ No se pudo descargar VEP ARBA: {e}")

        except Exception as e:
            self.log(f"   ⚠️ Error generando VEP ARBA: {e}")

        return vep_info

    # ------------------------------------------------------------------
    async def descargar_comprobante(self) -> str | None:
        """Descarga el PDF comprobante R-606M de la DJ enviada."""
        self.log("📄 Descargando comprobante R-606M...")
        try:
            async with self.page.expect_download(timeout=20000) as dl_info:
                await self.page.click(
                    "a:has-text('PDF'), a:has-text('Comprobante'), "
                    "a:has-text('R-606'), a:has-text('Imprimir'), button:has-text('PDF')"
                )
            download = await dl_info.value
            path = os.path.join(self.download_dir, download.suggested_filename or "IIBB_R606M.pdf")
            await download.save_as(path)
            self.log(f"✅ Comprobante guardado: {os.path.basename(path)}")
            return path
        except Exception as e:
            self.log(f"⚠️ No se pudo descargar comprobante: {e}")
            return None

    # ------------------------------------------------------------------
    async def liquidar(self, anio: int, mes: int, actividades: list,
                       saldo_favor_anterior: float = 0.0,
                       medio_pago: str = "qr") -> dict:
        """
        Flujo completo IIBB Local con bifurcación saldo a favor / saldo a pagar.

        Retorna dict con:
          - resultado: "saldo_a_favor" | "saldo_a_pagar" | "error"
          - comprobante: path al PDF R-606M
          - a_pagar: float (0 si hay saldo a favor)
          - a_favor: float (0 si hay saldo a pagar)
          - vep: dict con numero_vep y pdf_path (solo si saldo a pagar)
        """
        resultado = {
            "resultado": "ok",
            "comprobante": None,
            "a_pagar": 0.0,
            "a_favor": 0.0,
            "vep": None,
            "error": None,
        }
        try:
            await self.start(headless=HEADLESS)
            await self.login_arba(self.cuit, self.password)
            await self.navegar_a_iibb()
            await self.iniciar_dj(anio, mes)
            await self.cargar_actividades(actividades)
            await self.cerrar_dj(saldo_favor_anterior)

            # Leer resultado del cierre
            saldo_a_pagar, saldo_a_favor = await self.leer_saldo_cierre()
            resultado["a_pagar"] = saldo_a_pagar
            resultado["a_favor"] = saldo_a_favor

            if saldo_a_pagar > 0:
                # CASO B: hay deuda → generar VEP
                resultado["resultado"] = "saldo_a_pagar"
                self.log(f"💰 Saldo a pagar detectado: $ {saldo_a_pagar:,.2f}")
                vep = await self.generar_vep_arba(saldo_a_pagar, medio_pago)
                resultado["vep"] = vep
            else:
                # CASO A: saldo a favor → solo descargar comprobante
                resultado["resultado"] = "saldo_a_favor"
                self.log(f"✅ Saldo a favor del contribuyente: $ {saldo_a_favor:,.2f}")

            pdf = await self.descargar_comprobante()
            resultado["comprobante"] = pdf

        except Exception as e:
            resultado["resultado"] = "error"
            resultado["error"] = str(e)
            self.log(f"❌ Error en liquidación IIBB: {e}")
        finally:
            await self.close()
        return resultado


# ═══════════════════════════════════════════════════════════════════════════════
# CONVENIO MULTILATERAL — COMARB / SIFERE
# ═══════════════════════════════════════════════════════════════════════════════

class COMLiquidador(_BaseScraper):
    """
    Presenta DJ CM03 en SIFERE (sifereweb.comarb.gov.ar) y genera el VEP.

    El acceso a SIFERE requiere Clave Fiscal ARCA.
    El usuario debe tener delegado el servicio "SIFERE WEB - Declaraciones Juradas"
    en su portal de Clave Fiscal.

    Jurisdicciones estándar usadas:
        901 = CABA
        902 = Buenos Aires
    """

    ARCA_LOGIN_URL = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"
    SIFERE_URL     = "https://sifereweb.comarb.gov.ar/sifereweb/"
    PORTAL_ARCA    = "https://portalcf.cloud.afip.gob.ar/portal/app/"

    JURISDICCIONES = {
        "caba"  : "901",
        "bsas"  : "902",
    }

    # ------------------------------------------------------------------
    async def login_arca(self):
        """Login en ARCA con Clave Fiscal."""
        self.log("🔐 Iniciando sesión en ARCA (Clave Fiscal)...")
        await self.page.goto(self.ARCA_LOGIN_URL, wait_until="networkidle", timeout=30000)

        # CUIT
        await self.page.wait_for_selector("#F1\\:username, input[name='username']", timeout=15000)
        cuit_fmt = f"{self.cuit[:2]}-{self.cuit[2:10]}-{self.cuit[10:]}" if len(self.cuit) == 11 else self.cuit
        try:
            await self.page.fill("#F1\\:username", cuit_fmt)
        except Exception:
            await self.page.fill("input[name='username']", cuit_fmt)

        await self.page.click("#F1\\:btnSiguiente, button:has-text('Siguiente'), button:has-text('siguiente')")
        await self.page.wait_for_load_state("networkidle", timeout=15000)

        # Contraseña
        await self.page.wait_for_selector("#F1\\:password, input[type='password']", timeout=15000)
        try:
            await self.page.fill("#F1\\:password", self.password)
        except Exception:
            await self.page.fill("input[type='password']", self.password)

        await self.page.click("#F1\\:btnIngresar, button:has-text('Ingresar'), button[type='submit']")
        await self.page.wait_for_load_state("networkidle", timeout=30000)
        self.log("✅ Login ARCA exitoso")

    # ------------------------------------------------------------------
    async def navegar_sifere(self):
        """Navega al portal SIFERE desde ARCA."""
        self.log("🌐 Navegando a SIFERE...")
        try:
            # Ir directo a SIFERE (la sesión ARCA se comparte por cookie)
            await self.page.goto(self.SIFERE_URL, wait_until="networkidle", timeout=30000)

            # Si redirige al portal ARCA para elegir representado/perfil, manejarlo
            if "afip.gob.ar" in self.page.url or "portalcf" in self.page.url:
                self.log("   Seleccionando acceso a SIFERE desde portal ARCA...")
                try:
                    await self.page.click("text=SIFERE, text=Convenio Multilateral", timeout=10000)
                    await self.page.wait_for_load_state("networkidle", timeout=20000)
                except Exception:
                    # Intentar navegar por URL directa del servicio
                    await self.page.goto(
                        "https://portalcf.cloud.afip.gob.ar/portal/app/link/sifereweb",
                        wait_until="networkidle", timeout=20000
                    )
                    await self.page.wait_for_load_state("networkidle", timeout=20000)
                    await self.page.goto(self.SIFERE_URL, wait_until="networkidle", timeout=20000)
        except Exception as e:
            self.log(f"   ⚠️ Problema navegando a SIFERE: {e}")

        self.log("✅ En SIFERE")

    # ------------------------------------------------------------------
    async def abrir_dj_periodo(self, anio: int, mes: int):
        """Abre o crea la DJ CM03 para el período indicado."""
        periodo = f"{anio}{mes:02d}"
        self.log(f"📝 Abriendo DJ CM03 período {periodo}...")

        try:
            # Ir a Declaraciones Juradas Mensuales
            await self.page.click("text=Declaraciones Juradas Mensuales", timeout=10000)
            await self.page.wait_for_load_state("networkidle", timeout=15000)

            # Buscar el período en la lista o crear nueva
            try:
                # Buscar una DJ existente para el período
                row = await self.page.query_selector(f"td:has-text('{periodo}')")
                if row:
                    await row.click()
                    await self.page.wait_for_load_state("networkidle", timeout=15000)
                    self.log(f"✅ DJ existente {periodo} abierta")
                    return
            except Exception:
                pass

            # Crear nueva DJ
            await self.page.click("button:has-text('Nueva'), a:has-text('Nueva'), button:has-text('Agregar')")
            await self.page.wait_for_load_state("networkidle", timeout=15000)

            # Seleccionar período
            try:
                await self.page.select_option("select[name*='periodo'], select[id*='periodo']", periodo)
            except Exception:
                await self.page.fill("input[name*='periodo'], input[id*='periodo']", periodo)

            await self.page.click("button:has-text('Aceptar'), button:has-text('Confirmar'), button[type='submit']")
            await self.page.wait_for_load_state("networkidle", timeout=20000)
            self.log(f"✅ DJ {periodo} creada")

        except Exception as e:
            self.log(f"⚠️ Error abriendo DJ: {e}")

    # ------------------------------------------------------------------
    async def cargar_actividades_nivel_pais(self, base_caba: float, base_bsas: float):
        """
        Carga la base imponible por actividad a nivel país.
        La distribución por jurisdicción se hace en el paso siguiente.
        """
        total_base = base_caba + base_bsas
        self.log(f"📊 Cargando base imponible total nivel país: $ {total_base:,.2f}")

        try:
            # Ir a Datos de Actividades a Nivel País
            await self.page.click("text=Datos de Actividades", timeout=10000)
            await self.page.wait_for_load_state("networkidle", timeout=15000)

            # Para cada actividad en la lista, editar e ingresar el monto
            edit_btns = await self.page.query_selector_all(
                "a.edit-btn, a[title='Editar'], button:has-text('Editar'), "
                "td a svg, .edit-icon, a.editar"
            )

            if not edit_btns:
                # Buscar ícono de lápiz / edición
                edit_btns = await self.page.query_selector_all("td:first-child a, td a[href*='edit']")

            for i, btn in enumerate(edit_btns):
                try:
                    await btn.click()
                    await self.page.wait_for_load_state("networkidle", timeout=10000)

                    campo = await self.page.query_selector(
                        "input[name*='monto'], input[name*='base'], "
                        "input[id*='monto'], input[id*='base'], "
                        "input[name*='imponible']"
                    )
                    if campo:
                        await campo.triple_click()
                        # Si es la primera actividad, usar el total
                        await campo.fill(str(total_base).replace(".", ","))

                    await self.page.click("button:has-text('Guardar'), button:has-text('Aceptar')")
                    await self.page.wait_for_load_state("networkidle", timeout=10000)
                    self.log(f"   ✅ Actividad {i+1} nivel país cargada")
                    break  # Solo la primera actividad recibe el total
                except Exception as e:
                    self.log(f"   ⚠️ Actividad {i+1}: {e}")

        except Exception as e:
            self.log(f"⚠️ Error cargando actividades nivel país: {e}")

    # ------------------------------------------------------------------
    async def cargar_jurisdicciones(self, base_caba: float, base_bsas: float):
        """
        Carga la base imponible por jurisdicción (CABA y Buenos Aires).
        """
        self.log(f"🗺️ Cargando jurisdicciones — CABA: ${base_caba:,.2f} | Bs.As: ${base_bsas:,.2f}")

        jurisdicciones_data = [
            ("901", "CABA", base_caba),
            ("902", "Buenos Aires", base_bsas),
        ]

        try:
            # Ir a Actividades por Jurisdicción
            await self.page.click("text=Actividades por Jurisdicción, text=Jurisdicción", timeout=10000)
            await self.page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            # Intentar desde el árbol lateral
            try:
                await self.page.click("text=901, text=CABA", timeout=5000)
            except Exception:
                pass

        for cod, nombre, base in jurisdicciones_data:
            self.log(f"   Jurisdicción {cod}/{nombre}: $ {base:,.2f}")
            try:
                # Expandir o navegar a la jurisdicción
                try:
                    await self.page.click(f"text={cod}, text={nombre}", timeout=5000)
                    await self.page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    pass

                # Buscar y editar actividades en esa jurisdicción
                edit_btns = await self.page.query_selector_all(
                    "a[title*='Editar'], button:has-text('Editar'), td a"
                )
                for btn in edit_btns:
                    try:
                        await btn.click()
                        await self.page.wait_for_load_state("networkidle", timeout=10000)

                        campo = await self.page.query_selector(
                            "input[name*='base'], input[id*='base'], "
                            "input[name*='imponible'], input[name*='monto']"
                        )
                        if campo:
                            await campo.triple_click()
                            await campo.fill(str(base).replace(".", ","))

                        await self.page.click(
                            "button:has-text('Guardar'), button:has-text('Aceptar'), "
                            "button:has-text('Guardar Cambios')"
                        )
                        await self.page.wait_for_load_state("networkidle", timeout=10000)
                        self.log(f"   ✅ {nombre} cargada")
                        break
                    except Exception as e:
                        self.log(f"   ⚠️ {nombre}: {e}")

                # Actualizar total de ingresos (A+B+C)
                try:
                    await self.page.fill(
                        "input[name*='gravados'], input[id*='gravados']",
                        str(base).replace(".", ",")
                    )
                    await self.page.click(
                        "button:has-text('Actualizar'), button:has-text('Recalcular')"
                    )
                    await self.page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    pass

            except Exception as e:
                self.log(f"   ⚠️ Error en jurisdicción {nombre}: {e}")

    # ------------------------------------------------------------------
    async def finalizar_dj(self):
        """Navega a Liquidación Final y finaliza la DDJJ."""
        self.log("📤 Finalizando DDJJ...")
        try:
            await self.page.click("text=Liquidación Final, text=Finalizar DDJJ", timeout=10000)
            await self.page.wait_for_load_state("networkidle", timeout=15000)

            await self.page.click(
                "a:has-text('Finalizar DDJJ'), button:has-text('Finalizar'), "
                "a:has-text('Finalizar')"
            )
            await self.page.wait_for_load_state("networkidle", timeout=20000)
            self.log("✅ DDJJ finalizada")
        except Exception as e:
            self.log(f"⚠️ Error al finalizar DDJJ: {e}")

    # ------------------------------------------------------------------
    async def generar_y_descargar_vep(self) -> str | None:
        """Genera el VEP y descarga el PDF."""
        self.log("💳 Generando VEP...")
        vep_path = None
        try:
            # Navegar a sección de pago
            try:
                await self.page.click(
                    "text=Volantes de Pagos, text=Pago, text=VEP",
                    timeout=10000
                )
                await self.page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass

            # Seleccionar red de pago (Link por defecto)
            try:
                await self.page.click(
                    "label:has-text('Pagar-Red Link'), input[value*='LINK'], "
                    "label:has-text('Link')"
                )
            except Exception:
                pass

            # Click en "Enviar VEP"
            async with self.page.expect_download(timeout=30000) as dl_info:
                await self.page.click(
                    "button:has-text('Enviar VEP'), a:has-text('Enviar VEP'), "
                    "button:has-text('Generar VEP')"
                )
            download = await dl_info.value
            fname = download.suggested_filename or "VEP_COM.pdf"
            vep_path = os.path.join(self.download_dir, fname)
            await download.save_as(vep_path)
            self.log(f"✅ VEP descargado: {os.path.basename(vep_path)}")
        except Exception as e:
            self.log(f"⚠️ No se pudo descargar VEP: {e}")
        return vep_path

    # ------------------------------------------------------------------
    async def liquidar(self, anio: int, mes: int,
                       base_caba: float, base_bsas: float) -> dict:
        """Flujo completo Convenio Multilateral."""
        resultado = {"resultado": "ok", "vep": None, "total_a_pagar": 0.0, "error": None}
        try:
            await self.start(headless=HEADLESS)
            await self.login_arca()
            await self.navegar_sifere()
            await self.abrir_dj_periodo(anio, mes)
            await self.cargar_actividades_nivel_pais(base_caba, base_bsas)
            await self.cargar_jurisdicciones(base_caba, base_bsas)
            await self.finalizar_dj()

            # Leer total a pagar
            try:
                texto = await self.page.inner_text("body")
                match = re.search(r"Total\s+(?:General\s+)?a\s+Pagar[^\d]*\$?\s*([\d.,]+)", texto, re.IGNORECASE)
                if match:
                    resultado["total_a_pagar"] = float(
                        match.group(1).replace(".", "").replace(",", ".")
                    )
            except Exception:
                pass

            vep = await self.generar_y_descargar_vep()
            resultado["vep"] = vep
        except Exception as e:
            resultado["resultado"] = "error"
            resultado["error"] = str(e)
            self.log(f"❌ Error en liquidación COM: {e}")
        finally:
            await self.close()
        return resultado


# ═══════════════════════════════════════════════════════════════════════════════
# IVA — ARCA
# ═══════════════════════════════════════════════════════════════════════════════

class IVALiquidador(_BaseScraper):
    """
    Liquida la posición mensual de IVA en ARCA (ex-AFIP) y genera el VEP.

    Parámetros de entrada:
        anio              : int   — Año del período
        mes               : int   — Mes del período
        debito_fiscal     : float — Débito fiscal del período
        credito_fiscal    : float — Crédito fiscal del período
        retenciones       : float — Retenciones SICORE y otras
        saldo_favor_1p    : float — Saldo técnico a favor período anterior
        saldo_favor_2p    : float — Saldo libre disponibilidad período anterior
    """

    ARCA_LOGIN_URL = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"
    PORTAL_URL     = "https://portalcf.cloud.afip.gob.ar/portal/app/"

    # ------------------------------------------------------------------
    async def login_arca(self):
        """Login en ARCA con Clave Fiscal."""
        self.log("🔐 Iniciando sesión en ARCA...")
        await self.page.goto(self.ARCA_LOGIN_URL, wait_until="networkidle", timeout=30000)

        cuit_fmt = f"{self.cuit[:2]}-{self.cuit[2:10]}-{self.cuit[10:]}" if len(self.cuit) == 11 else self.cuit
        await self.page.wait_for_selector("#F1\\:username, input[name='username']", timeout=15000)
        try:
            await self.page.fill("#F1\\:username", cuit_fmt)
        except Exception:
            await self.page.fill("input[name='username']", cuit_fmt)

        await self.page.click("#F1\\:btnSiguiente, button:has-text('Siguiente')")
        await self.page.wait_for_load_state("networkidle", timeout=15000)

        await self.page.wait_for_selector("input[type='password']", timeout=15000)
        await self.page.fill("input[type='password']", self.password)
        await self.page.click("button:has-text('Ingresar'), button[type='submit']")
        await self.page.wait_for_load_state("networkidle", timeout=30000)
        self.log("✅ Login ARCA exitoso")

    # ------------------------------------------------------------------
    async def navegar_a_iva(self):
        """Navega al servicio de Declaración Jurada de IVA."""
        self.log("🧾 Navegando al módulo IVA...")
        try:
            # Desde el portal ARCA, buscar el servicio IVA
            await self.page.goto(self.PORTAL_URL, wait_until="networkidle", timeout=20000)
            await self.page.click("text=IVA", timeout=10000)
            await self.page.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            try:
                # URL directa del servicio IVA
                await self.page.goto(
                    "https://w1.afip.gob.ar/afip/ddjjIVA/",
                    wait_until="networkidle", timeout=20000
                )
            except Exception as e:
                self.log(f"⚠️ Error navegando a IVA: {e}")
        self.log("✅ En módulo IVA")

    # ------------------------------------------------------------------
    async def crear_dj_periodo(self, anio: int, mes: int):
        """Crea o abre la DJ IVA del período."""
        periodo = f"{anio}{mes:02d}"
        self.log(f"📝 Creando/abriendo DJ IVA {mes:02d}/{anio}...")
        try:
            # Buscar si ya existe una DJ para el período
            try:
                row = await self.page.query_selector(
                    f"td:has-text('{periodo}'), tr:has-text('{mes:02d}/{anio}')"
                )
                if row:
                    await row.click()
                    await self.page.wait_for_load_state("networkidle", timeout=15000)
                    self.log("✅ DJ existente abierta")
                    return
            except Exception:
                pass

            # Crear nueva
            await self.page.click(
                "button:has-text('Nueva'), a:has-text('Nueva declaración'), "
                "a:has-text('Agregar'), button:has-text('Agregar')"
            )
            await self.page.wait_for_load_state("networkidle", timeout=15000)

            # Seleccionar período
            try:
                await self.page.select_option(
                    "select[name*='periodo'], select[id*='periodo']", periodo
                )
            except Exception:
                try:
                    await self.page.select_option(
                        "select[name*='mes'], select[id*='mes']", str(mes)
                    )
                    await self.page.select_option(
                        "select[name*='anio'], select[id*='anio']", str(anio)
                    )
                except Exception:
                    pass

            await self.page.click("button:has-text('Continuar'), button[type='submit']")
            await self.page.wait_for_load_state("networkidle", timeout=20000)
            self.log(f"✅ DJ IVA {periodo} creada")
        except Exception as e:
            self.log(f"⚠️ Error creando DJ: {e}")

    # ------------------------------------------------------------------
    async def completar_dj(self,
                           # VENTAS a Consumidor Final
                           vta_cf_neto_21: float = 0.0,
                           vta_cf_iva_21:  float = 0.0,
                           vta_cf_neto_105: float = 0.0,
                           vta_cf_iva_105:  float = 0.0,
                           # VENTAS a Responsable Inscripto
                           vta_ri_neto_21: float = 0.0,
                           vta_ri_iva_21:  float = 0.0,
                           vta_ri_neto_105: float = 0.0,
                           vta_ri_iva_105:  float = 0.0,
                           # COMPRAS
                           cmp_neto_21: float = 0.0,
                           cmp_iva_21:  float = 0.0,
                           cmp_neto_105: float = 0.0,
                           cmp_iva_105:  float = 0.0,
                           # Retenciones / saldos anteriores
                           retenciones: float = 0.0,
                           saldo_favor_1p: float = 0.0,
                           saldo_favor_2p: float = 0.0):
        """
        Completa los campos detallados de la DJ IVA en ARCA.

        ARCA divide la pantalla en varias secciones:
          1. Ventas y/u operaciones gravadas (separadas por alícuota)
          2. Compras, locaciones y prestaciones gravadas
          3. Retenciones / Percepciones sufridas
          4. Saldos a favor períodos anteriores
        """
        debito_total = vta_cf_iva_21 + vta_cf_iva_105 + vta_ri_iva_21 + vta_ri_iva_105
        credito_total = cmp_iva_21 + cmp_iva_105
        self.log(
            f"📊 Completando DJ IVA — Débito: ${debito_total:,.2f} | Crédito: ${credito_total:,.2f}"
        )

        # Mapa de campos: (valor, selectores posibles en la página ARCA)
        # Los IDs reales de ARCA son estables entre versiones (F2002)
        campos_arca = [
            # Ventas CF 21%
            (vta_cf_neto_21,  ["neto21CF", "netoGravado21CF", "ventasCF21neto",
                                "input[name*='vcf'][name*='21'][name*='neto']"]),
            (vta_cf_iva_21,   ["iva21CF", "ivaGravado21CF",
                                "input[name*='vcf'][name*='21'][name*='iva']"]),
            # Ventas CF 10.5%
            (vta_cf_neto_105, ["neto105CF", "netoGravado105CF",
                                "input[name*='vcf'][name*='105'][name*='neto']"]),
            (vta_cf_iva_105,  ["iva105CF",
                                "input[name*='vcf'][name*='105'][name*='iva']"]),
            # Ventas RI 21%
            (vta_ri_neto_21,  ["neto21RI", "netoGravado21RI",
                                "input[name*='vri'][name*='21'][name*='neto']"]),
            (vta_ri_iva_21,   ["iva21RI",
                                "input[name*='vri'][name*='21'][name*='iva']"]),
            # Ventas RI 10.5%
            (vta_ri_neto_105, ["neto105RI",
                                "input[name*='vri'][name*='105'][name*='neto']"]),
            (vta_ri_iva_105,  ["iva105RI",
                                "input[name*='vri'][name*='105'][name*='iva']"]),
            # Compras 21%
            (cmp_neto_21,     ["netoCompra21", "compras21neto",
                                "input[name*='cmp'][name*='21'][name*='neto']"]),
            (cmp_iva_21,      ["ivaCompra21",
                                "input[name*='cmp'][name*='21'][name*='iva']"]),
            # Compras 10.5%
            (cmp_neto_105,    ["netoCompra105",
                                "input[name*='cmp'][name*='105'][name*='neto']"]),
            (cmp_iva_105,     ["ivaCompra105",
                                "input[name*='cmp'][name*='105'][name*='iva']"]),
            # Retenciones
            (retenciones,     ["retenciones", "retencion", "sicore",
                                "input[name*='ret']"]),
            # Saldos anteriores
            (saldo_favor_1p,  ["saldo1p", "saldoFavor1", "primerParrafo",
                                "input[name*='saldo'][name*='1']"]),
            (saldo_favor_2p,  ["saldo2p", "saldoFavor2", "segundoParrafo",
                                "input[name*='saldo'][name*='2']"]),
        ]

        for valor, selectores in campos_arca:
            if valor == 0:
                continue
            for sel in selectores:
                try:
                    # Si el selector parece un atributo CSS úsalo directo;
                    # si no, búscalo por id o name
                    if "[" in sel:
                        campo = await self.page.query_selector(sel)
                    else:
                        campo = await self.page.query_selector(
                            f"#{sel}, [name='{sel}'], [id*='{sel}'], [name*='{sel}']"
                        )
                    if campo:
                        await campo.triple_click()
                        await campo.fill(str(valor).replace(".", ","))
                        break
                except Exception:
                    continue

        # Intentar recalcular
        try:
            await self.page.click(
                "button:has-text('Calcular'), button:has-text('Recalcular'), "
                "button:has-text('Actualizar'), a:has-text('Calcular')"
            )
            await self.page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass

    # ------------------------------------------------------------------
    async def enviar_dj(self):
        """Envía/presenta la DJ."""
        self.log("📤 Presentando DJ IVA...")
        try:
            await self.page.click(
                "button:has-text('Presentar'), button:has-text('Enviar'), "
                "a:has-text('Presentar'), input[value*='Presentar']"
            )
            await self.page.wait_for_load_state("networkidle", timeout=30000)

            # Confirmar si aparece diálogo
            try:
                await self.page.click(
                    "button:has-text('Confirmar'), button:has-text('Aceptar'), "
                    "button:has-text('Sí')"
                )
                await self.page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass

            self.log("✅ DJ IVA presentada")
        except Exception as e:
            self.log(f"⚠️ Error al presentar DJ: {e}")

    # ------------------------------------------------------------------
    async def leer_posicion_f2051(self) -> tuple[float, float]:
        """
        Lee la posición IVA del F.2051 o de la página de vista previa.
        Retorna (importe_a_pagar, saldo_a_favor):
          - importe_a_pagar > 0 → hay deuda
          - saldo_a_favor > 0   → hay saldo a favor del contribuyente
        """
        try:
            texto = await self.page.inner_text("body")
        except Exception:
            return 0.0, 0.0

        importe_a_pagar = 0.0
        saldo_a_favor = 0.0

        # Detectar "Importe a ingresar" o "Saldo a pagar"
        match_pagar = re.search(
            r"(?:Importe\s+a\s+ingresar|Saldo\s+a\s+pagar|A\s+pagar|Impuesto\s+a\s+pagar)"
            r"[^\d]*\$?\s*([\d.,]+)",
            texto, re.IGNORECASE
        )
        if match_pagar:
            importe_a_pagar = float(match_pagar.group(1).replace(".", "").replace(",", "."))

        # Detectar "Saldo técnico a favor del contribuyente" (F.2051)
        match_favor = re.search(
            r"(?:Saldo\s+t[eé]cnico\s+a\s+favor|Saldo\s+a\s+favor\s+del\s+contribuyente|"
            r"A\s+favor\s+del\s+contribuyente)[^\d\-]*(-?\s*[\d.,]+)",
            texto, re.IGNORECASE
        )
        if match_favor:
            val_str = match_favor.group(1).replace(" ", "").replace(".", "").replace(",", ".")
            saldo_a_favor = abs(float(val_str))

        self.log(
            f"   📊 Posición IVA: "
            f"A pagar ${importe_a_pagar:,.2f} | A favor ${saldo_a_favor:,.2f}"
        )
        return importe_a_pagar, saldo_a_favor

    # ------------------------------------------------------------------
    async def descargar_f2051(self) -> str | None:
        """Descarga el PDF F.2051 si está disponible."""
        self.log("📄 Descargando F.2051...")
        try:
            async with self.page.expect_download(timeout=20000) as dl_info:
                await self.page.click(
                    "a:has-text('F.2051'), button:has-text('F.2051'), "
                    "a:has-text('Descargar'), a:has-text('PDF'), "
                    "button:has-text('Descargar formulario')",
                    timeout=10000
                )
            download = await dl_info.value
            fname = download.suggested_filename or "F2051_IVA.pdf"
            path = os.path.join(self.download_dir, fname)
            await download.save_as(path)
            self.log(f"✅ F.2051 guardado: {os.path.basename(path)}")
            return path
        except Exception as e:
            self.log(f"   ⚠️ No se pudo descargar F.2051: {e}")
            return None

    # ------------------------------------------------------------------
    async def generar_vep_seti(self, mes: int, anio: int,
                                importe: float,
                                medio_pago: str = "qr") -> dict:
        """
        Genera el VEP de IVA en ARCA SETI cuando hay saldo a pagar.
        Navega al Nuevo VEP con grupo/tipo IVA y selecciona el medio de pago.
        """
        self.log(f"💳 Generando VEP IVA en SETI — Importe: $ {importe:,.2f}...")
        vep_info = {"numero_vep": None, "pdf_path": None}

        SETI_NUEVO_VEP = "https://seti.afip.gob.ar/setiweb/#/pago/nuevo-vep?op=1"
        try:
            await self.page.goto(SETI_NUEVO_VEP, wait_until="networkidle", timeout=30000)
            await self.page.wait_for_timeout(2000)

            # Cerrar popup si hay
            try:
                await self.page.click(
                    "button:has-text('Entendido'), button:has-text('Cerrar')",
                    timeout=4000
                )
            except Exception:
                pass

            # Organismo → ARCA
            try:
                await self.page.select_option(
                    "select[formcontrolname*='organismo'], select[id*='organismo']",
                    label="ARCA", timeout=5000
                )
            except Exception:
                pass

            # Grupo → "IVA" o "Impuesto al Valor Agregado"
            for lbl in ["IVA", "Impuesto al Valor Agregado", "Impuesto al valor agregado"]:
                try:
                    await self.page.select_option(
                        "select[formcontrolname*='grupo'], select[id*='grupo']",
                        label=lbl, timeout=5000
                    )
                    await self.page.wait_for_timeout(1000)
                    break
                except Exception:
                    continue

            # Tipo → "IVA - Impuesto al Valor Agregado" o similar
            for lbl in ["IVA - Impuesto al Valor Agregado", "IVA General", "IVA"]:
                try:
                    await self.page.select_option(
                        "select[formcontrolname*='tipo'], select[id*='tipo']",
                        label=lbl, timeout=5000
                    )
                    break
                except Exception:
                    continue

            await self.page.click(
                "button:has-text('Siguiente'), input[value='Siguiente']"
            )
            await self.page.wait_for_load_state("networkidle", timeout=20000)
            await self.page.wait_for_timeout(2000)

            # Período
            try:
                await self.page.select_option(
                    "select[formcontrolname*='mes'], select[id*='mes']",
                    value=str(mes), timeout=5000
                )
            except Exception:
                pass
            try:
                await self.page.select_option(
                    "select[formcontrolname*='anio'], select[id*='anio'], "
                    "select[formcontrolname*='year']",
                    value=str(anio), timeout=5000
                )
            except Exception:
                pass

            # Importe
            try:
                campo_imp = await self.page.query_selector(
                    "input[formcontrolname*='importe'], input[id*='importe'], "
                    "input[formcontrolname*='monto']"
                )
                if campo_imp:
                    await campo_imp.triple_click()
                    await campo_imp.fill(str(importe).replace(".", ","))
            except Exception:
                pass

            await self.page.click("button:has-text('Siguiente'), input[value='Siguiente']")
            await self.page.wait_for_load_state("networkidle", timeout=20000)
            await self.page.wait_for_timeout(1500)

            # Marcar checkbox del VEP
            try:
                checkbox = await self.page.query_selector("input[type='checkbox']")
                if checkbox:
                    await checkbox.check()
                    await self.page.wait_for_timeout(500)
            except Exception:
                pass

            # Seleccioná medio de pago
            await self.page.click(
                "button:has-text('Seleccioná medio de pago'), "
                "button:has-text('Seleccionar medio'), a:has-text('medio de pago')",
                timeout=10000
            )
            await self.page.wait_for_load_state("networkidle", timeout=15000)
            await self.page.wait_for_timeout(1000)

            await _seleccionar_medio_pago(self.page, medio_pago, self.log)
            await self.page.wait_for_timeout(500)

            await self.page.click(
                "button:has-text('Aceptar'), button:has-text('Confirmar')",
                timeout=8000
            )
            await self.page.wait_for_load_state("networkidle", timeout=20000)
            await self.page.wait_for_timeout(2000)

            # Leer número de VEP
            texto = await self.page.inner_text("body")
            match = re.search(r"N[°º]\s*(?:de\s+)?VEP[:\s]*([\d]+)", texto, re.IGNORECASE)
            if not match:
                match = re.search(r"VEP[:\s#Nº]*\s*([\d]{6,})", texto, re.IGNORECASE)
            if match:
                vep_info["numero_vep"] = match.group(1)
                self.log(f"   VEP IVA N°: {vep_info['numero_vep']}")

            # Descargar PDF VEP
            try:
                async with self.page.expect_download(timeout=20000) as dl_info:
                    await self.page.click(
                        "button:has-text('Descargar VEP'), a:has-text('Descargar VEP'), "
                        "button:has-text('Descargar'), a:has-text('Descargar')",
                        timeout=10000
                    )
                download = await dl_info.value
                fname = download.suggested_filename or "VEP_IVA.pdf"
                path = os.path.join(self.download_dir, fname)
                await download.save_as(path)
                vep_info["pdf_path"] = path
                self.log(f"✅ VEP IVA descargado: {os.path.basename(path)}")
            except Exception as e:
                self.log(f"   ⚠️ No se pudo descargar VEP IVA: {e}")

        except Exception as e:
            self.log(f"   ⚠️ Error generando VEP IVA SETI: {e}")

        return vep_info

    # ------------------------------------------------------------------
    async def liquidar(self, anio: int, mes: int,
                       vta_cf_neto_21: float = 0.0, vta_cf_iva_21: float = 0.0,
                       vta_cf_neto_105: float = 0.0, vta_cf_iva_105: float = 0.0,
                       vta_ri_neto_21: float = 0.0, vta_ri_iva_21: float = 0.0,
                       vta_ri_neto_105: float = 0.0, vta_ri_iva_105: float = 0.0,
                       cmp_neto_21: float = 0.0, cmp_iva_21: float = 0.0,
                       cmp_neto_105: float = 0.0, cmp_iva_105: float = 0.0,
                       retenciones: float = 0.0,
                       saldo_favor_1p: float = 0.0,
                       saldo_favor_2p: float = 0.0,
                       medio_pago: str = "qr") -> dict:
        """
        Flujo completo IVA con bifurcación saldo a favor / saldo a pagar.

        Retorna dict con:
          - resultado: "saldo_a_favor" | "saldo_a_pagar" | "error"
          - posicion: float (positivo = a pagar, negativo = a favor)
          - pdf_f2051: path al PDF F.2051
          - vep: dict con numero_vep y pdf_path (solo si saldo a pagar)
        """
        resultado = {
            "resultado": "ok",
            "posicion": 0.0,
            "pdf_f2051": None,
            "vep": None,
            "error": None,
            # compatibilidad backward
            "vep_path": None,
        }
        try:
            await self.start(headless=HEADLESS)
            await self.login_arca()
            await self.navegar_a_iva()
            await self.crear_dj_periodo(anio, mes)
            await self.completar_dj(
                vta_cf_neto_21=vta_cf_neto_21, vta_cf_iva_21=vta_cf_iva_21,
                vta_cf_neto_105=vta_cf_neto_105, vta_cf_iva_105=vta_cf_iva_105,
                vta_ri_neto_21=vta_ri_neto_21, vta_ri_iva_21=vta_ri_iva_21,
                vta_ri_neto_105=vta_ri_neto_105, vta_ri_iva_105=vta_ri_iva_105,
                cmp_neto_21=cmp_neto_21, cmp_iva_21=cmp_iva_21,
                cmp_neto_105=cmp_neto_105, cmp_iva_105=cmp_iva_105,
                retenciones=retenciones,
                saldo_favor_1p=saldo_favor_1p,
                saldo_favor_2p=saldo_favor_2p,
            )

            await self.enviar_dj()

            # Leer posición desde F.2051 / página final
            importe_a_pagar, saldo_a_favor = await self.leer_posicion_f2051()
            resultado["posicion"] = importe_a_pagar if importe_a_pagar > 0 else -saldo_a_favor

            # Descargar F.2051
            pdf_f2051 = await self.descargar_f2051()
            resultado["pdf_f2051"] = pdf_f2051
            resultado["vep_path"] = pdf_f2051  # backward compat

            if importe_a_pagar > 0:
                # CASO B: saldo a pagar → generar VEP en SETI
                resultado["resultado"] = "saldo_a_pagar"
                self.log(f"💰 Saldo a pagar IVA: $ {importe_a_pagar:,.2f}")
                vep = await self.generar_vep_seti(mes, anio, importe_a_pagar, medio_pago)
                resultado["vep"] = vep
                if vep.get("pdf_path"):
                    resultado["vep_path"] = vep["pdf_path"]
            else:
                # CASO A: saldo a favor → solo PDF F.2051
                resultado["resultado"] = "saldo_a_favor"
                self.log(f"✅ Saldo a favor IVA: $ {saldo_a_favor:,.2f}")

        except Exception as e:
            resultado["resultado"] = "error"
            resultado["error"] = str(e)
            self.log(f"❌ Error en liquidación IVA: {e}")
        finally:
            await self.close()
        return resultado
