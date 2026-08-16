"""
Scrapers de liquidación impositiva
===================================
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
    async def descargar_comprobante(self) -> str | None:
        """Descarga el PDF comprobante de la DJ enviada."""
        self.log("📄 Descargando comprobante PDF...")
        try:
            async with self.page.expect_download(timeout=20000) as dl_info:
                await self.page.click(
                    "a:has-text('PDF'), a:has-text('Comprobante'), "
                    "a:has-text('Imprimir'), button:has-text('PDF')"
                )
            download = await dl_info.value
            path = os.path.join(self.download_dir, download.suggested_filename or "IIBB_comprobante.pdf")
            await download.save_as(path)
            self.log(f"✅ Comprobante guardado: {os.path.basename(path)}")
            return path
        except Exception as e:
            self.log(f"⚠️ No se pudo descargar comprobante: {e}")
            return None

    # ------------------------------------------------------------------
    async def liquidar(self, anio: int, mes: int, actividades: list,
                       saldo_favor_anterior: float = 0.0,
                       password_arba: str = None) -> dict:
        """
        Flujo completo IIBB Local.
        Retorna dict con 'resultado', 'comprobante', 'a_pagar'.
        """
        resultado = {"resultado": "ok", "comprobante": None, "a_pagar": 0.0, "error": None}
        try:
            await self.start(headless=HEADLESS)
            pwd = password_arba or self.password
            await self.login_arba(self.cuit, pwd)
            await self.navegar_a_iibb()
            await self.iniciar_dj(anio, mes)
            await self.cargar_actividades(actividades)
            await self.cerrar_dj(saldo_favor_anterior)

            # Intentar leer el monto a pagar de la página
            try:
                texto = await self.page.inner_text("body")
                import re as _re
                match = _re.search(r"A PAGAR[^\d]*\$?\s*([\d.,]+)", texto, _re.IGNORECASE)
                if match:
                    resultado["a_pagar"] = float(
                        match.group(1).replace(".", "").replace(",", ".")
                    )
            except Exception:
                pass

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
                import re as _re
                match = _re.search(r"Total\s+(?:General\s+)?a\s+Pagar[^\d]*\$?\s*([\d.,]+)", texto, _re.IGNORECASE)
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
    async def generar_vep(self) -> str | None:
        """Genera y descarga el VEP de IVA."""
        self.log("💳 Generando VEP IVA...")
        vep_path = None
        try:
            await self.page.click(
                "button:has-text('VEP'), a:has-text('VEP'), "
                "button:has-text('Volante'), a:has-text('Pagar')"
            )
            await self.page.wait_for_load_state("networkidle", timeout=20000)

            async with self.page.expect_download(timeout=30000) as dl_info:
                await self.page.click(
                    "button:has-text('Generar VEP'), button:has-text('Descargar'), "
                    "a:has-text('Descargar VEP'), button:has-text('Enviar VEP')"
                )
            download = await dl_info.value
            fname = download.suggested_filename or "VEP_IVA.pdf"
            vep_path = os.path.join(self.download_dir, fname)
            await download.save_as(vep_path)
            self.log(f"✅ VEP IVA descargado: {os.path.basename(vep_path)}")
        except Exception as e:
            self.log(f"⚠️ No se pudo generar VEP IVA: {e}")
        return vep_path

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
                       saldo_favor_2p: float = 0.0) -> dict:
        """Flujo completo IVA."""
        resultado = {"resultado": "ok", "vep": None, "posicion": 0.0, "error": None}
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

            # Leer posición (saldo a pagar o a favor)
            try:
                texto = await self.page.inner_text("body")
                import re as _re
                match = _re.search(r"(?:A PAGAR|SALDO)[^\d]*\$?\s*([\d.,]+)", texto, _re.IGNORECASE)
                if match:
                    resultado["posicion"] = float(
                        match.group(1).replace(".", "").replace(",", ".")
                    )
            except Exception:
                pass

            await self.enviar_dj()
            vep = await self.generar_vep()
            resultado["vep"] = vep
        except Exception as e:
            resultado["resultado"] = "error"
            resultado["error"] = str(e)
            self.log(f"❌ Error en liquidación IVA: {e}")
        finally:
            await self.close()
        return resultado
