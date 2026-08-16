"""
Scraper para ARCA (ex-AFIP)
- Login con CUIT y clave fiscal
- Mis Comprobantes: emitidos y recibidos por rango de fechas
- Mis Retenciones: SICORE 767
"""

import asyncio
import glob
import os
import re
import time as _time
import urllib.request
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError as PWTimeout


def _get_system_proxy():
    """Detecta el proxy del sistema en Windows."""
    try:
        proxies = urllib.request.getproxies()
        return proxies.get("https") or proxies.get("http") or None
    except Exception:
        return None


def _find_browser():
    """Encuentra Chrome o Edge instalado en el sistema."""
    import os
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


class ARCAScraper:
    LOGIN_URL = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"
    PORTAL_URL = "https://portalcf.cloud.afip.gob.ar/portal/app/"

    def __init__(self, cuit: str, password: str, log_fn=print):
        self.cuit = re.sub(r"[^0-9]", "", cuit)
        self.password = password
        self.log = log_fn
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    async def start(self, headless: bool = False):
        self.playwright = await async_playwright().start()
        proxy = _get_system_proxy()
        proxy_cfg = {"server": proxy} if proxy else None
        extra_args = ["--start-maximized", "--no-sandbox", "--disable-dev-shm-usage"]
        browser_path = _find_browser()

        launch_kwargs = dict(headless=headless, args=extra_args, proxy=proxy_cfg)
        if browser_path:
            self.log(f"🌐 Usando browser del sistema: {browser_path}")
            launch_kwargs["executable_path"] = browser_path

        try:
            self.browser = await self.playwright.chromium.launch(**launch_kwargs)
        except Exception:
            # Fallback sin executable_path
            launch_kwargs.pop("executable_path", None)
            try:
                self.browser = await self.playwright.chromium.launch(channel="chrome", **launch_kwargs)
            except Exception:
                self.browser = await self.playwright.chromium.launch(**launch_kwargs)
        self.context = await self.browser.new_context(
            viewport={"width": 1280, "height": 900},
            accept_downloads=True,
        )
        self.page = await self.context.new_page()
        self.page.set_default_timeout(90000)
        self.page.set_default_navigation_timeout(90000)

    async def close(self):
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

    # ------------------------------------------------------------------
    # Login ARCA
    # ------------------------------------------------------------------

    async def login(self):
        self.log("🔐 Iniciando sesión en ARCA...")
        await self.page.goto(self.LOGIN_URL, wait_until="load", timeout=90000)
        await self.page.wait_for_timeout(2000)

        # Paso 1: ingresar CUIT
        await self.page.wait_for_selector("#F1\\:username", timeout=30000)
        await self.page.fill("#F1\\:username", self.cuit)
        await self.page.click("#F1\\:btnSiguiente")
        await self.page.wait_for_load_state("load")
        await self.page.wait_for_timeout(2000)

        # Paso 2: ingresar contraseña
        try:
            await self.page.wait_for_selector("#F1\\:password", timeout=30000)
            await self.page.fill("#F1\\:password", self.password)
            await self.page.click("#F1\\:btnIngresar")
            await self.page.wait_for_load_state("load")
            await self.page.wait_for_timeout(3000)
        except PWTimeout:
            raise Exception("No apareció el campo de contraseña. Verifique el CUIT ingresado.")

        # Verificar login: esperar que la URL cambie
        await self.page.wait_for_timeout(2000)
        if "login" in self.page.url.lower():
            raise Exception("Login fallido. Verifique CUIT y clave fiscal.")

        self.log("✅ Sesión iniciada en ARCA")

    # ------------------------------------------------------------------
    # Navegar a un servicio desde el portal
    # ------------------------------------------------------------------

    async def _ir_a_servicio(self, nombre_servicio: str):
        """Busca y accede a un servicio en el portal de ARCA."""
        self.log(f"🔎 Buscando servicio: {nombre_servicio}...")
        await self.page.goto(self.PORTAL_URL, wait_until="load", timeout=90000)

        # Esperar a que aparezcan los servicios dinámicamente
        await self.page.wait_for_timeout(8000)

        # Intentar expandir "Ver todos" para ver más servicios
        try:
            ver_todos = self.page.locator("a:has-text('Ver todos'), button:has-text('Ver todos')").first
            if await ver_todos.count() > 0:
                await ver_todos.click(force=True)
                await self.page.wait_for_timeout(3000)
        except Exception:
            pass

        # Palabras clave para búsqueda parcial flexible
        keywords = nombre_servicio.lower().split()

        # 1) Intentar buscador del portal
        try:
            search = self.page.locator(
                "input[placeholder*='uscar'], input[type='search'], #buscador, input[aria-label*='uscar']"
            )
            if await search.count() > 0:
                await search.first.fill(nombre_servicio)
                await self.page.wait_for_timeout(2000)
        except Exception:
            pass

        # 2) Buscar link o botón con coincidencia parcial de texto
        async def _encontrar_y_clickear(termino: str) -> bool:
            for selector in [
                f"text={termino}",
                f"a:has-text('{termino}')",
                f"button:has-text('{termino}')",
                f"span:has-text('{termino}')",
                f"li:has-text('{termino}')",
            ]:
                try:
                    elem = self.page.locator(selector).first
                    if await elem.count() > 0:
                        self.log(f"   ✅ Encontrado '{termino}' — haciendo click...")
                        try:
                            async with self.context.expect_page(timeout=5000) as np_info:
                                await elem.click()
                            new_page = await np_info.value
                            await new_page.wait_for_load_state("load")
                            self.page = new_page
                        except Exception:
                            await elem.click()
                            await self.page.wait_for_load_state("load")
                        await self.page.wait_for_timeout(3000)
                        return True
                except Exception:
                    continue
            return False

        # Intentar nombre completo primero
        if await _encontrar_y_clickear(nombre_servicio):
            return

        # Intentar palabras clave individuales (ej. "Comprobantes", "Retenciones")
        for kw in keywords:
            if len(kw) > 4:  # ignorar palabras cortas como "mis"
                if await _encontrar_y_clickear(kw.capitalize()):
                    return

        # 3) Loguear todos los links visibles para debug
        self.log("⚠️ Servicios visibles en el portal:")
        try:
            links = await self.page.locator("a, button").all()
            textos_visibles = []
            for link in links[:40]:
                txt = (await link.inner_text()).strip()
                if txt and len(txt) > 2:
                    textos_visibles.append(txt)
            self.log("   " + " | ".join(textos_visibles[:20]))
        except Exception:
            pass

        raise Exception(
            f"No se encontró '{nombre_servicio}' en el portal. "
            "Revisá los servicios visibles en el log para ajustar el nombre."
        )

    # ------------------------------------------------------------------
    # Mis Comprobantes
    # ------------------------------------------------------------------

    async def get_comprobantes_emitidos(self, fecha_desde: str, fecha_hasta: str) -> list[dict]:
        self.log("📄 Obteniendo comprobantes emitidos...")
        await self._ir_a_servicio("Mis Comprobantes")
        await self.page.wait_for_timeout(4000)
        return await self._extraer_comprobantes("emitidos", fecha_desde, fecha_hasta)

    async def get_comprobantes_recibidos(self, fecha_desde: str, fecha_hasta: str) -> list[dict]:
        self.log("📄 Obteniendo comprobantes recibidos...")
        # Volver al servicio para asegurarnos de estar en el estado correcto
        await self._ir_a_servicio("Mis Comprobantes")
        await self.page.wait_for_timeout(4000)
        return await self._extraer_comprobantes("recibidos", fecha_desde, fecha_hasta)

    async def _click_tab(self, variantes: list[str]) -> bool:
        """Intenta hacer click en una pestaña probando múltiples nombres."""
        for txt in variantes:
            for sel in [f"a:has-text('{txt}')", f"button:has-text('{txt}')",
                        f"li:has-text('{txt}')", f"span:has-text('{txt}')",
                        f"text={txt}"]:
                try:
                    elem = self.page.locator(sel).first
                    if await elem.count() > 0:
                        await elem.click(force=True)
                        await self.page.wait_for_timeout(2000)
                        self.log(f"   ✅ Pestaña '{txt}' seleccionada")
                        return True
                except Exception:
                    continue
        return False

    async def _debug_snapshot(self, nombre: str):
        """Guarda screenshot + primeras líneas de HTML para debug."""
        import os
        out = os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(out, exist_ok=True)
        try:
            img_path = os.path.join(out, f"debug_{nombre}.png")
            await self.page.screenshot(path=img_path, full_page=False)
            self.log(f"   📸 Screenshot: output/debug_{nombre}.png")
        except Exception as e:
            self.log(f"   ⚠️ Screenshot falló: {e}")
        try:
            html_path = os.path.join(out, f"debug_{nombre}.html")
            html = await self.page.content()
            with open(html_path, "w", encoding="utf-8", errors="replace") as f:
                f.write(html)
            self.log(f"   💾 HTML: output/debug_{nombre}.html ({len(html)} bytes)")
        except Exception as e:
            self.log(f"   ⚠️ HTML dump falló: {e}")

    async def _extraer_comprobantes(self, tipo: str, fecha_desde: str, fecha_hasta: str) -> list[dict]:
        """Selecciona pestaña, ingresa fechas, busca y extrae comprobantes."""
        await self.page.wait_for_timeout(2000)

        # Seleccionar pestaña emitidos o recibidos
        if tipo == "emitidos":
            variantes = ["Emitidos", "Comprobantes Emitidos", "EMITIDOS", "Emitido"]
        else:
            variantes = ["Recibidos", "Comprobantes Recibidos", "RECIBIDOS", "Recibido"]

        encontrado = await self._click_tab(variantes)
        if not encontrado:
            self.log(f"⚠️ No se encontró pestaña para '{tipo}', continuando igual...")

        await self.page.wait_for_timeout(2000)

        # Ingresar rango de fechas
        await self._set_rango_fecha(fecha_desde, fecha_hasta)
        await self.page.wait_for_timeout(1000)

        # Click en Buscar
        buscado = False
        for btn_id in ["#buscarComprobantes", "#buscar", "#consultar", "#btnBuscar"]:
            try:
                btn = self.page.locator(btn_id).first
                if await btn.count() > 0:
                    await btn.click(force=True)
                    buscado = True
                    self.log(f"   🔍 Click Buscar por ID: {btn_id}")
                    break
            except Exception:
                continue

        if not buscado:
            for btn_text in ["Buscar", "Consultar", "Ver Comprobantes", "Buscar comprobantes"]:
                try:
                    btn = self.page.locator(
                        f"button:has-text('{btn_text}'), input[value='{btn_text}']"
                    ).first
                    if await btn.count() > 0:
                        await btn.click(force=True)
                        buscado = True
                        self.log(f"   🔍 Click Buscar por texto: '{btn_text}'")
                        break
                except Exception:
                    continue

        if not buscado:
            self.log("   ⚠️ No se encontró botón Buscar")

        # Esperar resultados (la pestaña "Resultados" puede activarse)
        await self.page.wait_for_timeout(5000)

        # Si hay pestaña "Resultados", hacer click en ella para ver los datos
        try:
            tab_resultados = self.page.locator(
                "a:has-text('Resultados'), button:has-text('Resultados'), "
                "li:has-text('Resultados'), [role='tab']:has-text('Resultados')"
            ).first
            if await tab_resultados.count() > 0:
                await tab_resultados.click(force=True)
                await self.page.wait_for_timeout(2000)
                self.log("   ✅ Click en pestaña 'Resultados'")
        except Exception:
            pass

        # Intentar descargar CSV si está disponible (más completo que parsear HTML)
        csv_data = await self._descargar_csv_resultados(tipo=tipo)
        if csv_data:
            self.log(f"   ✅ {len(csv_data)} filas obtenidas por CSV")
            return csv_data

        # Extraer tabla de resultados
        comprobantes = []
        page_num = 1
        while True:
            self.log(f"   📃 Página {page_num} de {tipo}...")
            rows = await self._extraer_tabla_comprobantes()
            comprobantes.extend(rows)
            self.log(f"   → {len(rows)} filas en página {page_num}")

            siguiente = self.page.locator(
                "a:has-text('Siguiente'), button:has-text('Siguiente'), [aria-label='Siguiente']"
            ).first
            if await siguiente.count() == 0:
                break
            disabled = await siguiente.get_attribute("disabled")
            if disabled is not None:
                break
            await siguiente.click()
            await self.page.wait_for_load_state("load")
            await self.page.wait_for_timeout(2000)
            page_num += 1

        self.log(f"   ✅ {len(comprobantes)} comprobantes {tipo} obtenidos")
        return comprobantes

    MESES_ES = {
        "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
        "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
    }

    async def _set_rango_fecha(self, fecha_desde: str, fecha_hasta: str):
        """
        Rellena el campo de rango de fechas de ARCA navegando el calendario:
        1. Click en el campo para abrir el popup
        2. Usar flechas ◄ ► para llegar al mes correcto
        3. Click en día 1 del mes
        4. Click en el último día del mes
        5. Click en Aplicar
        """
        import calendar as _cal
        self.log(f"   📅 Navegando calendario: {fecha_desde} → {fecha_hasta}")

        try:
            d1, m1, y1 = fecha_desde.split("/")
            mes_target  = int(m1)
            anio_target = int(y1)
            ultimo_dia  = int(fecha_hasta.split("/")[0])
        except Exception:
            self.log("   ⚠️ Formato de fecha inválido")
            return

        # Paso 1: abrir el popup clickeando el campo de fecha principal
        for inp in await self.page.locator("input").all():
            try:
                val = await inp.input_value()
                if re.search(r"\d{2}/\d{2}/\d{4}", val or ""):
                    await inp.click(force=True)
                    await self.page.wait_for_timeout(1500)
                    self.log("   🖱️ Popup abierto")
                    break
            except Exception:
                continue

        # Paso 2: leer el mes/año actual del calendario y navegar hasta el target
        await self._navegar_mes_calendario(mes_target, anio_target)

        # Paso 3: click en el día 1 del mes
        await self._click_dia_calendario(1)
        await self.page.wait_for_timeout(400)

        # Paso 4: click en el último día del mes
        await self._click_dia_calendario(ultimo_dia)
        await self.page.wait_for_timeout(600)

        # Paso 5: click en "Aplicar"
        await self._click_aplicar()

    async def _leer_mes_anio_calendario(self) -> tuple[int, int]:
        """Lee el mes y año del encabezado del calendario abierto."""
        try:
            # El encabezado dice algo como "Junio 2026"
            header = await self.page.evaluate("""() => {
                // Buscar el encabezado del calendario en el DOM
                const selectors = [
                    '.month', '.calendar-header', '[class*="monthYear"]',
                    '[class*="month-year"]', 'th.month', '.datepicker-switch',
                    '[class*="header"] b', '[class*="header"] span',
                ];
                for (const sel of selectors) {
                    const el = document.querySelector(sel);
                    if (el && el.textContent.trim().length > 3) {
                        return el.textContent.trim();
                    }
                }
                // Buscar cualquier texto que parezca "Mes Año"
                const all = document.querySelectorAll('*');
                for (const el of all) {
                    const t = el.textContent.trim();
                    if (/^[A-Za-záéíóúÁÉÍÓÚ]+ \\d{4}$/.test(t)) return t;
                }
                return '';
            }""")
            self.log(f"   📆 Encabezado calendario: '{header}'")
            partes = header.lower().split()
            if len(partes) >= 2:
                mes_str  = partes[0]
                anio_str = partes[-1]
                mes_num  = self.MESES_ES.get(mes_str, 0)
                anio_num = int(anio_str) if anio_str.isdigit() else 0
                return mes_num, anio_num
        except Exception as e:
            self.log(f"   ⚠️ Error leyendo encabezado: {e}")
        return 0, 0

    async def _navegar_mes_calendario(self, mes_target: int, anio_target: int):
        """Navega el calendario con las flechas hasta llegar al mes/año target."""
        for _ in range(24):  # máximo 24 clicks de navegación
            mes_actual, anio_actual = await self._leer_mes_anio_calendario()
            if mes_actual == 0:
                self.log("   ⚠️ No se pudo leer el mes del calendario")
                break

            diff = (anio_target - anio_actual) * 12 + (mes_target - mes_actual)
            self.log(f"   🗓️ Calendario en {mes_actual}/{anio_actual}, target {mes_target}/{anio_target}, diff={diff}")

            if diff == 0:
                self.log("   ✅ Mes correcto en el calendario")
                break
            elif diff > 0:
                # Ir hacia adelante (flecha derecha ►)
                await self._click_flecha_calendario("siguiente")
            else:
                # Ir hacia atrás (flecha izquierda ◄)
                await self._click_flecha_calendario("anterior")
            await self.page.wait_for_timeout(500)

    async def _click_flecha_calendario(self, direccion: str):
        """Hace click en la flecha de navegación del calendario."""
        if direccion == "siguiente":
            sels = ["button.next", ".next", "button:has-text('›')",
                    "button:has-text('>')", "[aria-label*='siguiente' i]",
                    "[aria-label*='next' i]", "[class*='next']", "th.next"]
        else:
            sels = ["button.prev", ".prev", "button:has-text('‹')",
                    "button:has-text('<')", "[aria-label*='anterior' i]",
                    "[aria-label*='prev' i]", "[class*='prev']", "th.prev"]

        for sel in sels:
            try:
                btn = self.page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(force=True)
                    self.log(f"   ▶ Flecha {direccion}: {sel}")
                    return
            except Exception:
                continue

        # Fallback JavaScript: buscar SVG de flecha o botón con <
        try:
            await self.page.evaluate(f"""() => {{
                const sels = {sels!r};
                for (const s of sels) {{
                    const el = document.querySelector(s);
                    if (el) {{ el.click(); return; }}
                }}
            }}""")
        except Exception:
            pass

    async def _click_dia_calendario(self, dia: int):
        """Hace click en un día específico del calendario visible."""
        self.log(f"   🖱️ Click en día {dia}")

        # Selectores típicos de celdas de días en datepickers
        sels = [
            f"td:has-text('{dia}')",
            f"[class*='day']:has-text('{dia}')",
            f"button:has-text('{dia}')",
            f"span:has-text('{dia}')",
        ]

        for sel in sels:
            try:
                # Filtrar: el día debe ser exactamente ese número (no el 1 en "21", etc.)
                elems = await self.page.locator(sel).all()
                for elem in elems:
                    txt = (await elem.inner_text()).strip()
                    cls = await elem.get_attribute("class") or ""
                    # Excluir días de otros meses (tienen clase "off", "disabled", "muted")
                    if (txt == str(dia) and
                        "off"      not in cls and
                        "disabled" not in cls and
                        "muted"    not in cls and
                        "other"    not in cls):
                        await elem.click(force=True)
                        self.log(f"   ✅ Día {dia} clickeado")
                        return
            except Exception:
                continue

        self.log(f"   ⚠️ No se encontró día {dia} en el calendario")

    async def _click_aplicar(self):
        """Hace click en el botón Aplicar del popup de fechas."""
        # Por texto
        for sel in ["button:has-text('Aplicar')", "button:has-text('Apply')",
                    ".applyBtn", "[class*='applyBtn']", "[class*='apply-btn']"]:
            try:
                btn = self.page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(force=True)
                    await self.page.wait_for_timeout(1000)
                    self.log("   ✅ Click en 'Aplicar'")
                    return
            except Exception:
                continue

        # Por JavaScript
        try:
            clicked = await self.page.evaluate("""() => {
                const btns = Array.from(document.querySelectorAll('button'));
                const b = btns.find(b =>
                    b.textContent.trim().toLowerCase().includes('aplicar') ||
                    b.textContent.trim().toLowerCase().includes('apply') ||
                    b.className.includes('apply')
                );
                if (b) { b.click(); return true; }
                return false;
            }""")
            if clicked:
                await self.page.wait_for_timeout(1000)
                self.log("   ✅ Click en 'Aplicar' (JS)")
                return
        except Exception:
            pass

        self.log("   ⚠️ Botón 'Aplicar' no encontrado — presionando Enter")
        await self.page.keyboard.press("Enter")
        await self.page.wait_for_timeout(500)

    async def _limpiar_y_escribir(self, campo, fecha: str):
        """Limpia un input de fecha y escribe la nueva fecha DD/MM/AAAA."""
        try:
            await campo.click(force=True)
            await self.page.wait_for_timeout(200)
            await self.page.keyboard.press("Control+a")
            await self.page.keyboard.press("Delete")
            await self.page.wait_for_timeout(100)
            await self.page.keyboard.type(fecha, delay=60)
            await self.page.wait_for_timeout(200)
            self.log(f"   ✅ Fecha escrita: {fecha}")
        except Exception as e:
            self.log(f"   ⚠️ Error escribiendo {fecha}: {e}")

    async def _fill_date_field(self, campo, valor: str):
        """
        Llena un campo de fecha individual.
        Maneja calendarios emergentes de Angular Material.
        valor esperado: DD/MM/AAAA
        """
        try:
            d, m, y = valor.split("/")
        except Exception:
            d, m, y = "01", "01", "2026"

        # Intentar con solo dígitos (DDMMAAAA) — Angular Material los acepta así
        solo_digitos = f"{d}{m}{y}"

        for intento, texto in enumerate([solo_digitos, valor]):
            try:
                await campo.click(force=True)
                await self.page.wait_for_timeout(300)

                # Cerrar calendario si se abrió (presionar Escape)
                await self.page.keyboard.press("Escape")
                await self.page.wait_for_timeout(200)

                # Volver a hacer click para focusear el input
                await campo.click(force=True)
                await self.page.wait_for_timeout(200)

                # Seleccionar todo y escribir
                await self.page.keyboard.press("Control+a")
                await self.page.keyboard.type(texto, delay=50)
                await self.page.wait_for_timeout(300)

                # Verificar si el valor quedó ingresado
                val_actual = await campo.input_value()
                if val_actual and val_actual.strip() not in ("", "dd/mm/aaaa", "dd/mm/yyyy"):
                    self.log(f"   ✅ Campo fecha = '{val_actual}' (intento {intento+1})")
                    return
            except Exception as e:
                self.log(f"   ⚠️ Intento {intento+1} falló: {e}")
                continue

        # Último recurso: JavaScript
        try:
            await self._js_set_value(campo, valor)
            self.log(f"   ✅ Campo fecha ingresado por JS: {valor}")
        except Exception as e:
            self.log(f"   ⚠️ JS también falló: {e}")

    async def _set_fecha(self, campo: str, valor: str):
        """Rellena un campo de fecha. Soporta formato DD/MM/AAAA."""
        es_desde = "desde" in campo.lower()

        # Convertir a distintos formatos
        formatos = [valor]
        try:
            d, m, y = valor.split("/")
            formatos.append(f"{y}-{m}-{d}")  # ISO
        except Exception:
            pass

        # Selectores específicos para ARCA Mis Comprobantes
        selectors = [
            "#fechaEmisionDesde" if es_desde else "#fechaEmisionHasta",
            "#fechaDesde"        if es_desde else "#fechaHasta",
            f"input[id*='Desde']" if es_desde else f"input[id*='Hasta']",
            f"input[name*='Desde']" if es_desde else f"input[name*='Hasta']",
            f"input[id*='desde']" if es_desde else f"input[id*='hasta']",
            "#fechaEmision",   # campo genérico — se usa con índice
            "#comprobanteDesde" if es_desde else "#comprobanteHasta",
        ]

        for sel in selectors:
            try:
                elems = await self.page.locator(sel).all()
                if not elems:
                    continue
                idx = 0 if es_desde else (len(elems) - 1)
                elem = elems[idx]

                # Intentar: JavaScript setNativeValue + eventos DOM
                for fmt in formatos:
                    try:
                        await self._js_set_value(elem, fmt)
                        self.log(f"   ✅ Fecha {campo} = {fmt} (JS)")
                        return
                    except Exception:
                        pass

                # Fallback: triple_click + type teclado
                for fmt in formatos:
                    try:
                        await elem.click(force=True)
                        await self.page.wait_for_timeout(100)
                        await elem.triple_click(force=True)
                        await self.page.keyboard.press("Control+a")
                        await self.page.keyboard.type(fmt, delay=50)
                        await self.page.keyboard.press("Tab")
                        await self.page.wait_for_timeout(300)
                        self.log(f"   ✅ Fecha {campo} = {fmt} (keyboard)")
                        return
                    except Exception:
                        pass
            except Exception:
                continue

        self.log(f"   ⚠️ No se pudo ingresar fecha {campo}={valor}")

    async def _js_set_value(self, elem, value: str):
        """Usa JavaScript para disparar eventos de input nativos (Locator.evaluate)."""
        await elem.evaluate("""(el, v) => {
            var nativeInputValueSetter = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'value').set;
            nativeInputValueSetter.call(el, v);
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
        }""", value)

    async def _descargar_csv_resultados(self, tipo: str = "emitidos") -> list[dict]:
        """
        Intenta descargar el CSV que ofrece ARCA en la página de resultados.
        Si lo descarga, lo parsea y devuelve la lista de comprobantes.
        Si no, devuelve lista vacía (el caller cae en parseo HTML).
        """
        import os, glob, time, csv, io

        for sel in ["a:has-text('CSV')", "button:has-text('CSV')",
                    "[class*='csv']", "a[href*='csv' i]"]:
            try:
                btn = self.page.locator(sel).first
                if await btn.count() == 0:
                    continue

                output_dir = os.path.join(os.path.dirname(__file__), "output")
                t_antes = time.time()

                try:
                    async with self.page.expect_download(timeout=15000) as dl:
                        await btn.click(force=True)
                    download = await dl.value
                    path = os.path.join(output_dir, f"arca_csv_{int(t_antes)}.csv")
                    await download.save_as(path)
                    self.log(f"   📥 CSV descargado: {path}")
                except Exception:
                    # Fallback: buscar en Descargas del usuario
                    await self.page.wait_for_timeout(4000)
                    descargas = os.path.expandvars(r"%USERPROFILE%\Downloads")
                    for pat in ["*.csv", "*.CSV"]:
                        for f in glob.glob(os.path.join(descargas, pat)):
                            if time.time() - os.path.getmtime(f) < 30:
                                path = f
                                self.log(f"   📥 CSV encontrado en Descargas: {path}")
                                break
                    else:
                        continue

                # Parsear CSV (ARCA descarga un ZIP que contiene el CSV)
                import zipfile
                if zipfile.is_zipfile(path):
                    with zipfile.ZipFile(path) as zf:
                        csv_name = next(
                            (n for n in zf.namelist() if n.lower().endswith(".csv")), None
                        )
                        if not csv_name:
                            continue
                        raw = zf.read(csv_name)
                        # ARCA exporta en UTF-8; fallback a latin-1 si falla
                        try:
                            contenido = raw.decode("utf-8")
                        except UnicodeDecodeError:
                            contenido = raw.decode("latin-1", errors="replace")
                        self.log(f"   📦 ZIP extraído: {csv_name}")
                else:
                    try:
                        with open(path, "r", encoding="utf-8") as f:
                            contenido = f.read()
                    except UnicodeDecodeError:
                        with open(path, "r", encoding="latin-1", errors="replace") as f:
                            contenido = f.read()

                separador = ";" if contenido.count(";") > contenido.count(",") else ","
                reader = csv.DictReader(io.StringIO(contenido), delimiter=separador)
                comprobantes = []
                for row in reader:
                    comp = self._parsear_fila_csv_arca(dict(row), tipo=tipo)
                    if comp:
                        comprobantes.append(comp)
                return comprobantes

            except Exception as e:
                self.log(f"   ⚠️ CSV ({sel}): {e}")
                continue

        return []

    # Mapeo de código numérico de tipo de comprobante → descripción
    TIPO_COMPROBANTE = {
        "1": "1 - Factura A", "2": "2 - Nota de Débito A", "3": "3 - Nota de Crédito A",
        "4": "4 - Recibo A",  "5": "5 - Nota de Venta A",
        "6": "6 - Factura B", "7": "7 - Nota de Débito B", "8": "8 - Nota de Crédito B",
        "9": "9 - Recibo B",  "10": "10 - Nota de Venta B",
        "11": "11 - Factura C", "12": "12 - Nota de Débito C", "13": "13 - Nota de Crédito C",
        "15": "15 - Recibo C",
        "51": "51 - Factura M", "52": "52 - Nota de Débito M", "53": "53 - Nota de Crédito M",
        "81": "81 - Tique Factura A", "82": "82 - Tique Factura B",
        "83": "83 - Nota de Crédito Tique", "111": "111 - Tique",
        "201": "201 - FCE Factura A", "202": "202 - FCE Nota de Débito A", "203": "203 - FCE Nota de Crédito A",
        "206": "206 - FCE Factura B", "207": "207 - FCE Nota de Débito B", "208": "208 - FCE Nota de Crédito B",
        "211": "211 - FCE Factura C", "212": "212 - FCE Nota de Débito C", "213": "213 - FCE Nota de Crédito C",
    }

    def _parsear_fila_csv_arca(self, row: dict, tipo: str = "emitidos") -> dict | None:
        """
        Parsea una fila del CSV de ARCA Mis Comprobantes.
        Columnas relevantes del CSV de ARCA:
          Fecha de Emisión | Tipo de Comprobante | Punto de Venta | Número Desde |
          Nro. Doc. Receptor | Denominación Receptor |
          IVA 10,5% | Imp. Neto Gravado IVA 10,5% | IVA 21% | Imp. Neto Gravado IVA 21% |
          Imp. Op. Exentas | Imp. Total
        """
        def m(s):
            try:
                return float(str(s).strip().replace(".", "").replace(",", ".").replace("$", "").replace(" ", ""))
            except Exception:
                return 0.0

        def f(s):
            s = str(s).strip()
            for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]:
                try:
                    return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
                except Exception:
                    pass
            return s

        # Búsqueda de columna: exacta con strip, luego por substring case-insensitive
        def get(*candidates):
            for candidate in candidates:
                # Exacto con strip
                for k, v in row.items():
                    if k.strip() == candidate:
                        return str(v).strip()
                # Substring case-insensitive
                c_lower = candidate.lower()
                for k, v in row.items():
                    if c_lower in k.strip().lower():
                        return str(v).strip()
            return ""

        # Fecha — "Fecha de Emisión"
        fecha = f(get("Fecha de Emisión", "Fecha"))

        # Tipo de comprobante — CSV trae código numérico ("6"), convertir a descripción
        tipo_cod = get("Tipo de Comprobante", "Tipo")
        tipo_desc = self.TIPO_COMPROBANTE.get(tipo_cod, f"{tipo_cod} - Comprobante") if tipo_cod else ""

        # N° Comprobante: XXXXX-YYYYYYYY (Punto de Venta + Número Desde)
        pv        = get("Punto de Venta")
        num_desde = get("Número Desde", "Numero Desde")
        try:
            nro_comp = f"{int(pv):05d}-{int(num_desde):08d}"
        except Exception:
            nro_comp = f"{pv}-{num_desde}"

        # CUIT: en emitidos el interlocutor es el Receptor (cliente);
        #       en recibidos el interlocutor es el Emisor (proveedor).
        # NOTA: se usa el parámetro `tipo` original ("emitidos"/"recibidos"),
        #       no `tipo_desc` (que ya contiene la descripción del comprobante).
        if tipo == "recibidos":
            cuit_raw = (get("Nro. Doc. Emisor") or
                        get("CUIT Emisor") or
                        get("Nro. Doc. Receptor") or "")
        else:
            cuit_raw = (get("Nro. Doc. Receptor") or
                        get("Nro. Doc. Emisor") or
                        get("CUIT Emisor") or "")
        cuit = re.sub(r"[^0-9]", "", cuit_raw) if cuit_raw else ""

        # Razón social: en emitidos = Receptor (cliente); en recibidos = Emisor (proveedor)
        if tipo == "recibidos":
            rs = (get("Denominación Emisor", "Denominacion Emisor") or
                  get("Razón Social Emisor", "Razon Social Emisor") or
                  get("Denominación Receptor", "Denominacion Receptor") or "")
        else:
            rs = (get("Denominación Receptor", "Denominacion Receptor") or
                  get("Denominación Emisor", "Denominacion Emisor") or
                  get("Razón Social Emisor", "Razon Social Emisor") or "")

        # IVA 21% — ARCA puede llamar la columna con o sin "Imp. Neto Gravado"
        iva_21  = m(get("IVA 21%", "Alicuota 21%", "Alícuota 21%"))
        neto_21 = m(get(
            "Imp. Neto Gravado IVA 21%",
            "Neto Gravado IVA 21%",
            "Base Imponible 21%",
            "Neto 21%",
        ))

        # IVA 10,5% — ARCA puede escribirlo con coma (10,5%) o punto (10.5%)
        iva_105  = m(get(
            "IVA 10,5%", "IVA 10.5%",
            "Alicuota 10,5%", "Alicuota 10.5%",
            "Alícuota 10,5%", "Alícuota 10.5%",
        ))
        neto_105 = m(get(
            "Imp. Neto Gravado IVA 10,5%", "Imp. Neto Gravado IVA 10.5%",
            "Neto Gravado IVA 10,5%",      "Neto Gravado IVA 10.5%",
            "Base Imponible 10,5%",         "Base Imponible 10.5%",
            "Neto 10,5%",                   "Neto 10.5%",
        ))

        # IVA 27%
        iva_27  = m(get("IVA 27%", "Alicuota 27%", "Alícuota 27%"))
        neto_27 = m(get(
            "Imp. Neto Gravado IVA 27%",
            "Neto Gravado IVA 27%",
            "Base Imponible 27%",
            "Neto 27%",
        ))

        # Exento y total
        exento = m(get("Imp. Op. Exentas", "Op. Exentas", "Exento", "Exenta"))
        total  = m(get("Imp. Total", "Total", "Importe Total"))

        # ── Inferir neto desde IVA cuando ARCA no lo da separado ──────────
        # ARCA a veces exporta solo el monto de IVA por alícuota sin el neto.
        # Si tenemos IVA pero neto = 0 → calculamos: neto = IVA / alícuota
        if iva_21 > 0 and neto_21 == 0:
            neto_21 = round(iva_21 / 0.21, 2)
        if iva_105 > 0 and neto_105 == 0:
            neto_105 = round(iva_105 / 0.105, 2)
        if iva_27 > 0 and neto_27 == 0:
            neto_27 = round(iva_27 / 0.27, 2)

        # Si el total vino del CSV pero todos los campos de neto+IVA son 0,
        # intentar analizar la alícuota comparando IVA con total
        if total > 0 and (neto_21 + iva_21 + neto_105 + iva_105 + neto_27 + iva_27 + exento) == 0:
            # Buscar campos de IVA alternativos más genéricos
            iva_raw = m(get("IVA", "Imp. IVA", "Impuesto IVA", "Monto IVA"))
            if iva_raw > 0:
                # Determinar alícuota por proporción: IVA/total
                ratio = iva_raw / total
                if abs(ratio - 0.21 / 1.21) < 0.01:       # ≈ 0.1736 → 21%
                    iva_21  = round(iva_raw, 2)
                    neto_21 = round(total - iva_21 - exento, 2)
                elif abs(ratio - 0.105 / 1.105) < 0.01:    # ≈ 0.0950 → 10.5%
                    iva_105  = round(iva_raw, 2)
                    neto_105 = round(total - iva_105 - exento, 2)
                elif abs(ratio - 0.27 / 1.27) < 0.01:      # ≈ 0.2126 → 27%
                    iva_27  = round(iva_raw, 2)
                    neto_27 = round(total - iva_27 - exento, 2)

        if not fecha and not tipo_cod:
            return None

        return {
            "fecha": fecha, "tipo": tipo_desc, "numero": nro_comp,
            "cuit_contraparte": cuit,
            "razon_social": rs,
            "neto_21": neto_21, "iva_21": iva_21,
            "neto_105": neto_105, "iva_105": iva_105,
            "neto_27": neto_27, "iva_27": iva_27,
            "exento": exento, "total": total,
        }

    async def _extraer_tabla_comprobantes(self) -> list[dict]:
        """Extrae filas de la tabla de comprobantes visible en pantalla."""
        comprobantes = []
        try:
            # Estrategia 1: tabla HTML clásica
            filas = await self.page.locator("table tbody tr").all()
            if filas:
                self.log(f"   🗂 Encontradas {len(filas)} filas en <table>")
                for fila in filas:
                    celdas = await fila.locator("td").all()
                    if len(celdas) < 4:
                        continue
                    textos = [await c.inner_text() for c in celdas]
                    comp = self._parsear_fila_comprobante(textos)
                    if comp:
                        comprobantes.append(comp)
                return comprobantes

            # Estrategia 2: filas con atributos de lista/grilla (Angular/React)
            for selector in [
                "tr[class*='row']",
                "div[class*='row']:not([class*='header'])",
                "li[class*='item']",
                "[role='row']",
                "[class*='comprobante']",
                "[class*='resultado']",
            ]:
                filas = await self.page.locator(selector).all()
                if len(filas) > 1:
                    self.log(f"   🗂 Encontradas {len(filas)} filas con '{selector}'")
                    for fila in filas:
                        textos_raw = await fila.inner_text()
                        partes = [p.strip() for p in textos_raw.split("\n") if p.strip()]
                        if len(partes) >= 4:
                            comp = self._parsear_fila_comprobante(partes)
                            if comp:
                                comprobantes.append(comp)
                    if comprobantes:
                        return comprobantes

            # Estrategia 3: extraer toda la tabla como texto y parsear
            for sel_tabla in ["table", "[class*='grid']", "[class*='tabla']", "[class*='listado']"]:
                tabla = self.page.locator(sel_tabla).first
                if await tabla.count() > 0:
                    texto = await tabla.inner_text()
                    self.log(f"   🗂 Texto de tabla ({sel_tabla}): {texto[:200]}")
                    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
                    for linea in lineas:
                        partes = linea.split("\t")
                        if len(partes) < 4:
                            partes = re.split(r"\s{2,}", linea)
                        comp = self._parsear_fila_comprobante(partes)
                        if comp:
                            comprobantes.append(comp)
                    if comprobantes:
                        return comprobantes

        except Exception as e:
            self.log(f"⚠️ Error extrayendo tabla: {e}")

        if not comprobantes:
            self.log("   ⚠️ No se encontraron filas de datos en la página")
        return comprobantes

    def _parsear_fila_comprobante(self, textos: list[str]) -> dict | None:
        """
        Parsea una fila de comprobante. La estructura típica de Mis Comprobantes ARCA es:
        Fecha | Tipo | Punto Venta | Número | CUIT | Razón Social | Imp. Neto No Grav | Neto Grav 21% | IVA 21% | Neto Grav 10.5% | IVA 10.5% | Exento | Total
        """
        if len(textos) < 4:
            return None

        def parse_monto(s: str) -> float:
            s = s.strip().replace(".", "").replace(",", ".").replace("$", "").replace(" ", "")
            try:
                return float(s)
            except Exception:
                return 0.0

        def parse_fecha(s: str) -> str:
            s = s.strip()
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                try:
                    return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
                except Exception:
                    pass
            return s

        try:
            # Estructura flexible: buscar columnas por posición aproximada
            # Típico: Fecha|Tipo|PV-Num|CUIT|RS|NetoNoGrav|Neto21|IVA21|Neto10.5|IVA10.5|Exento|Total
            n = len(textos)
            neto_21  = parse_monto(textos[6]) if n > 6 else 0.0
            iva_21   = parse_monto(textos[7]) if n > 7 else 0.0
            neto_105 = parse_monto(textos[8]) if n > 8 else 0.0
            iva_105  = parse_monto(textos[9]) if n > 9 else 0.0
            exento   = parse_monto(textos[10]) if n > 10 else 0.0
            total    = parse_monto(textos[-1]) if textos else 0.0

            # Fallback: si la tabla tiene menos columnas,
            # el neto puede estar en posición 5 y el IVA en 6
            if neto_21 == 0 and iva_21 == 0 and n in (8, 9, 10):
                neto_21 = parse_monto(textos[5]) if n > 5 else 0.0
                iva_21  = parse_monto(textos[6]) if n > 6 else 0.0

            # Inferir neto desde IVA si ARCA solo da el monto de impuesto
            if iva_21 > 0 and neto_21 == 0:
                neto_21 = round(iva_21 / 0.21, 2)
            if iva_105 > 0 and neto_105 == 0:
                neto_105 = round(iva_105 / 0.105, 2)

            comp = {
                "fecha": parse_fecha(textos[0]) if n > 0 else "",
                "tipo": textos[1].strip() if n > 1 else "",
                "numero": textos[2].strip() if n > 2 else "",
                "cuit_contraparte": textos[3].strip() if n > 3 else "",
                "razon_social": textos[4].strip() if n > 4 else "",
                "neto_21": neto_21, "iva_21": iva_21,
                "neto_105": neto_105, "iva_105": iva_105,
                "neto_27": 0.0, "iva_27": 0.0,
                "exento": exento, "total": total,
            }
            # Filtrar filas vacías o de encabezado
            if not comp["fecha"] or comp["tipo"] in ("", "Tipo"):
                return None
            return comp
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Mis Retenciones - SICORE 767
    # ------------------------------------------------------------------

    async def get_retenciones_sicore(self, fecha_desde: str, fecha_hasta: str) -> list[dict]:
        """
        Navega a Mis Retenciones → versión anterior (click aquí) →
        rellena el formulario:
          - CUIT del retenido  : dropdown → valor = self.cuit
          - Impuesto retenido  : input "767" + select "767 - SICORE..."
          - Radio              : "Retención y/o Percepción"
          - Fecha Ret./Perc.   : Desde / Hasta en formato ddMMaaaa
        → CONSULTAR → Exportar a Excel → parsea el .xls descargado.

        Estructura del Excel:
          CUIT Agente Ret./Perc. | Denominación o Razón Social | Impuesto |
          Descripción Impuesto | Régimen | Descripción Régimen |
          Fecha Ret./Perc. | Número Certificado | Descripción Operación |
          Importe Ret./Perc. | Número Comprobante | Fecha Comprobante |
          Descripción Comprobante | Fecha Registración DJ Ag.Ret.
        """
        # os, glob y _time ya importados a nivel de módulo

        self.log("🔒 Obteniendo retenciones SICORE 767...")

        # Convertir fechas de DD/MM/AAAA → ddMMaaaa (formato ARCA versión anterior)
        def _a_ddmmaaaa(fecha_slash: str) -> str:
            partes = fecha_slash.replace("-", "/").split("/")
            if len(partes) == 3:
                d, m, a = partes[0].zfill(2), partes[1].zfill(2), partes[2]
                return f"{d}{m}{a}"
            return fecha_slash

        fecha_desde_fmt = _a_ddmmaaaa(fecha_desde)
        fecha_hasta_fmt = _a_ddmmaaaa(fecha_hasta)
        self.log(f"   📅 Fechas formato SICORE: {fecha_desde_fmt} → {fecha_hasta_fmt}")

        # ── Paso 1: ir a Mis Retenciones ────────────────────────────────
        await self._ir_a_servicio("Mis Retenciones")
        await self.page.wait_for_timeout(4000)

        # ── Paso 2: click en "click aquí" (versión anterior) ────────────
        self.log("   🔗 Abriendo versión anterior...")
        for sel in [
            "a:has-text('click aquí')",
            "a:has-text('click aqui')",
            "a:has-text('versión anterior')",
            "a:has-text('version anterior')",
            "a:has-text('Acceder')",
            "text=click aquí",
        ]:
            try:
                elem = self.page.locator(sel).first
                if await elem.count() == 0:
                    continue
                try:
                    async with self.context.expect_page(timeout=6000) as np_info:
                        await elem.click(force=True)
                    nueva = await np_info.value
                    await nueva.wait_for_load_state("load")
                    self.page = nueva
                except Exception:
                    await elem.click(force=True)
                    await self.page.wait_for_load_state("load")
                await self.page.wait_for_timeout(3000)
                self.log(f"   ✅ Versión anterior cargada ({sel})")
                break
            except Exception:
                continue

        await self.page.wait_for_timeout(2000)

        # ── Paso 3: CUIT del retenido (primer <select> de la página) ────
        self.log(f"   🔎 Seleccionando CUIT del retenido: {self.cuit}...")
        try:
            # El select "CUIT del retenido" suele ser el primero en la página
            selects = await self.page.locator("select").all()
            for sel_elem in selects:
                opts = await sel_elem.locator("option").all()
                for opt in opts:
                    val = (await opt.get_attribute("value") or "").strip()
                    txt = (await opt.inner_text()).strip()
                    if self.cuit in val or self.cuit in txt:
                        await sel_elem.select_option(value=val if val else txt)
                        await self.page.wait_for_timeout(1500)
                        self.log("   ✅ CUIT del retenido seleccionado")
                        break
                else:
                    continue
                break
            else:
                self.log("   ℹ️ CUIT ya seleccionado o único")
        except Exception as e:
            self.log(f"   ⚠️ Selección CUIT: {e}")

        # ── Paso 4: Impuesto retenido = 767 ─────────────────────────────
        await self.page.wait_for_timeout(1000)
        self.log("   📋 Seleccionando impuesto 767-SICORE...")
        try:
            # Hay un input numérico y un select descriptivo
            # Primero llenamos el input de código con "767"
            for inp_sel in [
                "input[name*='mpuesto']", "input[id*='mpuesto']",
                "input[name*='impuesto']", "input[id*='impuesto']",
            ]:
                inp = self.page.locator(inp_sel).first
                if await inp.count() > 0:
                    await self._js_set_value(inp, "767")
                    await self.page.keyboard.press("Tab")
                    await self.page.wait_for_timeout(800)
                    self.log("   ✅ Código impuesto '767' ingresado")
                    break

            # Luego seleccionamos el select descriptivo que contiene "767"
            for sel_elem in await self.page.locator("select").all():
                opts = await sel_elem.locator("option").all()
                for opt in opts:
                    val = (await opt.get_attribute("value") or "").strip()
                    txt = (await opt.inner_text()).strip()
                    if "767" in val or "767" in txt:
                        await sel_elem.select_option(value=val if val else txt)
                        await self.page.wait_for_timeout(800)
                        self.log(f"   ✅ Select impuesto '767' seleccionado")
                        break
                else:
                    continue
                break
        except Exception as e:
            self.log(f"   ⚠️ Selección impuesto: {e}")

        # ── Paso 5: Radio "Retención y/o Percepción" ────────────────────
        await self.page.wait_for_timeout(500)
        self.log("   📻 Seleccionando 'Retención y/o Percepción'...")
        try:
            # Buscar radio por texto cercano o por value
            for radio_sel in [
                "input[type='radio'][value*='3']",   # suele ser el 3ro
                "input[type='radio'][value*='R']",
                "input[type='radio']:near(:text('Retención y/o'))",
                "input[type='radio']:near(:text('Retenci'))",
            ]:
                radios = await self.page.locator(radio_sel).all()
                if radios:
                    await radios[-1].click(force=True)   # "Retención y/o" suele ser el último
                    await self.page.wait_for_timeout(500)
                    self.log("   ✅ Radio 'Retención y/o Percepción' seleccionado")
                    break

            # Fallback: clicar directamente por texto cercano
            if not radios:
                label = self.page.locator("label:has-text('Retención y/o'), label:has-text('Retenci')").last
                if await label.count() > 0:
                    await label.click(force=True)
                    self.log("   ✅ Radio seleccionado via label")
        except Exception as e:
            self.log(f"   ⚠️ Selección radio: {e}")

        # ── Paso 6: Fechas Ret./Perc. Desde / Hasta (formato ddMMaaaa) ──
        await self.page.wait_for_timeout(500)
        self.log(f"   📅 Ingresando fechas {fecha_desde_fmt} → {fecha_hasta_fmt}...")

        # Buscar todos los inputs de texto que puedan ser fechas
        # El formulario tiene: Fecha Ret./Perc. Desde | Hasta
        #                       Fecha Comprobante  Desde | Hasta (los dejamos vacíos)
        # Buscamos los primeros dos inputs de fecha (los de Ret./Perc.)
        async def _fill_date_input(selector: str, value: str, label: str):
            try:
                elem = self.page.locator(selector).first
                if await elem.count() == 0:
                    return False
                await elem.click(force=True)
                await self.page.wait_for_timeout(100)
                await self.page.keyboard.press("Control+a")
                await self.page.keyboard.press("Delete")
                await self.page.keyboard.type(value, delay=40)
                await self.page.keyboard.press("Tab")
                await self.page.wait_for_timeout(200)
                val = await elem.input_value()
                self.log(f"   ✅ {label} = {val or value}")
                return True
            except Exception:
                return False

        # Intentar por nombres/ids típicos del formulario SICORE versión anterior
        fecha_desde_ok = False
        fecha_hasta_ok = False

        for sel in ["input[name='fechaRetencionDesde']", "input[id='fechaRetencionDesde']",
                    "input[name*='etDesde']", "input[id*='etDesde']",
                    "input[name*='ercDesde']", "input[id*='ercDesde']",
                    "input[name*='Desde']", "input[id*='Desde']"]:
            if await _fill_date_input(sel, fecha_desde_fmt, "Fecha Desde"):
                fecha_desde_ok = True
                break

        for sel in ["input[name='fechaRetencionHasta']", "input[id='fechaRetencionHasta']",
                    "input[name*='etHasta']", "input[id*='etHasta']",
                    "input[name*='ercHasta']", "input[id*='ercHasta']",
                    "input[name*='Hasta']", "input[id*='Hasta']"]:
            if await _fill_date_input(sel, fecha_hasta_fmt, "Fecha Hasta"):
                fecha_hasta_ok = True
                break

        # Fallback: encontrar todos los inputs de tipo texto visibles y llenar los primeros dos vacíos
        if not fecha_desde_ok or not fecha_hasta_ok:
            self.log("   ⚠️ Fallback: buscando inputs de fecha por posición...")
            date_inputs = []
            for inp in await self.page.locator("input[type='text'], input:not([type])").all():
                try:
                    visible = await inp.is_visible()
                    if not visible:
                        continue
                    placeholder = (await inp.get_attribute("placeholder") or "").lower()
                    name = (await inp.get_attribute("name") or "").lower()
                    # Excluir inputs que son CUIT del agente u otros
                    if "cuit" in name or "cuit" in placeholder:
                        continue
                    date_inputs.append(inp)
                except Exception:
                    continue

            # Los primeros dos inputs de fecha del formulario = Ret/Perc Desde y Hasta
            if len(date_inputs) >= 1 and not fecha_desde_ok:
                try:
                    await date_inputs[0].triple_click(force=True)
                    await self.page.keyboard.type(fecha_desde_fmt, delay=40)
                    await self.page.keyboard.press("Tab")
                    self.log(f"   ✅ Fecha Desde (fallback pos[0]) = {fecha_desde_fmt}")
                except Exception as e:
                    self.log(f"   ⚠️ Fecha Desde fallback: {e}")
            if len(date_inputs) >= 2 and not fecha_hasta_ok:
                try:
                    await date_inputs[1].triple_click(force=True)
                    await self.page.keyboard.type(fecha_hasta_fmt, delay=40)
                    await self.page.keyboard.press("Tab")
                    self.log(f"   ✅ Fecha Hasta (fallback pos[1]) = {fecha_hasta_fmt}")
                except Exception as e:
                    self.log(f"   ⚠️ Fecha Hasta fallback: {e}")

        await self.page.wait_for_timeout(500)

        # ── Paso 7: CONSULTAR ────────────────────────────────────────────
        self.log("   🔍 Haciendo click en CONSULTAR...")
        consultado = False
        for sel in [
            "input[type='submit'][value='CONSULTAR']",
            "input[value='CONSULTAR']",
            "button:has-text('CONSULTAR')",
            "input[type='submit'][value*='onsultar']",
            "button:has-text('Consultar')",
            "input[value='Consultar']",
        ]:
            try:
                btn = self.page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(force=True)
                    consultado = True
                    self.log(f"   ✅ CONSULTAR ({sel})")
                    break
            except Exception:
                continue
        if not consultado:
            self.log("   ⚠️ Botón CONSULTAR no encontrado")
        await self.page.wait_for_timeout(6000)

        # ── Paso 8: Exportar a Excel ─────────────────────────────────────
        self.log("   📥 Buscando botón 'Exportar a Excel'...")
        output_dir = os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(output_dir, exist_ok=True)
        t_antes = _time.time()
        excel_path = None

        # Log de links/botones visibles para debug
        try:
            links = await self.page.locator("a, input[type='button'], input[type='submit'], button").all()
            textos = []
            for lnk in links[:30]:
                t = (await lnk.inner_text()).strip() or (await lnk.get_attribute("value") or "")
                if t:
                    textos.append(t.strip())
            self.log(f"   🔍 Elementos clickeables en página: {textos[:15]}")
        except Exception:
            pass

        # ── Encontrar el link "Exportar a Excel" iterando todos los <a> ────
        # El link tiene una imagen antes del texto ("X Exportar a Excel")
        # Playwright's has-text() a veces falla con imagen+texto, por eso iteramos.
        link_excel = None
        try:
            todos_los_links = await self.page.locator("a").all()
            for lnk in todos_los_links:
                try:
                    texto = (await lnk.inner_text()).strip().lower()
                    if "excel" in texto:
                        link_excel = lnk
                        self.log(f"   ✅ Link Excel encontrado: '{texto}'")
                        break
                except Exception:
                    continue
        except Exception as e:
            self.log(f"   ⚠️ Error buscando link Excel: {e}")

        if not link_excel:
            # Fallback: buscar por img alt o input value
            for css in ["img[alt*='Excel' i]", "input[value*='Excel' i]",
                        "input[value*='Exportar' i]"]:
                try:
                    elem = self.page.locator(css).first
                    if await elem.count() > 0:
                        # Subir al padre <a> si es imagen
                        parent = elem.locator("xpath=ancestor::a[1]")
                        if await parent.count() > 0:
                            link_excel = parent
                        else:
                            link_excel = elem
                        self.log(f"   ✅ Link Excel (fallback CSS): {css}")
                        break
                except Exception:
                    continue

        # ── Hacer click y capturar la descarga ───────────────────────────
        if link_excel:
            try:
                async with self.page.expect_download(timeout=30000) as dl_info:
                    await link_excel.click(force=True)
                dl = await dl_info.value
                fname = dl.suggested_filename or f"sicore_{int(t_antes)}.xls"
                excel_path = os.path.join(output_dir, fname)
                await dl.save_as(excel_path)
                self.log(f"   📥 Excel SICORE guardado: {excel_path}")
            except Exception as e:
                self.log(f"   ⚠️ Click/download falló: {e} — buscando en Descargas...")

        # ── Fallback: buscar archivo recién descargado en Descargas ──────
        if not excel_path or not os.path.exists(excel_path):
            await self.page.wait_for_timeout(6000)
            descargas = os.path.expandvars(r"%USERPROFILE%\Downloads")
            for pat in ["MisRetenciones*.xls*", "Retenciones*.xls*", "*.xls"]:
                for fp in sorted(glob.glob(os.path.join(descargas, pat)),
                                 key=os.path.getmtime, reverse=True):
                    if _time.time() - os.path.getmtime(fp) < 90:
                        excel_path = fp
                        self.log(f"   📥 Excel encontrado en Descargas: {excel_path}")
                        break
                if excel_path and os.path.exists(excel_path):
                    break

        if not excel_path or not os.path.exists(excel_path):
            self.log("   ⚠️ No se pudo descargar Excel — extrayendo de tabla HTML como fallback...")
            return await self._extraer_tabla_retenciones_html()

        # ── Paso 9: parsear el Excel ─────────────────────────────────────
        retenciones = self._parsear_excel_sicore(excel_path)
        self.log(f"   ✅ {len(retenciones)} retenciones SICORE del Excel")
        return retenciones

    def _parsear_excel_sicore(self, path: str) -> list[dict]:
        """
        Parsea el Excel .xls exportado por ARCA Mis Retenciones SICORE 767.
        Columnas del archivo:
          0: CUIT Agente Ret./Perc.
          1: Denominación o Razón Social
          6: Fecha Ret./Perc.
          9: Importe Ret./Perc.
          10: Número Comprobante
          11: Fecha Comprobante
        """
        def _m(s) -> float:
            try:
                # xlrd devuelve celdas numéricas como float nativo → no manipular
                if isinstance(s, (int, float)):
                    return float(s)
                s = str(s).strip().replace("$", "").replace(" ", "")
                if not s:
                    return 0.0
                # Formato con punto Y coma: 11.138,26 → 11138.26
                if "," in s and "." in s:
                    s = s.replace(".", "").replace(",", ".")
                elif "," in s:
                    # Solo coma como decimal: 11138,26 → 11138.26
                    s = s.replace(",", ".")
                elif "." in s:
                    partes = s.split(".")
                    # 1 o 2 dígitos después del punto → es decimal: 13114.6 o 11138.26
                    # 3 dígitos → separador de miles: 11.138 → 11138
                    if len(partes) == 2 and 1 <= len(partes[1]) <= 2:
                        pass
                    else:
                        # Punto como miles (o múltiples puntos): 11.138 → 11138
                        s = s.replace(".", "")
                return float(s)
            except Exception:
                return 0.0

        def _f(s) -> str:
            # Fechas numéricas de xlrd (número serial de Excel)
            try:
                n = float(str(s).strip())
                if 10000 < n < 100000:
                    try:
                        import xlrd as _xlrd
                        dt = _xlrd.xldate_as_datetime(n, 0)
                        return dt.strftime("%d/%m/%Y")
                    except Exception:
                        pass
            except Exception:
                pass
            s = str(s).strip()
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                try:
                    return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
                except Exception:
                    pass
            return s

        def _procesar_filas(filas: list) -> list[dict]:
            """Recibe lista de listas/tuplas, detecta encabezado y extrae datos."""
            col_map = {}
            header_found = False
            retenciones = []
            for row in filas:
                row = [str(c or "").strip() for c in row]
                joined = " ".join(row).lower()
                # Detectar fila de encabezado
                if not header_found and "cuit" in joined and "importe" in joined:
                    for j, t in enumerate(row):
                        tl = t.lower()
                        if "cuit agente" in tl or ("cuit" in tl and "agente" in tl):
                            col_map["cuit"] = j
                        elif "denominaci" in tl or "raz" in tl:
                            col_map["razon"] = j
                        elif "fecha ret" in tl:
                            col_map["fecha"] = j
                        elif "importe" in tl:
                            col_map["importe"] = j
                        elif "número comprobante" in tl or "numero comprobante" in tl:
                            col_map["nro_comp"] = j
                        elif "fecha comprobante" in tl:
                            col_map["fecha_comp"] = j
                    # Si no se mapeó nada, usar posiciones fijas conocidas
                    if not col_map:
                        col_map = {"cuit":0,"razon":1,"fecha":6,"importe":9,"nro_comp":10,"fecha_comp":11}
                    self.log(f"   📋 Columnas SICORE: {col_map}")
                    header_found = True
                    continue
                if not header_found:
                    continue
                cuit_ag = re.sub(r"[^0-9]", "", row[col_map.get("cuit", 0)] if row else "")
                if not cuit_ag or len(cuit_ag) < 8:
                    continue
                retenciones.append({
                    "fecha":               _f(row[col_map.get("fecha", 6)]   if len(row) > col_map.get("fecha",6) else ""),
                    "cuit_agente":         cuit_ag,
                    "razon_social_agente": row[col_map.get("razon", 1)]       if len(row) > col_map.get("razon",1) else "",
                    "tipo_retencion":      "IVA",
                    "numero_comprobante":  row[col_map.get("nro_comp", 10)]   if len(row) > col_map.get("nro_comp",10) else "",
                    "fecha_comprobante":   _f(row[col_map.get("fecha_comp",11)] if len(row) > col_map.get("fecha_comp",11) else ""),
                    "importe":             _m(row[col_map.get("importe", 9)] if len(row) > col_map.get("importe",9) else "0"),
                })
            return retenciones

        retenciones = []

        # ── Intentar con xlrd (nativo para .xls) ────────────────────────
        try:
            import xlrd
            wb = xlrd.open_workbook(path)
            ws = wb.sheet_by_index(0)
            filas = [ws.row_values(r) for r in range(ws.nrows)]
            self.log(f"   📂 xlrd: {len(filas)} filas leídas de '{os.path.basename(path)}'")
            retenciones = _procesar_filas(filas)
            if retenciones:
                return retenciones
        except ImportError:
            self.log("   ⚠️ xlrd no instalado — instalá con: pip install xlrd")
        except Exception as e:
            self.log(f"   ⚠️ xlrd falló: {e}")

        # ── Intentar con openpyxl (si es .xlsx) ─────────────────────────
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            filas = [list(row) for row in ws.iter_rows(values_only=True)]
            self.log(f"   📂 openpyxl: {len(filas)} filas")
            retenciones = _procesar_filas(filas)
            if retenciones:
                return retenciones
        except Exception as e:
            self.log(f"   ⚠️ openpyxl falló: {e}")

        # ── Intentar con pandas ──────────────────────────────────────────
        try:
            import pandas as pd
            df = pd.read_excel(path, dtype=str, header=None)
            filas = df.values.tolist()
            self.log(f"   📂 pandas: {len(filas)} filas")
            retenciones = _procesar_filas(filas)
            if retenciones:
                return retenciones
        except Exception as e:
            self.log(f"   ⚠️ pandas falló: {e}")

        self.log("   ❌ No se pudo leer el Excel SICORE con ningún método")
        return retenciones

    async def _extraer_tabla_retenciones_html(self) -> list[dict]:
        """Fallback: extrae retenciones de la tabla HTML si no hay Excel."""
        retenciones = []
        page_num = 1
        while True:
            self.log(f"   📃 Página {page_num} de retenciones HTML...")
            try:
                filas = await self.page.locator("table tbody tr").all()
                for fila in filas:
                    celdas = await fila.locator("td").all()
                    if len(celdas) < 3:
                        continue
                    textos = [await c.inner_text() for c in celdas]
                    ret = self._parsear_fila_retencion_html(textos)
                    if ret:
                        retenciones.append(ret)
            except Exception as e:
                self.log(f"   ⚠️ Error extrayendo tabla retenciones: {e}")
                break

            siguiente = self.page.locator(
                "a:has-text('Siguiente'), button:has-text('Siguiente')"
            ).first
            if await siguiente.count() == 0:
                break
            disabled = await siguiente.get_attribute("disabled")
            if disabled is not None:
                break
            await siguiente.click(force=True)
            await self.page.wait_for_timeout(2000)
            page_num += 1

        self.log(f"   ✅ {len(retenciones)} retenciones SICORE (HTML)")
        return retenciones

    def _parsear_fila_retencion_html(self, textos: list[str]) -> dict | None:
        """
        Parsea fila HTML del listado SICORE:
        CUIT Agente | Denominación | Impuesto | Régimen | Fecha | N°Cert | Descripción | Importe | N°Comp | FechaComp
        """
        if len(textos) < 3:
            return None

        def m(s):
            s = s.strip().replace(".", "").replace(",", ".").replace("$", "").replace(" ", "")
            try:
                return float(s)
            except Exception:
                return 0.0

        def f(s):
            s = s.strip()
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                try:
                    return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
                except Exception:
                    pass
            return s

        try:
            cuit_ag = re.sub(r"[^0-9]", "", textos[0])
            if not cuit_ag:
                return None
            return {
                "fecha":               f(textos[6]) if len(textos) > 6 else "",
                "cuit_agente":         cuit_ag,
                "razon_social_agente": textos[1].strip() if len(textos) > 1 else "",
                "tipo_retencion":      "IVA",
                "numero_comprobante":  textos[10].strip() if len(textos) > 10 else "",
                "fecha_comprobante":   f(textos[11]) if len(textos) > 11 else "",
                "importe":             m(textos[9]) if len(textos) > 9 else 0.0,
            }
        except Exception:
            return None
