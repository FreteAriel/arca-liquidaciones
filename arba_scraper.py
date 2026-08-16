"""
Scraper para ARBA (Agencia de Recaudación de la Provincia de Buenos Aires)
- Login con CUIT y clave fiscal (delegada desde AFIP)
- Deducciones informadas por agentes de recaudación
- Descarga de archivo por período
- Parseo de: Percepción, SITRAC, Retenciones Bancarias/PSP
"""

import asyncio
import os
import urllib.request


def _find_chrome():
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


def _get_system_proxy():
    try:
        proxies = urllib.request.getproxies()
        return proxies.get("https") or proxies.get("http") or None
    except Exception:
        return None
import re
import csv
import io
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError as PWTimeout


class ARBAScraper:
    LOGIN_URL = "https://web.arba.gov.ar/"

    def __init__(self, cuit: str, password: str, context=None, log_fn=print):
        self.cuit = re.sub(r"[^0-9]", "", cuit)
        self.password = password
        self.external_context = context  # Reutilizar contexto de ARCA si se comparte sesión
        self.log = log_fn
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.download_dir = os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(self.download_dir, exist_ok=True)
        self._content_frame = None  # iframe donde carga el contenido de Deducciones

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    async def start(self, headless: bool = False):
        if self.external_context:
            self.context = self.external_context
            self.page = await self.context.new_page()
            return

        self.playwright = await async_playwright().start()
        proxy = _get_system_proxy()
        proxy_cfg = {"server": proxy} if proxy else None

        extra_args = [
            "--start-maximized",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--safebrowsing-disable-download-protection",
            "--disable-popup-blocking",
        ]
        browser_path = _find_chrome()

        launch_kwargs = dict(
            headless=headless,
            args=extra_args,
            proxy=proxy_cfg,
            downloads_path=self.download_dir,
        )
        if browser_path:
            self.log(f"🌐 ARBA usando browser: {browser_path}")
            launch_kwargs["executable_path"] = browser_path

        try:
            self.browser = await self.playwright.chromium.launch(**launch_kwargs)
        except Exception:
            launch_kwargs.pop("executable_path", None)
            self.browser = await self.playwright.chromium.launch(**launch_kwargs)

        self.context = await self.browser.new_context(
            viewport={"width": 1280, "height": 900},
            accept_downloads=True,
        )
        self.page = await self.context.new_page()

    async def close(self):
        if not self.external_context and self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

    # ------------------------------------------------------------------
    # Login ARBA
    # ------------------------------------------------------------------

    async def login(self):
        """
        Login en web.arba.gov.ar usando Clave de Identificación Tributaria (CIT).
        Flujo: web.arba.gov.ar → Autogestión → completar CUIT + clave → Ingresar
        """
        self.log("🔐 Iniciando sesión en ARBA (web.arba.gov.ar)...")
        await self.page.goto(self.LOGIN_URL, wait_until="load", timeout=90000)
        await self.page.wait_for_timeout(3000)

        # Paso 1: click en "Ingresá" (botón teal en el panel Autogestión)
        # Puede ser <a>, <button>, <div> o cualquier elemento con ese texto
        clicked = False
        for txt in ["Ingresá", "Ingresar", "Autogestion", "Autogestión", "Acceder"]:
            for sel in [
                f"a:has-text('{txt}')",
                f"button:has-text('{txt}')",
                f"div:has-text('{txt}')",
                f"span:has-text('{txt}')",
                f"[class*='btn']:has-text('{txt}')",
                f"text={txt}",
            ]:
                try:
                    elem = self.page.locator(sel).first
                    if await elem.count() > 0:
                        await elem.click(force=True)
                        await self.page.wait_for_load_state("load")
                        await self.page.wait_for_timeout(2000)
                        self.log(f"   ✅ Click en '{txt}' ({sel})")
                        clicked = True
                        break
                except Exception:
                    continue
            if clicked:
                break

        # Fallback: JavaScript — busca el primer elemento visible que contenga "Ingresá"
        if not clicked:
            self.log("   ⚠️ Botón no encontrado por CSS — intentando JS...")
            try:
                await self.page.evaluate("""() => {
                    const textos = ['Ingresá', 'Ingresar', 'Acceder'];
                    for (const t of textos) {
                        const all = Array.from(document.querySelectorAll('*'));
                        for (const el of all) {
                            if (el.children.length === 0 && el.textContent.trim() === t) {
                                el.click();
                                return;
                            }
                        }
                    }
                }""")
                await self.page.wait_for_load_state("load")
                await self.page.wait_for_timeout(2000)
                self.log("   ✅ Click JS ejecutado")
            except Exception as e:
                self.log(f"   ⚠️ JS también falló: {e}")

        # Screenshot de debug para ver qué cargó después del click
        try:
            img = os.path.join(self.download_dir, "debug_arba_login.png")
            await self.page.screenshot(path=img)
            self.log(f"   📸 Screenshot: output/debug_arba_login.png")
        except Exception:
            pass

        # Paso 2: llenar CUIT (11 dígitos sin guiones)
        await self.page.wait_for_timeout(1500)
        cuit_filled = False
        for sel in [
            "input[placeholder*='11 dígitos']",
            "input[placeholder*='CUIT']",
            "input[placeholder*='C.U.I.T']",
            "input[name*='cuit' i]",
            "input[id*='cuit' i]",
            "input[type='text']",
        ]:
            try:
                elem = self.page.locator(sel).first
                if await elem.count() > 0:
                    await elem.fill(self.cuit)
                    cuit_filled = True
                    self.log(f"   ✅ CUIT ingresado ({sel})")
                    break
            except Exception:
                continue
        if not cuit_filled:
            raise Exception("No se encontró campo CUIT en ARBA")

        # Paso 3: llenar clave
        try:
            pwd = self.page.locator("input[type='password']").first
            if await pwd.count() > 0:
                await pwd.fill(self.password)
                self.log("   ✅ Clave ingresada")
        except Exception as e:
            self.log(f"   ⚠️ Campo clave: {e}")

        # Paso 4: click en "Ingresar"
        await self.page.wait_for_timeout(300)
        for sel in [
            "button:has-text('Ingresar')",
            "input[value='Ingresar']",
            "button[type='submit']",
            "input[type='submit']",
        ]:
            try:
                btn = self.page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(force=True)
                    break
            except Exception:
                continue

        await self.page.wait_for_load_state("load")
        await self.page.wait_for_timeout(4000)

        content = await self.page.content()
        if any(w in content.lower() for w in ["clave incorrecta", "cuit incorrecto", "error de autenticación"]):
            raise Exception("CUIT o clave incorrecta en ARBA")

        self.log("✅ Sesión iniciada en ARBA")

    # ------------------------------------------------------------------
    # Deducciones informadas por agentes de recaudación
    # ------------------------------------------------------------------

    async def get_deducciones(self, mes: int, anio: int) -> dict:
        self.log(f"📊 Obteniendo deducciones ARBA para {mes:02d}/{anio}...")

        # Flujo exacto: Ingresar → Mis accesos frecuentes → Deducciones informadas
        # → Deducciones → Período → Descargar para importar
        await self._navegar_deducciones(mes, anio)

        archivo = await self._esperar_descarga(mes, anio)
        if not archivo:
            self.log("⚠️ No se descargó archivo de deducciones ARBA")
            return {"percepciones": [], "sitrac": [], "retenciones_bancarias": []}

        return self._parsear_archivo_deducciones(archivo)

    async def _click_link(self, textos: list, esperar_carga: bool = True) -> bool:
        """Intenta hacer click en un link/botón por lista de textos posibles."""
        for txt in textos:
            for sel in [f"a:has-text('{txt}')", f"button:has-text('{txt}')",
                        f"span:has-text('{txt}')", f"li:has-text('{txt}')",
                        f"td:has-text('{txt}')"]:
                try:
                    elem = self.page.locator(sel).first
                    if await elem.count() > 0:
                        await elem.click(force=True)
                        if esperar_carga:
                            await self.page.wait_for_load_state("load")
                        await self.page.wait_for_timeout(2000)
                        self.log(f"   ✅ Click en '{txt}'")
                        return True
                except Exception:
                    continue
        return False

    async def _click_link_en_frame(self, frame, textos: list) -> bool:
        """Como _click_link pero opera dentro de un frame (iframe) específico."""
        for txt in textos:
            for sel in [f"a:has-text('{txt}')", f"button:has-text('{txt}')",
                        f"input[value='{txt}']", f"span:has-text('{txt}')",
                        f"li:has-text('{txt}')", f"td:has-text('{txt}')"]:
                try:
                    elem = frame.locator(sel).first
                    if await elem.count() > 0:
                        await elem.click(force=True)
                        await self.page.wait_for_timeout(2000)
                        self.log(f"   ✅ Click (frame) en '{txt}'")
                        return True
                except Exception:
                    continue
        return False

    async def _screenshot(self, nombre: str):
        """Guarda screenshot en output/ para debug."""
        try:
            path = os.path.join(self.download_dir, f"debug_arba_{nombre}.png")
            await self.page.screenshot(path=path, full_page=False)
            self.log(f"   📸 Screenshot: output/debug_arba_{nombre}.png")
        except Exception:
            pass

    async def _log_links_visibles(self):
        """Loguea todos los links/botones visibles y todos los frames disponibles."""
        try:
            self.log(f"   🌐 URL: {self.page.url}")
            # Loguear elementos en la página principal
            elems = await self.page.locator("a, button, li, span[onclick], td[onclick]").all()
            textos = []
            for e in elems[:40]:
                t = (await e.inner_text()).strip().replace("\n", " ")
                if t and len(t) < 80:
                    textos.append(t)
            self.log("   📋 Página principal: " + " | ".join(textos[:20]))
            # Loguear todos los frames y sus textos clave
            frames = self.page.frames
            self.log(f"   🖼️ Frames ({len(frames)}):")
            for i, f in enumerate(frames):
                try:
                    url = f.url
                    # Buscar textos clave en el frame
                    clave = await f.evaluate("""() => {
                        const textos = ['Deducciones', 'Descarga para Importar',
                                        'Régimen', 'Consultar', 'Descargar'];
                        return textos.filter(t =>
                            document.body && document.body.innerText.includes(t)
                        );
                    }""")
                    self.log(f"      [{i}] {url} — contiene: {clave}")
                except Exception as fe:
                    self.log(f"      [{i}] error: {fe}")
        except Exception as e:
            self.log(f"   ⚠️ _log_links_visibles error: {e}")

    async def _navegar_deducciones(self, mes: int, anio: int):
        """
        Flujo completo en ARBA:
        1. Después del login estamos en la home autenticada ("Hola, XXXX")
        2. Click "Ingresá" para entrar al panel de Autogestión
        3. Navegar a Mis Retenciones → Deducciones informadas
        4. Descarga para Importar → período → Consultar → Descargar
        """
        self.log("   🧭 Navegando a Deducciones en ARBA...")

        # Screenshot post-login: muestra home con "Ingresá a tu panel de Autogestión"
        await self._screenshot("01_post_login")
        await self._log_links_visibles()
        url_inicio = self.page.url
        self.log(f"   🌐 URL inicial: {url_inicio}")

        # ── Paso 0: Click en "Ingresá" para entrar al panel de Autogestión ──────
        # Después del login hay un botón teal "Ingresá" que lleva al panel real.
        self.log("   🖱️ Entrando al panel de Autogestión (click 'Ingresá')...")
        entro_panel = False
        # Intentar por texto exacto primero
        for sel in [
            "a:has-text('Ingresá')",
            "button:has-text('Ingresá')",
            "a:has-text('Ingresar')",
            "button:has-text('Ingresar')",
            "[class*='btn']:has-text('Ingresá')",
        ]:
            try:
                elem = self.page.locator(sel).first
                if await elem.count() > 0:
                    await elem.click(force=True)
                    await self.page.wait_for_timeout(4000)
                    self.log(f"   ✅ Click 'Ingresá' ({sel})")
                    entro_panel = True
                    break
            except Exception:
                continue

        # Fallback JS
        if not entro_panel:
            self.log("   ⚠️ No encontró 'Ingresá' por CSS — intentando JS...")
            try:
                await self.page.evaluate("""() => {
                    for (const t of ['Ingresá', 'Ingresar']) {
                        for (const el of document.querySelectorAll('a, button, div, span')) {
                            if (el.textContent.trim() === t) { el.click(); return; }
                        }
                    }
                }""")
                await self.page.wait_for_timeout(4000)
                self.log("   ✅ Click JS ejecutado")
                entro_panel = True
            except Exception as e:
                self.log(f"   ⚠️ JS también falló: {e}")

        await self._screenshot("02_panel_autogestion")
        await self._log_links_visibles()
        self.log(f"   🌐 URL tras click Ingresá: {self.page.url}")

        # ── Paso 1: "Mis accesos frecuentes" en el menú lateral ──────────────────
        ok = await self._click_link([
            "Mis accesos frecuentes",
            "Accesos frecuentes",
            "MIS ACCESOS FRECUENTES",
        ])
        if ok:
            await self.page.wait_for_timeout(2000)
            await self._screenshot("03_mis_accesos_frecuentes")
            await self._log_links_visibles()
        else:
            self.log("   ⚠️ No se encontró 'Mis accesos frecuentes'")
            await self._screenshot("03_mis_accesos_frecuentes")
            await self._log_links_visibles()

        # ── Paso 2: "Ingresos Brutos" para desplegar el submenú ──────────────────
        ok = await self._click_link([
            "Ingresos Brutos",
            "INGRESOS BRUTOS",
            "Ingresos brutos",
        ])
        if ok:
            await self.page.wait_for_timeout(2000)
            await self._screenshot("04_ingresos_brutos")
            await self._log_links_visibles()
        else:
            self.log("   ⚠️ No se encontró 'Ingresos Brutos'")
            await self._screenshot("04_ingresos_brutos")
            await self._log_links_visibles()

        # ── Paso 3: "Deducciones informadas" — puede abrir en nueva pestaña ────────
        # Registrar páginas abiertas ANTES del click para detectar nueva pestaña
        paginas_antes = list(self.context.pages)
        self.log(f"   📋 Pestañas antes del click: {len(paginas_antes)}")

        ok = await self._click_link([
            "Deducciones informadas por los Agentes de Recaudación",
            "Deducciones informadas por los agentes de recaudación",
            "Deducciones informadas",
            "Agentes de Recaudación",
            "Agentes de recaudación",
        ])
        if not ok:
            self.log("   ⚠️ No se encontró 'Deducciones informadas'")

        await self.page.wait_for_timeout(4000)

        # ── Detectar si se abrió en nueva pestaña ─────────────────────────────────
        paginas_despues = list(self.context.pages)
        nuevas = [p for p in paginas_despues if p not in paginas_antes]
        self.log(f"   📋 Pestañas después del click: {len(paginas_despues)} (nuevas: {len(nuevas)})")

        if nuevas:
            # "Deducciones informadas" abrió en nueva pestaña — cambiar foco
            nueva_pagina = nuevas[-1]
            await nueva_pagina.wait_for_load_state("load")
            self.log(f"   ✅ Nueva pestaña detectada: {nueva_pagina.url}")
            self.page = nueva_pagina  # IMPORTANTE: todas las ops posteriores sobre esta pestaña
        else:
            self.log(f"   📋 Sin nueva pestaña — URL actual: {self.page.url}")

        await self._screenshot("05_deducciones_informadas")
        await self._log_links_visibles()

        # ── Detectar iframe dentro de la página actual ─────────────────────────────
        # En algunos casos el contenido tabular carga en un <iframe> embebido.
        self.log("   🖼️ Verificando iframes en la página activa...")
        content_frame = self.page  # fallback: página principal

        # Esperar a que aparezca un <iframe> en el DOM (puede tardar si es AJAX)
        try:
            await self.page.wait_for_selector("iframe", timeout=4000)
            self.log("   ✅ iframe detectado en el DOM")
        except Exception:
            self.log("   ⚠️ No hay <iframe> en el DOM (o ya está listo)")

        await self.page.wait_for_timeout(1000)

        frames = self.page.frames
        self.log(f"   📋 Total frames: {len(frames)}")
        for i, f in enumerate(frames):
            try:
                self.log(f"      Frame {i}: {f.url}")
            except Exception:
                pass

        # Buscar el frame que contenga "Deducciones" como tab (texto exacto)
        for frame in frames:
            if frame == self.page.main_frame:
                continue
            try:
                has_tab = await frame.evaluate("""() => {
                    const els = Array.from(document.querySelectorAll('a, li, td, span'));
                    return els.some(el => el.textContent.trim() === 'Deducciones');
                }""")
                if has_tab:
                    content_frame = frame
                    self.log(f"   ✅ Frame con tab 'Deducciones' en: {frame.url}")
                    break
            except Exception as fe:
                self.log(f"      Frame error: {fe}")
                continue

        if content_frame is self.page:
            self.log("   ⚠️ No se encontró iframe con tab 'Deducciones' — usando página principal")

        self._content_frame = content_frame  # guardar para _esperar_descarga

        # ── Paso 4+5: Click en tab "Deducciones" → "Descarga para Importar" ────────
        # El tab usa CSS hover; estrategias dentro del content_frame detectado.
        self.log("   🖱️ Activando tab 'Deducciones' → 'Descarga para Importar'...")
        clicked_descarga = False

        # Intento A: extraer href del link oculto en el DOM del frame
        try:
            info = await content_frame.evaluate("""() => {
                for (const a of document.querySelectorAll('a')) {
                    const t = a.textContent.trim();
                    if (t.includes('Descarga') && t.includes('Importar')) {
                        return {href: a.href, onclick: a.getAttribute('onclick'), text: t};
                    }
                }
                for (const el of document.querySelectorAll('[onclick]')) {
                    const t = el.textContent.trim();
                    if (t.includes('Descarga') && t.includes('Importar')) {
                        return {href: null, onclick: el.getAttribute('onclick'), text: t};
                    }
                }
                return null;
            }""")
            if info:
                self.log(f"   🔗 DOM frame: text='{info.get('text')}' href='{info.get('href')}'")
                href = info.get('href', '') or ''
                onclick = info.get('onclick', '') or ''
                if href and href.startswith('http'):
                    await self.page.goto(href)
                    await self.page.wait_for_timeout(3000)
                    clicked_descarga = True
                    self.log(f"   ✅ Navegando directo a href")
                elif onclick:
                    await content_frame.evaluate(f"() => {{ {onclick} }}")
                    await self.page.wait_for_timeout(2000)
                    clicked_descarga = True
                    self.log(f"   ✅ Ejecutado onclick")
            else:
                self.log("   ⚠️ 'Descarga para Importar' no hallado en DOM del frame (submenu oculto)")
        except Exception as e:
            self.log(f"   ⚠️ Búsqueda DOM frame falló: {e}")

        # Intento B: JS fuerza visibilidad CSS del dropdown dentro del frame
        if not clicked_descarga:
            self.log("   🖱️ Forzando visibilidad CSS del dropdown en frame...")
            try:
                resultado = await content_frame.evaluate("""() => {
                    // Encontrar el tab "Deducciones"
                    const todos = Array.from(document.querySelectorAll('a, li, td, span, div'));
                    const ded = todos.find(el =>
                        el.children.length < 5 &&
                        el.textContent.trim() === 'Deducciones'
                    );
                    if (!ded) return 'NO_TAB_DEDUCCIONES';

                    // Disparar hover
                    ['mouseover','mouseenter','focus'].forEach(evt => {
                        ded.dispatchEvent(new MouseEvent(evt, {bubbles:true, cancelable:true}));
                    });

                    // Forzar visibilidad del contenedor padre + hijos
                    const parent = ded.closest('li') || ded.closest('td') || ded.parentElement;
                    if (parent) {
                        parent.querySelectorAll('*').forEach(el => {
                            const cs = window.getComputedStyle(el);
                            if (cs.display === 'none' || cs.visibility === 'hidden') {
                                el.style.display = 'block';
                                el.style.visibility = 'visible';
                                el.style.opacity = '1';
                            }
                        });
                    }

                    // Buscar y clickear "Descarga para Importar"
                    for (const el of document.querySelectorAll('a, li, td, span')) {
                        const t = el.textContent.trim();
                        if (t.includes('Descarga') && t.includes('Importar') && el.children.length < 3) {
                            el.click();
                            return 'CLICKED:' + t;
                        }
                    }
                    return 'NO_SUBMENU';
                }""")
                self.log(f"   📋 JS forzado resultado: {resultado}")
                if resultado and resultado.startswith('CLICKED:'):
                    clicked_descarga = True
                    await self.page.wait_for_timeout(2000)
                    self.log("   ✅ Click con CSS forzado en frame")
            except Exception as e:
                self.log(f"   ⚠️ JS forzado falló: {e}")

        # Intento C: Playwright hover real sobre el tab + force click en el subitem
        if not clicked_descarga:
            self.log("   🖱️ Fallback C: hover real + force click en frame...")
            try:
                for sel_ded in ["a:has-text('Deducciones')", "li:has-text('Deducciones')",
                                 "td:has-text('Deducciones')", "span:has-text('Deducciones')"]:
                    elems = content_frame.locator(sel_ded)
                    cnt = await elems.count()
                    for i in range(cnt):
                        elem = elems.nth(i)
                        box = await elem.bounding_box()
                        if box:
                            await self.page.mouse.move(
                                box['x'] + box['width'] / 2,
                                box['y'] + box['height'] / 2
                            )
                            await self.page.wait_for_timeout(800)
                            for sel_sub in ["a:has-text('Descarga para Importar')",
                                            "li:has-text('Descarga para Importar')"]:
                                sub = content_frame.locator(sel_sub).first
                                if await sub.count() > 0:
                                    await sub.click(force=True)
                                    self.log(f"   ✅ Force click '{sel_sub}'")
                                    clicked_descarga = True
                                    await self.page.wait_for_timeout(2000)
                                    break
                            if clicked_descarga:
                                break
                    if clicked_descarga:
                        break
            except Exception as e:
                self.log(f"   ⚠️ Fallback C falló: {e}")

        if not clicked_descarga:
            self.log("   ❌ No se pudo clickear 'Descarga para Importar' — continúa de todos modos")

        await self._screenshot("06_descarga_importar")
        await self._log_links_visibles()

        # ── Paso 6: Ingresar período (año + mes) en el frame ─────────────────────
        self.log(f"   📅 Ingresando período {mes:02d}/{anio}...")
        await self._seleccionar_periodo(mes, anio, frame=content_frame)
        await self.page.wait_for_timeout(800)
        await self._screenshot("08_periodo_ingresado")

        # ── Paso 7: Click "Consultar" en el frame ─────────────────────────────────
        ok = await self._click_link_en_frame(content_frame, ["Consultar", "CONSULTAR"])
        if not ok:
            for sel in ["input[type='submit']", "button[type='submit']"]:
                btn = content_frame.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(force=True)
                    self.log("   ✅ Submit por fallback (frame)")
                    break
        await self.page.wait_for_timeout(4000)
        await self._screenshot("09_resultado_consulta")
        self.log(f"   ✅ Consulta realizada — período {anio}{mes:02d}")

    async def _esperar_descarga(self, mes: int, anio: int) -> str | None:
        """
        Hace click en 'Descargar' e intercepta el ZIP descargado de ARBA.
        Guarda el ZIP en output/ y devuelve la ruta.
        Usa self._content_frame si está disponible (iframe de ARBA).
        """
        import time as _time
        self.log("   ⬇️ Buscando botón Descargar...")
        ctx = self._content_frame if self._content_frame is not None else self.page

        # Log elementos visibles para debug
        try:
            elems = await ctx.locator("a, button, input[type='submit'], input[type='button']").all()
            textos = [(await e.inner_text()).strip() or (await e.get_attribute("value") or "") for e in elems[:30]]
            textos = [t for t in textos if t]
            self.log(f"   🔍 Elementos en pantalla: {' | '.join(textos[:15])}")
        except Exception:
            pass

        btn = None
        for sel in [
            "button:has-text('Descargar')",
            "input[value='Descargar']",
            "a:has-text('Descargar')",
            "input[value*='escargar']",
            "button[type='submit']",
            "input[type='submit']",
        ]:
            try:
                elem = ctx.locator(sel).first
                if await elem.count() > 0:
                    btn = elem
                    self.log(f"   ✅ Botón Descargar: {sel}")
                    break
            except Exception:
                continue

        if not btn:
            self.log("   ⚠️ No se encontró botón Descargar")
            return None

        t_antes = _time.time()
        nombre_base = f"arba_deducciones_{anio}{mes:02d}"

        # Intentar capturar descarga con Playwright
        try:
            async with self.page.expect_download(timeout=30000) as dl_info:
                await btn.click(force=True)
            download = await dl_info.value
            ext = os.path.splitext(download.suggested_filename or "")[1] or ".zip"
            path = os.path.join(self.download_dir, f"{nombre_base}{ext}")
            await download.save_as(path)
            self.log(f"   ✅ Descargado: {os.path.basename(path)}")
            return path
        except Exception as e:
            self.log(f"   ⚠️ expect_download falló: {e}")

        # Fallback: buscar archivo reciente en output/ y Descargas
        import glob
        await self.page.wait_for_timeout(6000)
        for carpeta in [self.download_dir, os.path.expandvars(r"%USERPROFILE%\Downloads")]:
            for pat in ["*.zip", "*.ZIP", "*.txt", "*.TXT"]:
                for fp in sorted(glob.glob(os.path.join(carpeta, pat)),
                                 key=os.path.getmtime, reverse=True):
                    if os.path.getmtime(fp) > t_antes:
                        import shutil
                        ext = os.path.splitext(fp)[1]
                        dst = os.path.join(self.download_dir, f"{nombre_base}{ext}")
                        if fp != dst:
                            shutil.copy2(fp, dst)
                        self.log(f"   ✅ Archivo encontrado: {os.path.basename(fp)}")
                        return dst

        self.log("   ⚠️ No se encontró archivo descargado")
        return None

    async def _buscar_archivo_descargado(self, mes: int, anio: int, t_antes: float = 0.0) -> str | None:
        """Busca el archivo descargado en la carpeta Descargas del usuario."""
        import glob, time, shutil
        descargas = os.path.expandvars(r"%USERPROFILE%\Downloads")
        if not t_antes:
            t_antes = time.time() - 90  # últimos 90 segundos por defecto

        for pat in ["*.txt", "*.zip", "*.csv", "*.TXT", "*.ZIP", "*.CSV"]:
            archivos = glob.glob(os.path.join(descargas, pat))
            recientes = [f for f in archivos if os.path.getmtime(f) > t_antes]
            if recientes:
                src = max(recientes, key=os.path.getmtime)
                dst = os.path.join(self.download_dir, f"deducciones_{mes:02d}_{anio}.txt")
                shutil.copy2(src, dst)
                self.log(f"   ✅ Archivo encontrado en Descargas: {src}")
                return dst

        self.log("   ⚠️ No se encontró archivo descargado en Descargas")
        return None

    async def _seleccionar_periodo(self, mes: int, anio: int, frame=None):
        """
        Completa el campo de período en ARBA - "Consulta de deducciones".
        La pantalla muestra: [ YYYY ] - [ MM ] con dos inputs.
        Acepta un frame opcional (iframe) para buscar dentro del iframe correcto.
        """
        mes_str  = f"{mes:02d}"
        anio_str = str(anio)
        ctx = frame if frame is not None else self.page  # operar en el frame correcto

        # ── Intento 1: inputs con nombre explícito ────────────────────────
        anio_ok = False
        for sel in [
            "input[name*='anio' i]", "input[id*='anio' i]",
            "input[name*='year' i]", "input[id*='year' i]",
            "input[name*='periodo_anio' i]",
        ]:
            try:
                elem = ctx.locator(sel).first
                if await elem.count() > 0:
                    await elem.triple_click(force=True)
                    await elem.fill(anio_str)
                    anio_ok = True
                    self.log(f"   ✅ Año {anio_str} ({sel})")
                    break
            except Exception:
                continue

        mes_ok = False
        for sel in [
            "input[name*='mes' i]", "input[id*='mes' i]",
            "input[name*='month' i]", "input[id*='month' i]",
            "input[name*='periodo_mes' i]",
        ]:
            try:
                elem = ctx.locator(sel).first
                if await elem.count() > 0:
                    await elem.triple_click(force=True)
                    await elem.fill(mes_str)
                    mes_ok = True
                    self.log(f"   ✅ Mes {mes_str} ({sel})")
                    break
            except Exception:
                continue

        # ── Intento 2: campo único "período" YYYYMM o MMYYYY ─────────────
        if not anio_ok and not mes_ok:
            for sel in ["input[name*='periodo' i]", "input[id*='periodo' i]",
                        "input[name*='period' i]", "input[id*='period' i]"]:
                try:
                    elem = ctx.locator(sel).first
                    if await elem.count() > 0:
                        await elem.triple_click(force=True)
                        await elem.fill(f"{mes_str}{anio_str}")
                        self.log(f"   ✅ Período {mes_str}{anio_str} (campo único)")
                        anio_ok = mes_ok = True
                        break
                except Exception:
                    continue

        # ── Intento 3: select de mes ──────────────────────────────────────
        if not mes_ok:
            for sel in ["select[name*='mes' i]", "select[id*='mes' i]", "select"]:
                try:
                    selects = await ctx.locator(sel).all()
                    for s in selects:
                        opts = await s.locator("option").all()
                        for opt in opts:
                            v = (await opt.get_attribute("value") or "").strip()
                            t = (await opt.inner_text()).strip()
                            if v in (mes_str, str(mes)) or t in (mes_str, str(mes)):
                                await s.select_option(value=v)
                                mes_ok = True
                                self.log(f"   ✅ Mes {mes_str} (select)")
                                break
                        if mes_ok:
                            break
                    if mes_ok:
                        break
                except Exception:
                    continue

        # ── Intento 4: selector amplio (cualquier input que no sea hidden/submit) ───
        if not mes_ok or not anio_ok:
            self.log("   ⚠️ Intento 4: selector amplio de inputs...")
            try:
                sel_amplio = ("input:not([type='hidden']):not([type='submit'])"
                              ":not([type='button']):not([type='reset'])"
                              ":not([type='checkbox']):not([type='radio'])")
                inputs = await ctx.locator(sel_amplio).all()
                vacios = []
                for inp in inputs:
                    try:
                        val = await inp.input_value()
                        if not val or not val.strip():
                            vacios.append(inp)
                    except Exception:
                        continue
                self.log(f"   📋 Inputs vacíos (selector amplio): {len(vacios)}")
                if len(vacios) >= 2:
                    # En ARBA: primer campo = AÑO (YYYY), segundo = MES (MM)
                    await vacios[0].click(force=True)
                    await vacios[0].fill(anio_str)
                    await vacios[1].click(force=True)
                    await vacios[1].fill(mes_str)
                    anio_ok = mes_ok = True
                    self.log(f"   ✅ Período {anio_str}/{mes_str} (selector amplio)")
                elif len(vacios) == 1:
                    await vacios[0].click(force=True)
                    await vacios[0].fill(f"{mes_str}{anio_str}")
                    anio_ok = mes_ok = True
                    self.log(f"   ✅ Período {mes_str}{anio_str} (campo único amplio)")
            except Exception as e:
                self.log(f"   ⚠️ Intento 4 falló: {e}")

        # ── Intento 5 (FALLBACK JS): fill directo via DOM + dispatchEvent ─────────
        # Último recurso: llenar los inputs directamente desde JavaScript.
        # Primero loguea todos los inputs para diagnóstico, luego los rellena.
        if not mes_ok or not anio_ok:
            self.log("   ⚠️ Intento 5: JS fill directo + dispatchEvent...")
            try:
                resultado = await ctx.evaluate(f"""() => {{
                    const anio = '{anio_str}';
                    const mes  = '{mes_str}';
                    // Loguear todos los inputs
                    const todos = Array.from(document.querySelectorAll('input'));
                    const info = todos.map(i => ({{n:i.name, id:i.id, t:i.type, v:i.value, ml:i.maxLength}}));

                    // Filtrar inputs editables vacíos (excluir hidden, submit, etc.)
                    const excluir = ['hidden','submit','button','reset','checkbox','radio','file'];
                    const editables = todos.filter(i =>
                        !excluir.includes(i.type) && !i.disabled && !i.readOnly
                    );

                    function llenar(inp, valor) {{
                        // Soporte para frameworks que usan native input value setter
                        const nativeSetter = Object.getOwnPropertyDescriptor(
                            window.HTMLInputElement.prototype, 'value'
                        ).set;
                        nativeSetter.call(inp, valor);
                        inp.dispatchEvent(new Event('input',  {{bubbles:true}}));
                        inp.dispatchEvent(new Event('change', {{bubbles:true}}));
                        inp.dispatchEvent(new Event('blur',   {{bubbles:true}}));
                    }}

                    let anioFilled = null, mesFilled = null;
                    if (editables.length >= 2) {{
                        llenar(editables[0], anio);
                        llenar(editables[1], mes);
                        anioFilled = editables[0].value;
                        mesFilled  = editables[1].value;
                    }} else if (editables.length === 1) {{
                        llenar(editables[0], mes + anio);
                        anioFilled = mesFilled = editables[0].value;
                    }}

                    return {{
                        totalInputs: todos.length,
                        editables: editables.length,
                        anioFilled, mesFilled,
                        inputsInfo: info.slice(0, 10)
                    }};
                }}""")
                self.log(f"   📋 JS fill resultado: {resultado}")
                if resultado and resultado.get('anioFilled'):
                    anio_ok = mes_ok = True
                    self.log(f"   ✅ Período rellenado via JS: {resultado.get('anioFilled')}/{resultado.get('mesFilled')}")
            except Exception as e:
                self.log(f"   ⚠️ Intento 5 JS falló: {e}")

    def _nombre_mes(self, mes: int) -> str:
        meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                 "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
        return meses[mes - 1] if 1 <= mes <= 12 else str(mes)

    # ------------------------------------------------------------------
    # Parseo del archivo de deducciones (ZIP → TXT)
    # ------------------------------------------------------------------

    def _parsear_archivo_deducciones(self, ruta_archivo: str) -> dict:
        """
        Parsea el ZIP (o TXT) descargado de ARBA Descarga para Importar.
        El ZIP contiene varios TXT con este formato por línea:
          {prefix}{CUIT_con_guiones}{DD/MM/YYYY}{PV_5d}{numero_padded}{TIPO_2L}{importe_con_coma}
        Todos los registros van a ret_iibb (Ingresos Brutos).
        Incluye numero_formateado para matching por comprobante.
        """
        percepciones = []
        sitrac       = []
        retenciones_bancarias = []

        if not ruta_archivo or not os.path.exists(ruta_archivo):
            return {"percepciones": percepciones, "sitrac": sitrac,
                    "retenciones_bancarias": retenciones_bancarias}

        self.log(f"   🔍 Parseando archivo ARBA: {os.path.basename(ruta_archivo)}")

        import zipfile

        def _procesar_txt(contenido: str, nombre_archivo: str = ""):
            """Parsea el contenido de un TXT ARBA y devuelve lista de registros."""
            registros = []
            for linea in contenido.splitlines():
                reg = self._parsear_linea_txt_arba(linea.strip())
                if reg:
                    registros.append(reg)
            return registros

        todos = []
        try:
            if zipfile.is_zipfile(ruta_archivo):
                with zipfile.ZipFile(ruta_archivo) as zf:
                    self.log(f"   📦 ZIP con {len(zf.namelist())} archivos: {zf.namelist()}")
                    for nombre in zf.namelist():
                        try:
                            raw = zf.read(nombre)
                            try:
                                contenido = raw.decode("latin-1")
                            except Exception:
                                contenido = raw.decode("utf-8", errors="replace")
                            regs = _procesar_txt(contenido, nombre)
                            self.log(f"   📄 {nombre}: {len(regs)} registros")
                            todos.extend(regs)
                        except Exception as e:
                            self.log(f"   ⚠️ Error en {nombre}: {e}")
            else:
                # Archivo TXT directo
                try:
                    with open(ruta_archivo, "r", encoding="latin-1", errors="replace") as f:
                        contenido = f.read()
                except Exception:
                    with open(ruta_archivo, "r", encoding="utf-8", errors="replace") as f:
                        contenido = f.read()
                todos = _procesar_txt(contenido, os.path.basename(ruta_archivo))

        except Exception as e:
            self.log(f"   ⚠️ Error parseando archivo ARBA: {e}")

        # Separar por tipo de registro
        for reg in todos:
            t = reg.get("tipo", "")
            if t == "PERCEPCION":
                percepciones.append(reg)
            elif t == "SITRAC":
                sitrac.append(reg)
            elif t == "BANCARIA":
                retenciones_bancarias.append(reg)
            else:
                percepciones.append(reg)  # fallback

        self.log(f"   ✅ ARBA: {len(percepciones)} CP / {len(sitrac)} CT (SITRAC) / {len(retenciones_bancarias)} CB (Bancarias)")

        return {
            "percepciones": percepciones,
            "sitrac": sitrac,
            "retenciones_bancarias": retenciones_bancarias,
        }

    def _parsear_linea_txt_arba(self, linea: str) -> dict | None:
        """
        Parsea una línea de los TXT ARBA 'Descarga para Importar'.
        El ZIP contiene 3 tipos de archivo con formatos distintos:

        CP (Percepciones):
          {prefix}{CUIT-XX-XXXXXXXX-X}{DD/MM/YYYY}{PV_5d}{numero_padded}{TIPO_1-2L}{-?importe}
          Ejemplo: 0290230-71796465-503/06/20260000700000000000000001625FA000000005651,43
          → cuit=30717964655, pv=7, num=1625, num_fmt="00007-00001625", imp=5651.43

        CT (SITRAC - Convenio Multilateral):
          Mismo CUIT + DD/MM/YYYY, pero el importe es los últimos 11 chars fijos.
          Ejemplo: 90230-70308853-415/06/202600010000161442290435O 0000000016144229043500000660,00
          → cuit=30703088534, imp=660.00 (sin numero_formateado para matching exacto)

        CB (Retenciones Bancarias):
          {prefix}{CUIT-XX-XXXXXXXX-X}{YYYY/MM}...{TIPO_2-3L}{importe}
          Ejemplo: 0433-99924210-92026/060140152903700560584859CAP000000000096,00
          → cuit=33999242109, imp=96.00
        """
        if not linea or len(linea) < 20:
            return None

        # ── Formato 1: CP (Percepciones IIBB) ────────────────────────────────
        # Identificable porque tiene: CUIT + DD/MM/YYYY + PV(5d) + num + tipo(1-2L) + importe(≤15d)
        m = re.search(
            r'(\d{2}-\d{8}-\d)'        # CUIT con guiones
            r'(\d{2}/\d{2}/\d{4})'     # fecha DD/MM/YYYY
            r'(\d{5})'                  # punto de venta (5 dígitos)
            r'(\d+)'                    # número comprobante (padded)
            r'([A-Z]{1,2})\s*'          # tipo doc (1 o 2 letras) + espacio opcional
            r'(-?\d{1,15},\d{2})',      # importe: máx 15 dígitos antes de la coma
            linea,
        )
        if m:
            try:
                cuit_agente       = re.sub(r'[^0-9]', '', m.group(1))
                pv                = int(m.group(3))
                numero            = int(m.group(4))
                numero_formateado = f"{pv:05d}-{numero:08d}"
                importe           = float(m.group(6).replace(',', '.'))
                return {
                    "tipo":              "PERCEPCION",
                    "cuit_agente":       cuit_agente,
                    "fecha":             m.group(2),
                    "punto_venta":       pv,
                    "numero":            numero,
                    "numero_formateado": numero_formateado,
                    "tipo_doc":          m.group(5).strip(),
                    "importe":           importe,
                }
            except Exception:
                pass

        # ── Formato 2: CT (SITRAC - Convenio Multilateral) ───────────────────
        # Misma fecha DD/MM/YYYY pero tipo de 1 letra + espacio, importe al final (11 chars fijos).
        # El CP regex falla en CT porque los dígitos antes de la coma son > 15.
        m_ct = re.search(r'(\d{2}-\d{8}-\d)(\d{2}/\d{2}/\d{4})', linea)
        if m_ct:
            try:
                tail    = linea.strip()
                imp_str = tail[-11:]   # últimos 11 chars: "XXXXXXXX,XX"
                if re.match(r'\d+,\d{2}$', imp_str):
                    importe = float(imp_str.replace(',', '.'))
                    return {
                        "tipo":              "SITRAC",
                        "cuit_agente":       re.sub(r'[^0-9]', '', m_ct.group(1)),
                        "fecha":             m_ct.group(2),
                        "punto_venta":       0,
                        "numero":            0,
                        "numero_formateado": "",
                        "tipo_doc":          "",
                        "importe":           importe,
                    }
            except Exception:
                pass

        # ── Formato 3: CB (Retenciones Bancarias / PSP) ──────────────────────
        # Fecha en formato YYYY/MM (no DD/MM/YYYY), tipo de 2-3 letras (CAP, CCP, etc.)
        m_cb = re.search(
            r'(\d{2}-\d{8}-\d)'        # CUIT con guiones
            r'(\d{4}/\d{2})'           # fecha YYYY/MM
            r'.*?'                      # contenido variable
            r'([A-Z]{2,3})'            # tipo (CAP, CCP, ...)
            r'(-?\d{1,15},\d{2})\s*$', # importe
            linea,
        )
        if m_cb:
            try:
                importe = float(m_cb.group(4).replace(',', '.'))
                return {
                    "tipo":              "BANCARIA",
                    "cuit_agente":       re.sub(r'[^0-9]', '', m_cb.group(1)),
                    "fecha":             m_cb.group(2),
                    "punto_venta":       0,
                    "numero":            0,
                    "numero_formateado": "",
                    "tipo_doc":          m_cb.group(3),
                    "importe":           importe,
                }
            except Exception:
                pass

        return None
