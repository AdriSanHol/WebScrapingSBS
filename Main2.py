import os
import calendar
from datetime import date, timedelta
from playwright.sync_api import sync_playwright
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================
# CONFIG
# =========================
URL = "https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIActivaTipoCreditoEmpresa.aspx?tip=B"
CHROME_PATH = r"C:\chromium\chrome-win64\chrome.exe"

# =========================
# 📅 FECHAS SBS
# =========================
def ayer_habil():
    d = date.today() - timedelta(days=1)
    if d.weekday() == 5:      # sábado
        d -= timedelta(days=1)
    elif d.weekday() == 6:    # domingo
        d -= timedelta(days=2)
    return d

def ultimo_habil_mes(anio, mes):
    d = date(anio, mes, calendar.monthrange(anio, mes)[1])
    if d.weekday() == 5:
        d -= timedelta(days=1)
    elif d.weekday() == 6:
        d -= timedelta(days=2)
    return d

def primer_habil_mes(anio, mes):
    d = date(anio, mes, 1)
    if d.weekday() == 5:      # sábado
        d += timedelta(days=2)
    elif d.weekday() == 6:    # domingo
        d += timedelta(days=1)
    return d

# =========================
# 📂 CARPETAS
# =========================
def crear_ruta(base, anio, mes):
    nombre_mes = f"{mes:02d}_{calendar.month_name[mes]}"
    ruta = os.path.join(base, str(anio), nombre_mes)
    os.makedirs(ruta, exist_ok=True)
    return ruta

# =========================
# 🤖 DESCARGA DIARIA
# =========================
def descargar_diario(ruta_base):
    hoy = date.today()
    limite = ayer_habil()

    anio = hoy.year
    mes = hoy.month

    if limite.month != mes:
        return  # nada que correr aún

    ruta_mes = crear_ruta(ruta_base, anio, mes)
    inicio = primer_habil_mes(anio, mes)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            executable_path=CHROME_PATH,
            headless=False
        )
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=60000)

        fecha = inicio
        while fecha <= limite:

            if fecha.weekday() >= 5:
                fecha += timedelta(days=1)
                continue

            fecha_sbs = fecha.strftime("%d/%m/%Y")
            fecha_arch = fecha.strftime("%Y-%m-%d")

            mn = os.path.join(ruta_mes, f"MN_{fecha_arch}.xlsx")
            me = os.path.join(ruta_mes, f"ME_{fecha_arch}.xlsx")

            if os.path.exists(mn) and os.path.exists(me):
                fecha += timedelta(days=1)
                continue

            print(f"📅 Diario: {fecha_arch}")

            page.locator("input[id*='dateInput']").first.fill(fecha_sbs)
            page.keyboard.press("Enter")
            page.wait_for_timeout(4000)

            if not os.path.exists(mn):
                with page.expect_download() as d:
                    page.locator("#ctl00_cphContent_btnExportar").click()
                d.value.save_as(mn)

            page.locator("#ctl00_cphContent_lbtnMex").click()
            page.wait_for_timeout(3000)

            if not os.path.exists(me):
                with page.expect_download() as d:
                    page.locator("#ctl00_cphContent_btnExportar").click()
                d.value.save_as(me)

            fecha += timedelta(days=1)

        browser.close()

# =========================
# 🤖 DESCARGA MENSUAL
# =========================
def descargar_mensual(ruta_base, anio_ini, mes_ini, anio_fin, mes_fin):
    hoy = date.today()
    ayer = ayer_habil()

    with sync_playwright() as p:
        browser = p.chromium.launch(
            executable_path=CHROME_PATH,
            headless=False
        )
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=60000)

        anio, mes = anio_ini, mes_ini

        while (anio < anio_fin) or (anio == anio_fin and mes <= mes_fin):

            if anio == hoy.year and mes == hoy.month:
                fecha = ayer
            else:
                fecha = ultimo_habil_mes(anio, mes)

            ruta_mes = crear_ruta(ruta_base, anio, mes)

            fecha_sbs = fecha.strftime("%d/%m/%Y")
            fecha_arch = fecha.strftime("%Y-%m-%d")

            print(f"📅 Mensual: {anio}-{mes:02d} → {fecha_arch}")

            page.locator("input[id*='dateInput']").first.fill(fecha_sbs)
            page.keyboard.press("Enter")
            page.wait_for_timeout(4000)

            with page.expect_download() as d:
                page.locator("#ctl00_cphContent_btnExportar").click()
            d.value.save_as(os.path.join(ruta_mes, f"MN_{fecha_arch}.xlsx"))

            page.locator("#ctl00_cphContent_lbtnMex").click()
            page.wait_for_timeout(3000)

            with page.expect_download() as d:
                page.locator("#ctl00_cphContent_btnExportar").click()
            d.value.save_as(os.path.join(ruta_mes, f"ME_{fecha_arch}.xlsx"))

            mes += 1
            if mes > 12:
                mes = 1
                anio += 1

        browser.close()

# =========================
# 🖥️ FRONT
# =========================
def ejecutar():
    ruta = filedialog.askdirectory(title="Selecciona carpeta destino")
    if not ruta:
        return

    try:
        if modo.get() == "diario":
            descargar_diario(ruta)
        else:
            descargar_mensual(
                ruta,
                int(anio_ini.get()), int(mes_ini.get()),
                int(anio_fin.get()), int(mes_fin.get())
            )

        messagebox.showinfo("Listo", "Proceso completado correctamente ✅")
    except Exception as e:
        messagebox.showerror("Error", str(e))

app = tk.Tk()
app.title("Descarga Tasas SBS")

modo = tk.StringVar(value="mensual")

tk.Radiobutton(app, text="Mensual", variable=modo, value="mensual").grid(row=0, column=0)
tk.Radiobutton(app, text="Diario (día hábil)", variable=modo, value="diario").grid(row=0, column=1, columnspan=2)

tk.Label(app, text="Año inicio").grid(row=1, column=0)
tk.Label(app, text="Mes inicio").grid(row=1, column=1)
tk.Label(app, text="Año fin").grid(row=1, column=2)
tk.Label(app, text="Mes fin").grid(row=1, column=3)

anio_ini = tk.Entry(app, width=6); anio_ini.insert(0, "2025")
mes_ini = tk.Entry(app, width=4); mes_ini.insert(0, "1")
anio_fin = tk.Entry(app, width=6); anio_fin.insert(0, "2025")
mes_fin = tk.Entry(app, width=4); mes_fin.insert(0, "6")

anio_ini.grid(row=2, column=0)
mes_ini.grid(row=2, column=1)
anio_fin.grid(row=2, column=2)
mes_fin.grid(row=2, column=3)

tk.Button(
    app,
    text="Seleccionar carpeta y ejecutar",
    command=ejecutar,
    bg="#4CAF50",
    fg="white"
).grid(row=3, column=0, columnspan=4, pady=10)

app.mainloop()