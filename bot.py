"""Travel Advisor Bot - Lucila Barros"""
import logging, os, asyncio, json
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from scraper import scrape_flights, scrape_packages, get_usd_rate
from report_generator import generate_daily_report
from database import init_db
from dotenv import load_dotenv

load_dotenv()

TOKEN   = os.getenv("TELEGRAM_VIAJES_TOKEN") or os.getenv("TELEGRAM_TOKEN", "8784357296:AAHifcn_XSOovS08VS3j6DnxJHA58TOQnyw")
CHAT_ID = os.getenv("CHAT_ID", "8700942418")
ALERTA_PP = 300
HIST_FILE = "/tmp/historial.json"

logging.basicConfig(format="%(asctime)s [%(levelname)s] %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

FECHAS = [("2026-01-03","2026-01-09"),("2026-01-10","2026-01-16"),
          ("2026-01-17","2026-01-23"),("2026-01-24","2026-01-30")]

def cargar_hist():
    return json.load(open(HIST_FILE)) if os.path.exists(HIST_FILE) else {}

def guardar_hist(h):
    json.dump(h, open(HIST_FILE,"w"), indent=2)

def alertas(vuelos, hist, fecha):
    out = []
    pp = [v["price_usd"] for v in vuelos if v.get("price_usd")]
    if not pp: return out
    minp = min(pp)
    clave = f"min_{fecha.replace('-','')}"
    prev = hist.get(clave)
    if prev and minp < prev:
        out.append(f"BAJA DE PRECIO {fecha}!\nAntes: USD {prev:.0f} -> Ahora: USD {minp:.0f}\nAhorras USD {(prev-minp)*5:.0f} para 5 personas")
    if minp <= ALERTA_PP:
        m = min(vuelos, key=lambda v: v.get("price_usd",9999))
        out.append(f"OFERTA! Menos de USD {ALERTA_PP}/pp\n{fecha} | {m.get('airline','')} | USD {minp:.0f}/pp | Total x5: USD {minp*5:.0f}")
    hist[clave] = minp
    return out

async def scrape_all(fechas=None):
    loop = asyncio.get_event_loop()
    vuelos = []
    for fi, fv in (fechas or FECHAS):
        try:
            v = await loop.run_in_executor(None, scrape_flights, "EZE", "GIG", fi, fv, 5)
            vuelos.extend(v)
        except Exception as e:
            log.error(f"Scrape error {fi}: {e}")
        try:
            await loop.run_in_executor(None, scrape_packages, "EZE", "GIG", fi, fv, 5)
        except Exception as e:
            log.error(f"Pkg error {fi}: {e}")
    return vuelos

async def tarea_diaria(app=None):
    log.info("Iniciando tarea diaria...")
    hist = cargar_hist()
    todos = await scrape_all()
    for fi, _ in FECHAS:
        vv = [v for v in todos if v.get("departure_date") == fi]
        for a in alertas(vv, hist, fi):
            if app:
                try: await app.bot.send_message(chat_id=CHAT_ID, text=a, parse_mode="HTML")
                except Exception as e: log.error(e)
    guardar_hist(hist)

    excel, msg = generate_daily_report()
    if not excel:
        log.warning("Sin datos para reporte.")
        return
    if app:
        try:
            await app.bot.send_message(chat_id=CHAT_ID, text=msg, parse_mode="HTML")
            with open(excel, "rb") as f:
                await app.bot.send_document(
                    chat_id=CHAT_ID, document=f,
                    filename=f"vuelos_buzios_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    caption="Excel PRO: Resumen - Vuelos - Paquetes - Historico - Por Fecha")
            log.info("Reporte PRO enviado.")
        except Exception as e:
            log.error(f"Error enviando: {e}")

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"Hola Lucila! Bot de viajes a Buzios activo.\n\n"
        f"/reporte - Excel PRO ahora mismo\n"
        f"/buscar - Busqueda completa\n"
        f"/alerta - Ver historico de precios\n"
        f"Reporte automatico todos los dias a las 21hs")

async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cid = update.effective_chat.id
    await update.message.reply_text("Preparando tu reporte PRO...")
    excel, msg = generate_daily_report()
    if not excel:
        await context.bot.send_message(cid, "Sin datos aun. Buscando... (2-3 min)")
        try:
            await scrape_all(FECHAS[:2])
            excel, msg = generate_daily_report()
        except Exception as e:
            log.error(e)
    if excel and msg:
        try:
            await context.bot.send_message(cid, msg, parse_mode="HTML")
            with open(excel, "rb") as f:
                await context.bot.send_document(
                    cid, document=f,
                    filename=f"vuelos_buzios_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                    caption="Excel PRO: Resumen - Vuelos - Paquetes - Historico - Por Fecha")
        except Exception as e:
            await context.bot.send_message(cid, f"Error: {e}")
    else:
        await context.bot.send_message(cid, "No se pudo obtener datos. Proba con /buscar primero.")

async def cmd_buscar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Buscando vuelos... tarda 2-3 min.")
    try:
        await tarea_diaria(app=context.application)
    except Exception as e:
        await update.message.reply_text(f"Error: {e}")

async def cmd_alerta(update: Update, context: ContextTypes.DEFAULT_TYPE):
    hist = cargar_hist()
    if not hist:
        await update.message.reply_text("Sin historial aun. Usa /buscar primero.")
        return
    lines = ["Minimos de precio registrados:\n"]
    for k, p in sorted(hist.items()):
        f = k.replace("min_","")
        f2 = f"{f[6:8]}/{f[4:6]}/{f[:4]}"
        lines.append(f"{'OFERTA' if p<=ALERTA_PP else 'Normal'} | {f2}: USD {p:.0f}/pp | x5: USD {p*5:.0f}")
    await update.message.reply_text("\n".join(lines))

def main():
    init_db()
    log.info("Bot iniciado en Railway")
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("reporte", cmd_reporte))
    app.add_handler(CommandHandler("buscar", cmd_buscar))
    app.add_handler(CommandHandler("buscar_ahora", cmd_buscar))
    app.add_handler(CommandHandler("alerta", cmd_alerta))
    sched = AsyncIOScheduler(timezone="America/Argentina/Buenos_Aires")
    sched.add_job(tarea_diaria, "cron", hour=21, minute=0, kwargs={"app":app})
    sched.add_job(tarea_diaria, "cron", hour="3,9,15", minute=0, kwargs={"app":app})
    sched.start()
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
