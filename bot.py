"""
✈️  TRAVEL ADVISOR BOT — Lucila Barros
Buenos Aires → Búzios / Río de Janeiro | Enero 2026 | 5 pasajeras
Scraping real con Playwright + Excel prolijo + Alertas Telegram 24/7
"""

import logging
import os
import asyncio
import json
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from scraper import scrape_flights, scrape_packages, get_usd_rate
from report_generator import generate_daily_report
from database import init_db
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_TOKEN   = os.getenv("TELEGRAM_TOKEN", "8784357296:AAHifcn_XSOovS08VS3j6DnxJHA58TOQnyw")
CHAT_ID          = os.getenv("CHAT_ID", "8700942418")
PRECIO_ALERTA_PP = 300
HISTORIAL_FILE   = "/tmp/historial_precios.json"

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%d/%m/%Y %H:%M:%S",
    level=logging.INFO
)
log = logging.getLogger(__name__)

ENERO_FECHAS = [
    ("2026-01-03", "2026-01-09"),
    ("2026-01-10", "2026-01-16"),
    ("2026-01-17", "2026-01-23"),
    ("2026-01-24", "2026-01-30"),
]

# ─────────────────────────────────────────
# HISTORIAL Y ALERTAS
# ─────────────────────────────────────────
def cargar_historial():
    if os.path.exists(HISTORIAL_FILE):
        with open(HISTORIAL_FILE) as f:
            return json.load(f)
    return {}

def guardar_historial(h):
    with open(HISTORIAL_FILE, "w") as f:
        json.dump(h, f, indent=2)

def detectar_alertas(vuelos, historial, fecha_ida):
    alertas = []
    if not vuelos:
        return alertas
    precios = [v["price_usd"] for v in vuelos if v.get("price_usd")]
    if not precios:
        return alertas
    precio_actual   = min(precios)
    clave           = f"min_pp_{fecha_ida.replace('-','')}"
    precio_anterior = historial.get(clave)
    if precio_anterior and precio_actual < precio_anterior:
        caida = round(precio_anterior - precio_actual, 0)
        alertas.append(
            f"📉 *¡BAJÓ el precio para {fecha_ida}!*\n"
            f"Antes: USD {precio_anterior:.0f} → Ahora: USD {precio_actual:.0f}\n"
            f"Ahorrás USD {caida:.0f}/pp = USD {caida*5:.0f} para 5 personas"
        )
    if precio_actual <= PRECIO_ALERTA_PP:
        mejor = min(vuelos, key=lambda v: v.get("price_usd", 9999))
        alertas.append(
            f"🔥 *¡OFERTA! Precio bajo USD {PRECIO_ALERTA_PP}/pp*\n"
            f"Fecha: {fecha_ida}\n"
            f"Aerolínea: {mejor.get('airline','N/D')} | USD {precio_actual:.0f}/pp\n"
            f"Total x5: USD {precio_actual*5:.0f}\n"
            f"🔗 {mejor.get('url','')}"
        )
    historial[clave] = precio_actual
    return alertas

# ─────────────────────────────────────────
# SCRAPING EN BACKGROUND (no bloquea el bot)
# ─────────────────────────────────────────
async def scrape_todas_las_fechas(fechas=None):
    """Corre el scraping de todas las fechas en un thread separado."""
    if fechas is None:
        fechas = ENERO_FECHAS
    loop = asyncio.get_event_loop()
    todos_vuelos = []
    for fecha_ida, fecha_vuelta in fechas:
        log.info(f"Scraping {fecha_ida} → {fecha_vuelta}...")
        try:
            vuelos = await loop.run_in_executor(
                None, scrape_flights, "EZE", "GIG", fecha_ida, fecha_vuelta, 5
            )
            todos_vuelos.extend(vuelos)
        except Exception as e:
            log.error(f"Error scraping vuelos {fecha_ida}: {e}")
        try:
            await loop.run_in_executor(
                None, scrape_packages, "EZE", "GIG", fecha_ida, fecha_vuelta, 5
            )
        except Exception as e:
            log.error(f"Error scraping paquetes {fecha_ida}: {e}")
    return todos_vuelos

# ─────────────────────────────────────────
# TAREA DIARIA
# ─────────────────────────────────────────
async def run_daily_task(app=None):
    log.info("Iniciando tarea diaria...")
    historial = cargar_historial()
    target_id = CHAT_ID

    # Scraping completo
    todos_vuelos = await scrape_todas_las_fechas()

    # Detectar alertas fecha por fecha
    for fecha_ida, _ in ENERO_FECHAS:
        vuelos_fecha = [v for v in todos_vuelos if v.get("departure_date") == fecha_ida]
        alertas = detectar_alertas(vuelos_fecha, historial, fecha_ida)
        if app and alertas:
            for alerta in alertas:
                try:
                    await app.bot.send_message(chat_id=target_id, text=alerta, parse_mode="Markdown")
                except Exception as e:
                    log.error(f"Error enviando alerta: {e}")

    guardar_historial(historial)

    # Generar y enviar Excel
    report_file = generate_daily_report()
    if not report_file:
        log.warning("Sin datos para generar reporte diario.")
        return

    usd_rate = get_usd_rate()
    tips = (
        f"💡 *Tips de tu Asesor de Viajes*\n"
        f"_(Dólar blue hoy: ${usd_rate:.0f})_\n\n"
        f"✅ *LATAM*: Mejor relación precio-servicio. Equipaje incluido.\n"
        f"💰 *Gol*: La más barata pero el equipaje se paga aparte.\n"
        f"⚠️ *Aerolíneas Arg.*: Sale de Aeroparque pero historial de demoras.\n"
        f"🔄 *Copa*: Confiable, escala en Panamá = más tiempo de viaje.\n\n"
        f"📊 Excel adjunto con todos los detalles y links directos para comprar."
    )

    if app:
        try:
            await app.bot.send_message(chat_id=target_id, text=tips, parse_mode="Markdown")
            with open(report_file, "rb") as f:
                await app.bot.send_document(
                    chat_id=target_id,
                    document=f,
                    filename=f"informe_vuelos_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    caption=f"📊 *Reporte diario* — {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                    parse_mode="Markdown"
                )
            log.info("Reporte diario enviado.")
        except Exception as e:
            log.error(f"Error enviando reporte diario: {e}")

# ─────────────────────────────────────────
# COMANDOS TELEGRAM
# ─────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_chat.id
    await update.message.reply_text(
        f"👋 *¡Hola Lucila!* Soy tu bot de viajes a Búzios 🇧🇷\n\n"
        f"Tu chat ID: `{uid}`\n\n"
        f"*Comandos disponibles:*\n"
        f"📊 /reporte → Excel con todos los vuelos ahora mismo\n"
        f"🔍 /buscar → Búsqueda completa + tips del asesor\n"
        f"📈 /alerta → Ver mínimos de precio registrados\n"
        f"📩 Reporte automático todos los días a las *21hs*\n"
        f"🔔 Alertas cuando el precio baja de USD {PRECIO_ALERTA_PP}/pp",
        parse_mode="Markdown"
    )

async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Genera y envía el Excel ahora mismo.
    Si la DB está vacía, scrapea automáticamente antes de generar.
    """
    chat_id = update.effective_chat.id
    await update.message.reply_text("⏳ Preparando tu reporte...")

    # Primero intentar con datos ya en DB
    report_file = generate_daily_report()

    # Si no hay datos, scrapear 2 semanas rápido y reintentar
    if not report_file:
        await context.bot.send_message(
            chat_id=chat_id,
            text="📡 Todavía no hay datos guardados. Iniciando búsqueda rápida (2-3 min)..."
        )
        try:
            await scrape_todas_las_fechas(fechas=ENERO_FECHAS[:2])
            report_file = generate_daily_report()
        except Exception as e:
            log.error(f"Error en scraping de /reporte: {e}")

    if report_file:
        try:
            with open(report_file, "rb") as f:
                await context.bot.send_document(
                    chat_id=chat_id,
                    document=f,
                    filename=f"informe_vuelos_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                    caption=(
                        f"📊 *Reporte bajo pedido* — {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
                        f"3 solapas: Vuelos · Paquetes · Histórico\n"
                        f"🔗 Hacé clic en 'Ver oferta' para comprar directo"
                    ),
                    parse_mode="Markdown"
                )
        except Exception as e:
            log.error(f"Error enviando Excel: {e}")
            await context.bot.send_message(
                chat_id=chat_id,
                text=f"❌ Error al enviar el archivo: {e}"
            )
    else:
        await context.bot.send_message(
            chat_id=chat_id,
            text=(
                "⚠️ No se pudieron obtener vuelos ahora.\n"
                "Los sitios pueden estar bloqueando el scraper temporalmente.\n"
                "Probá de nuevo en unos minutos con /reporte o /buscar."
            )
        )

async def cmd_buscar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Búsqueda manual completa con tips."""
    await update.message.reply_text(
        "🔍 Iniciando búsqueda completa en todos los sitios...\n"
        "Esto puede tardar 2-3 minutos, te aviso cuando esté listo."
    )
    try:
        await run_daily_task(app=context.application)
    except Exception as e:
        log.error(f"Error en /buscar: {e}")
        await update.message.reply_text(f"❌ Error durante la búsqueda: {e}")

async def cmd_alerta(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Muestra los mínimos de precio registrados."""
    historial = cargar_historial()
    if not historial:
        await update.message.reply_text(
            "📭 No hay historial de precios aún.\n"
            "Usá /buscar para iniciar el monitoreo."
        )
        return
    lines = ["📊 *Mínimos registrados por fecha:*\n"]
    for clave, precio in sorted(historial.items()):
        fecha = clave.replace("min_pp_", "")
        fecha_fmt = f"{fecha[6:8]}/{fecha[4:6]}/{fecha[:4]}"
        emoji = "🔥" if precio <= PRECIO_ALERTA_PP else "✈️"
        lines.append(
            f"{emoji} {fecha_fmt}: USD {precio:.0f}/pp "
            f"(total x5: USD {precio*5:.0f})"
        )
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────
def main():
    init_db()
    log.info("🤖 Travel Advisor Bot iniciado — corriendo 24/7 en Railway")

    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start",        cmd_start))
    app.add_handler(CommandHandler("reporte",      cmd_reporte))
    app.add_handler(CommandHandler("buscar",       cmd_buscar))
    app.add_handler(CommandHandler("buscar_ahora", cmd_buscar))  # alias viejo
    app.add_handler(CommandHandler("alerta",       cmd_alerta))

    # Scheduler en zona horaria Argentina
    scheduler = AsyncIOScheduler(timezone="America/Argentina/Buenos_Aires")

    # Reporte diario a las 21hs
    scheduler.add_job(
        run_daily_task, "cron", hour=21, minute=0,
        kwargs={"app": app}, id="reporte_21hs"
    )
    # Scraping de monitoreo cada 6hs (3, 9, 15, 21)
    scheduler.add_job(
        run_daily_task, "cron", hour="3,9,15", minute=0,
        kwargs={"app": app}, id="scraping_6hs"
    )
    scheduler.start()

    log.info("Scheduler: scraping a las 3, 9, 15hs | Reporte a las 21hs (ARG)")
    log.info(f"Alerta de precio: USD {PRECIO_ALERTA_PP}/pp")

    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
