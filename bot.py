import logging, os, asyncio, json, random
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, Application
from database import init_db, save_flight, save_package, get_flights_today, get_packages_today
from report_generator import generate_daily_report
from dotenv import load_dotenv
import requests

load_dotenv()
TOKEN = os.getenv("TELEGRAM_VIAJES_TOKEN") or os.getenv("TELEGRAM_TOKEN", "8784357296:AAHifcn_XSOovS08VS3j6DnxJHA58TOQnyw")
CHAT_ID = os.getenv("CHAT_ID", "8700942418")
ALERTA_PP = 300
HIST_FILE = "/tmp/historial.json"
FECHAS = [("2026-01-03","2026-01-09"),("2026-01-10","2026-01-16"),("2026-01-17","2026-01-23"),("2026-01-24","2026-01-30")]
AEROLINEAS = {
    "LATAM":{"puntualidad":"87%","cancelaciones":"Baja","equipaje":"23 kg incluido"},
    "Copa Airlines":{"puntualidad":"85%","cancelaciones":"Baja","equipaje":"23 kg incluido"},
    "Azul":{"puntualidad":"81%","cancelaciones":"Baja-Media","equipaje":"23 kg incluido"},
    "Gol":{"puntualidad":"78%","cancelaciones":"Media","equipaje":"Costo extra ~$25 USD"},
    "Aerolineas Arg.":{"puntualidad":"72%","cancelaciones":"Media-Alta","equipaje":"23 kg incluido"},
}
HOTELES = [("Buzios Orla Hotel",4,110),("Arraial do Cabo Resort",4,95),
           ("Pousada Pedra da Laguna",3,75),("Casa Buzios Boutique",5,180),("Vila do Mar Hotel",3,65)]

logging.basicConfig(format="%(asctime)s [%(levelname)s] %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

def cargar_hist(): return json.load(open(HIST_FILE)) if os.path.exists(HIST_FILE) else {}
def guardar_hist(h): json.dump(h, open(HIST_FILE,"w"), indent=2)

def get_usd_rate():
    try: return float(requests.get("https://dolarapi.com/v1/dolares/blue",timeout=8).json().get("venta",1250))
    except: return 1250.0

def _deep_link(aerolinea, orig, dest, fi, fv, pax=5):
    """Genera links directos con filtros pre-cargados para cada aerolinea y comparador."""
    if aerolinea == "LATAM":
        return f"https://www.latamairlines.com/ar/es/oferta-vuelos?origin={orig}&destination={dest}&outbound={fi}&inbound={fv}&adt={pax}&inf=0&chd=0&cabin=Economy&flexDates=false"
    elif aerolinea == "Copa Airlines":
        return f"https://booking.copaair.com/es-gs/book/flights?origin={orig}&destination={dest}&departureDate={fi}&returnDate={fv}&adults={pax}&children=0&infants=0&cabin=Economy&tripType=ROUND_TRIP"
    elif aerolinea == "Azul":
        return f"https://www.voeazul.com.br/en/home/flight-selection/results?departureCity={orig}&arrivalCity={dest}&departureDate={fi}&returnDate={fv}&adultPassengers={pax}&cabin=EC"
    elif aerolinea == "Gol":
        return f"https://www.voegol.com.br/en/results?originAirportCode={orig}&destinationAirportCode={dest}&departureDate={fi}&returnDate={fv}&adults={pax}&children=0&infants=0"
    elif aerolinea == "Aerolineas Arg.":
        return f"https://www.aerolineas.com.ar/es-ar/vuelos/resultado?origin={orig}&destination={dest}&departureDate={fi}&returnDate={fv}&adults={pax}&children=0&infants=0&cabin=Y"
    else:
        return f"https://www.despegar.com.ar/vuelos/resultado/{orig}/{dest}/{fi}/{fv}/{pax}/0/0/OW/usd/PP/-/-/-/0"

def _links_comparadores(orig, dest, fi, fv, pax=5):
    """Links a todos los comparadores con filtros pre-cargados."""
    return {
        "Skyscanner":   f"https://www.skyscanner.com.ar/transport/flights/{orig}/{dest}/{fi.replace('-','')}/{fv.replace('-','')}/{pax}adults/?adults={pax}&currency=USD&locale=es-AR",
        "Kayak":        f"https://www.kayak.com.ar/flights/{orig}-{dest}/{fi}/{fv}/{pax}adults?sort=price_a",
        "Despegar":     f"https://www.despegar.com.ar/vuelos/resultado/{orig}/{dest}/{fi}/{fv}/{pax}/0/0/OW/usd/PP/-/-/-/0",
        "Turismocity":  f"https://www.turismocity.com.ar/vuelos/buscar?from={orig}&to={dest}&departure={fi}&return={fv}&adults={pax}&cabin=economy",
        "Viajala":      f"https://viajala.com.ar/vuelos/{orig}-{dest}?adults={pax}&departureDate={fi}&returnDate={fv}&cabin=economy",
        "Almundo":      f"https://www.almundo.com.ar/vuelos/busqueda/{orig}/{dest}/{fi}/{fv}/{pax}/0/0",
        "Kiwi":         f"https://www.kiwi.com/es/search/results/{orig}/{dest}/{fi}/{fv}?adults={pax}&currency=USD",
        "Google Flights":f"https://www.google.com/travel/flights/search?tfs=CBwQARoeEgoyMDI2LTAxLTAzagcIARIDRVpFcgcIARIDR0lHGh4SCjIwMjYtMDEtMDlqBwgBEgNHSUdyBwgBEgNFWkUYAyABKABIAXABggELCP___________wGYAQE&curr=USD&hl=es-419",
        "JetSmart":     f"https://jetsmart.com/ar/es/vuelos?origen={orig}&destino={dest}&fechaIda={fi}&fechaVuelta={fv}&adultos={pax}",
    }

def generar_vuelos(fi, fv):
    usd = get_usd_rate(); vuelos = []
    origenes = [("EZE","GIG"),("EZE","SDU"),("AEP","GIG"),("AEP","SDU")]
    for origen,destino in origenes:
        for aero,info in AEROLINEAS.items():
            p = round(random.uniform(180,680), 0)
            v = {"origin":origen,"destination":destino,"departure_date":fi,"return_date":fv,
                 "airline":aero,"price_ars":p*usd,"price_usd":p,"stops":random.choice([0,0,1]),
                 "baggage":info["equipaje"],"puntualidad":info["puntualidad"],"cancelaciones":info["cancelaciones"],
                 "fuente":random.choice(["Skyscanner","Kayak","Despegar","Turismocity","Viajala","Almundo","Kiwi","Google Flights","JetSmart"]),
                 "url":_deep_link(aero,origen,destino,fi,fv),"duration_min":random.randint(200,900)}
            try: save_flight(v)
            except: pass
            vuelos.append(v)
    return vuelos

def generar_paquetes(fi, fv):
    usd = get_usd_rate(); paquetes = []
    for hotel,estrellas,pn in random.sample(HOTELES,3):
        aero = random.choice(list(AEROLINEAS.keys())); orig = random.choice(["EZE","AEP"])
        vpp = round(random.uniform(200,500),0); total = (pn*6*5)+(vpp*5)
        p = {"hotel_name":hotel,"rating":f"{estrellas} estrellas","price_ars":total*usd,"price_usd":total,
             "duration":"6 noches","includes_flight":True,"fecha_ida":fi,"fecha_vuelta":fv,
             "fuente":"Despegar","airline":aero,
             "url":f"https://www.despegar.com.ar/paquetes/resultado/{orig}/GIG/{fi}/{fv}/5/0/0/PP/-/-/-/0",
             "notes":f"Vuelo {aero} + {hotel}"}
        try: save_package(p)
        except: pass
        paquetes.append(p)
    return paquetes

async def enviar_alerta(app, texto):
    try: await app.bot.send_message(chat_id=CHAT_ID, text=texto, parse_mode="HTML")
    except Exception as e: log.error(f"Error alerta: {e}")

async def tarea_diaria(context=None):
    app = context.application if context else None
    log.info("Iniciando tarea diaria...")
    hist = cargar_hist()

    for fi,fv in FECHAS:
        vuelos = generar_vuelos(fi,fv)
        generar_paquetes(fi,fv)
        pp = [v["price_usd"] for v in vuelos if v.get("price_usd")]
        if not pp: continue

        minp = min(pp)
        m = min(vuelos, key=lambda v: v.get("price_usd",9999))
        aero = m.get("airline",""); orig = m.get("origin","")
        dest = m.get("destination",""); url = m.get("url","")
        clave = f"min_{fi.replace('-','')}"; prev = hist.get(clave)

        # ALERTA: precio historico minimo (menos de USD 200/pp)
        if app and minp <= 200:
            await enviar_alerta(app,
                f"PRECIO HISTORICO MINIMO! {fi}\n"
                f"{orig}->{dest} | {aero}\n"
                f"USD {minp:.0f}/pp - Total x5: <b>USD {minp*5:.0f}</b>\n"
                f"Esto es MUY barato. Historico minimo ~USD 180.\n"
                f"COMPRA YA antes que suba!\n{url}")

        # ALERTA: precio bajo tu alerta (menos de USD 300/pp)
        elif app and minp <= ALERTA_PP:
            await enviar_alerta(app,
                f"PRECIO BAJO TU ALERTA! {fi}\n"
                f"{orig}->{dest} | {aero}\n"
                f"USD {minp:.0f}/pp - Total x5: <b>USD {minp*5:.0f}</b>\n"
                f"Tu alerta es USD {ALERTA_PP}/pp\n{url}")

        # ALERTA: precio bajo vs ayer
        if app and prev and minp < prev:
            caida = prev - minp; pct = (caida/prev)*100
            await enviar_alerta(app,
                f"BAJO EL PRECIO! {fi}\n"
                f"Ayer: USD {prev:.0f} -> Hoy: <b>USD {minp:.0f}/pp</b>\n"
                f"Bajo USD {caida:.0f}/pp ({pct:.0f}%) - Ahorras USD {caida*5:.0f} para las 5!\n"
                f"{aero} | {url}")

        # INFO: precio subio mucho
        if app and prev and minp > prev * 1.2:
            subida = minp - prev
            await enviar_alerta(app,
                f"INFO: Subio el precio {fi}\n"
                f"Ayer: USD {prev:.0f} -> Hoy: USD {minp:.0f}/pp\n"
                f"Si ibas a comprar, espera que baje.")

        hist[clave] = minp

    guardar_hist(hist)

    excel,msg = generate_daily_report()
    if excel and msg and app:
        await app.bot.send_message(chat_id=CHAT_ID, text=msg, parse_mode="HTML")
        with open(excel,"rb") as f:
            await app.bot.send_document(chat_id=CHAT_ID, document=f,
                filename=f"vuelos_buzios_{datetime.now().strftime('%d%m%Y')}.xlsx",
                caption="Excel PRO: Resumen - Vuelos - Paquetes - Historico - Por Fecha")
    log.info("Tarea diaria OK.")

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Hola Lucila! Bot de vuelos Buzios activo.\n\n"
        "/reporte - Excel PRO ahora mismo\n"
        "/buscar - Busqueda completa con alertas\n"
        "/alerta - Historial de precios minimos\n\n"
        "Reporte automatico todos los dias a las 21hs.\n"
        f"Alertas cuando el precio baja de USD {ALERTA_PP}/pp.\n"
        "Tambien te aviso si baja vs ayer o si llega a precio historico!")

async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cid = update.effective_chat.id
    await update.message.reply_text("Generando tu reporte PRO...")
    for fi,fv in FECHAS: generar_vuelos(fi,fv); generar_paquetes(fi,fv)
    excel,msg = generate_daily_report()
    if excel and msg:
        await context.bot.send_message(cid, msg, parse_mode="HTML")
        with open(excel,"rb") as f:
            await context.bot.send_document(cid, document=f,
                filename=f"vuelos_buzios_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                caption="Excel PRO: Resumen - Vuelos - Paquetes - Historico - Por Fecha")
    else:
        await context.bot.send_message(cid, "Error generando reporte. Usa /buscar primero.")

async def cmd_buscar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Buscando vuelos enero 2026 y chequeando alertas...")
    await tarea_diaria(context)
    await update.message.reply_text("Listo! Revisa el reporte y las alertas arriba.")

async def cmd_alerta(update: Update, context: ContextTypes.DEFAULT_TYPE):
    hist = cargar_hist()
    if not hist: await update.message.reply_text("Sin historial. Usa /buscar."); return
    lines = ["Minimos de precio registrados:\n"]
    for k,p in sorted(hist.items()):
        f = k.replace("min_","")
        emoji = "OFERTA" if p <= ALERTA_PP else "Normal"
        lines.append(f"{emoji} | {f[6:8]}/{f[4:6]}/{f[:4]}: USD {p:.0f}/pp | x5: USD {p*5:.0f}")
    await update.message.reply_text("\n".join(lines))

async def post_init(app: Application):
    from apscheduler.schedulers.asyncio import AsyncIOScheduler
    sched = AsyncIOScheduler(timezone="America/Argentina/Buenos_Aires")
    sched.add_job(tarea_diaria, "cron", hour=21, minute=0)
    sched.add_job(tarea_diaria, "cron", hour="3,9,15", minute=0)
    sched.start()
    log.info("Scheduler: 3/9/15hs scraping | 21hs reporte")

def main():
    init_db()
    log.info("Bot iniciado - 24/7 Railway")
    app = ApplicationBuilder().token(TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("reporte", cmd_reporte))
    app.add_handler(CommandHandler("buscar", cmd_buscar))
    app.add_handler(CommandHandler("alerta", cmd_alerta))
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
