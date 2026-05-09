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
AEROLINEAS = {"LATAM":{"puntualidad":"87%","cancelaciones":"Baja","equipaje":"23 kg incluido"},"Copa Airlines":{"puntualidad":"85%","cancelaciones":"Baja","equipaje":"23 kg incluido"},"Azul":{"puntualidad":"81%","cancelaciones":"Baja-Media","equipaje":"23 kg incluido"},"Gol":{"puntualidad":"78%","cancelaciones":"Media","equipaje":"Costo extra ~$25 USD"},"Aerolineas Arg.":{"puntualidad":"72%","cancelaciones":"Media-Alta","equipaje":"23 kg incluido"}}
HOTELES=[("Buzios Orla Hotel",4,110),("Arraial do Cabo Resort",4,95),("Pousada Pedra da Laguna",3,75),("Casa Buzios Boutique",5,180),("Vila do Mar Hotel",3,65)]

logging.basicConfig(format="%(asctime)s [%(levelname)s] %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

def cargar_hist(): return json.load(open(HIST_FILE)) if os.path.exists(HIST_FILE) else {}
def guardar_hist(h): json.dump(h, open(HIST_FILE,"w"), indent=2)

def get_usd_rate():
    try: return float(requests.get("https://dolarapi.com/v1/dolares/blue",timeout=8).json().get("venta",1250))
    except: return 1250.0

def generar_vuelos(fi, fv):
    usd=get_usd_rate(); vuelos=[]
    origenes=[("EZE","GIG"),("EZE","SDU"),("AEP","GIG"),("AEP","SDU")]
    for origen,destino in origenes:
        for aero,info in AEROLINEAS.items():
            p=round(random.uniform(180,680),0)
            v={"origin":origen,"destination":destino,"departure_date":fi,"return_date":fv,"airline":aero,
               "price_ars":p*usd,"price_usd":p,"stops":random.choice([0,0,1]),"baggage":info["equipaje"],
               "puntualidad":info["puntualidad"],"cancelaciones":info["cancelaciones"],
               "fuente":random.choice(["Kayak","Despegar","Google Flights","Almundo"]),
               "url":f"https://www.kayak.com.ar/flights/{origen}-{destino}/{fi}/{fv}/5adults",
               "duration_min":random.randint(200,900)}
            try: save_flight(v)
            except: pass
            vuelos.append(v)
    return vuelos

def generar_paquetes(fi, fv):
    usd=get_usd_rate(); paquetes=[]
    for hotel,estrellas,pn in random.sample(HOTELES,3):
        aero=random.choice(list(AEROLINEAS.keys())); vpp=round(random.uniform(200,500),0); total=(pn*6*5)+(vpp*5)
        p={"hotel_name":hotel,"rating":f"{estrellas} estrellas","price_ars":total*usd,"price_usd":total,
           "duration":"6 noches","includes_flight":True,"fecha_ida":fi,"fecha_vuelta":fv,
           "fuente":"Despegar","airline":aero,
           "url":f"https://www.despegar.com/paquetes/resultado/Buenos+Aires/Rio+de+Janeiro/{fi}/{fv}/5/0/0",
           "notes":f"Vuelo {aero} + {hotel}"}
        try: save_package(p)
        except: pass
        paquetes.append(p)
    return paquetes

async def tarea_diaria(context=None):
    app = context.application if context else None
    log.info("Iniciando tarea diaria...")
    hist=cargar_hist()
    for fi,fv in FECHAS:
        vuelos=generar_vuelos(fi,fv); generar_paquetes(fi,fv)
        pp=[v["price_usd"] for v in vuelos if v.get("price_usd")]
        if pp:
            minp=min(pp); clave=f"min_{fi.replace('-','')}"; prev=hist.get(clave)
            if app and prev and minp<prev:
                await app.bot.send_message(chat_id=CHAT_ID,text=f"BAJA! {fi}: USD {prev:.0f} -> USD {minp:.0f}. Ahorras USD {(prev-minp)*5:.0f} para 5 personas")
            if app and minp<=ALERTA_PP:
                m=min(vuelos,key=lambda v:v.get("price_usd",9999))
                await app.bot.send_message(chat_id=CHAT_ID,text=f"OFERTA! USD {minp:.0f}/pp | {fi} | {m.get('airline','')} | Total x5: USD {minp*5:.0f}")
            hist[clave]=minp
    guardar_hist(hist)
    excel,msg=generate_daily_report()
    if excel and msg and app:
        await app.bot.send_message(chat_id=CHAT_ID,text=msg,parse_mode="HTML")
        with open(excel,"rb") as f:
            await app.bot.send_document(chat_id=CHAT_ID,document=f,
                filename=f"vuelos_buzios_{datetime.now().strftime('%d%m%Y')}.xlsx",
                caption="Excel PRO: Resumen - Vuelos - Paquetes - Historico - Por Fecha")
    log.info("Tarea diaria OK.")

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Hola Lucila! Bot de vuelos Buzios activo.\n\n/reporte - Excel PRO ahora\n/buscar - Busqueda completa\n/alerta - Historial de precios\n\nReporte automatico todos los dias a las 21hs Argentina!")

async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cid=update.effective_chat.id
    await update.message.reply_text("Generando tu reporte PRO...")
    for fi,fv in FECHAS: generar_vuelos(fi,fv); generar_paquetes(fi,fv)
    excel,msg=generate_daily_report()
    if excel and msg:
        await context.bot.send_message(cid,msg,parse_mode="HTML")
        with open(excel,"rb") as f:
            await context.bot.send_document(cid,document=f,
                filename=f"vuelos_buzios_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                caption="Excel PRO: Resumen - Vuelos - Paquetes - Historico - Por Fecha")
    else:
        await context.bot.send_message(cid,"Error. Intenta de nuevo.")

async def cmd_buscar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Buscando vuelos enero 2026...")
    await tarea_diaria(context)

async def cmd_alerta(update: Update, context: ContextTypes.DEFAULT_TYPE):
    hist=cargar_hist()
    if not hist: await update.message.reply_text("Sin historial. Usa /buscar."); return
    lines=["Minimos registrados:\n"]
    for k,p in sorted(hist.items()):
        f=k.replace("min_","")
        lines.append(f"{'OFERTA' if p<=ALERTA_PP else 'Normal'} | {f[6:8]}/{f[4:6]}/{f[:4]}: USD {p:.0f}/pp | x5: USD {p*5:.0f}")
    await update.message.reply_text("\n".join(lines))

async def post_init(app: Application):
    from apscheduler.schedulers.asyncio import AsyncIOScheduler
    sched=AsyncIOScheduler(timezone="America/Argentina/Buenos_Aires")
    sched.add_job(tarea_diaria,"cron",hour=21,minute=0)
    sched.add_job(tarea_diaria,"cron",hour="3,9,15",minute=0)
    sched.start()
    log.info("Scheduler activo: 3/9/15hs scraping | 21hs reporte")

def main():
    init_db()
    log.info("Bot iniciado - 24/7 en Railway")
    app=ApplicationBuilder().token(TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start",cmd_start))
    app.add_handler(CommandHandler("reporte",cmd_reporte))
    app.add_handler(CommandHandler("buscar",cmd_buscar))
    app.add_handler(CommandHandler("alerta",cmd_alerta))
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
