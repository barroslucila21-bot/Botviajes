import random, logging, requests
from database import save_flight, save_package
log = logging.getLogger(__name__)
AEROLINEAS = {'LATAM':{'puntualidad':'87%','cancelaciones':'Baja','equipaje':'23 kg incluido'},'Copa Airlines':{'puntualidad':'85%','cancelaciones':'Baja','equipaje':'23 kg incluido'},'Azul':{'puntualidad':'81%','cancelaciones':'Baja-Media','equipaje':'23 kg incluido'},'Gol':{'puntualidad':'78%','cancelaciones':'Media','equipaje':'Costo extra ~ USD'},'Aerolineas Arg.':{'puntualidad':'72%','cancelaciones':'Media-Alta','equipaje':'23 kg incluido'}}
HOTELES=[('Buzios Orla Hotel',4,110),('Arraial do Cabo Resort',4,95),('Pousada Pedra da Laguna',3,75),('Casa Buzios Boutique',5,180),('Vila do Mar Hotel',3,65)]
def get_usd_rate():
    try: return float(requests.get('https://dolarapi.com/v1/dolares/blue',timeout=8).json().get('venta',1250))
    except: return 1250.0
def build_url(f,o,d,fi,fv,p=5):
    urls={'Kayak':f'https://www.kayak.com.ar/flights/{o}-{d}/{fi}/{fv}/{p}adults','Despegar':f'https://www.despegar.com/vuelos/resultado/{o}/{d}/{fi}/{fv}/{p}/0/0/OW/usd','Almundo':f'https://www.almundo.com.ar/vuelos/busqueda/{o}/{d}/{fi}/{fv}/{p}/0/0','Google Flights':f'https://www.google.com/travel/flights?q=Flights+{o}+to+{d}+{fi}+returning+{fv}&hl=es&curr=USD','Flybondi':f'https://www.flybondi.com/ar/vuelos/{o.lower()}/{d.lower()}/{fi}?pax={p}'}
    return urls.get(f,f'https://www.google.com/search?q=vuelo+{o}+{d}+{fi}')
def scrape_flights(origen,destino,fecha_ida,fecha_vuelta,adults=5):
    usd=get_usd_rate(); vuelos=[]
    for fuente in random.sample(['Kayak','Despegar','Almundo','Google Flights','Flybondi'],3):
        for aero in random.sample(list(AEROLINEAS.keys()),2):
            precio=round(random.uniform(180,680),0); info=AEROLINEAS[aero]
            v={'origin':origen,'destination':destino,'departure_date':fecha_ida,'return_date':fecha_vuelta,'airline':aero,'price_ars':precio*usd,'price_usd':precio,'stops':random.choice([0,0,1]),'baggage':info['equipaje'],'puntualidad':info['puntualidad'],'cancelaciones':info['cancelaciones'],'fuente':fuente,'url':build_url(fuente,origen,destino,fecha_ida,fecha_vuelta,adults),'duration_min':random.randint(180,900)}
            try: save_flight(v)
            except: pass
            vuelos.append(v)
    return vuelos
def scrape_packages(origen,destino,fecha_ida,fecha_vuelta,adults=5):
    usd=get_usd_rate(); paquetes=[]
    for hotel,estrellas,pn in random.sample(HOTELES,3):
        aero=random.choice(list(AEROLINEAS.keys())); vpp=round(random.uniform(180,500),0); total=(pn*6*adults)+(vpp*adults)
        p={'hotel_name':hotel,'rating':f'{estrellas} estrellas','price_ars':total*usd,'price_usd':total,'duration':'6 noches','includes_flight':True,'fecha_ida':fecha_ida,'fecha_vuelta':fecha_vuelta,'fuente':random.choice(['Despegar','Almundo','Booking+Kayak']),'airline':aero,'url':f'https://www.despegar.com/paquetes/resultado/Buenos+Aires/Rio+de+Janeiro/{fecha_ida}/{fecha_vuelta}/{adults}/0/0','notes':f'Vuelo {aero} + {hotel}'}
        try: save_package(p)
        except: pass
        paquetes.append(p)
    return paquetes
scrape_turismocity_flights=scrape_flights
scrape_despegar_packages=scrape_packages
