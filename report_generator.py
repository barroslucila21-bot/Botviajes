"""
Genera el Excel PRO con 5 solapas + mensaje Telegram con resumen ejecutivo.
"""

import logging
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from database import get_flights_today, get_packages_today, get_historical

log = logging.getLogger(__name__)

OUTPUT_FILE = "/tmp/informe_viajes_buzios_PRO.xlsx"
PASAJERAS   = 5
NOCHES      = 6
ARS_RATE    = 1250  # estimado dólar blue

AEROLINEAS_INFO = {
    "LATAM":           {"puntualidad":"87%","cancelaciones":"Baja","equipaje":"23 kg incluido","estrellas":"★★★★★","nota":"Mejor balance precio-calidad. La más recomendada.","color":"EBF5FB","fg":"1B4F72"},
    "Copa Airlines":   {"puntualidad":"85%","cancelaciones":"Baja","equipaje":"23 kg incluido","estrellas":"★★★★☆","nota":"Muy confiable. Escala en Panamá (+2hs). Buena opción.","color":"D6EAF8","fg":"2E86C1"},
    "Azul":            {"puntualidad":"81%","cancelaciones":"Baja-Media","equipaje":"23 kg incluido","estrellas":"★★★★☆","nota":"Buena low-cost brasileña. Tiene vuelos directos.","color":"D5F5E3","fg":"1E8449"},
    "Gol":             {"puntualidad":"78%","cancelaciones":"Media","equipaje":"Costo extra ~$25 USD","estrellas":"★★★☆☆","nota":"La más barata pero equipaje se paga aparte.","color":"FFF3CD","fg":"E67E22"},
    "Aerolíneas Arg.": {"puntualidad":"72%","cancelaciones":"Media-Alta","equipaje":"23 kg incluido","estrellas":"★★★☆☆","nota":"Sale de Aeroparque. Historial de demoras.","color":"FDEBD0","fg":"C0392B"},
    "JetSmart":        {"puntualidad":"80%","cancelaciones":"Media","equipaje":"Costo extra","estrellas":"★★★☆☆","nota":"Low-cost. Verificar equipaje antes de comprar.","color":"FFF3CD","fg":"E67E22"},
}

# ── helpers ───────────────────────────────────────────────────────────────────
def fill(c): return PatternFill("solid", fgColor=c)
def brd(color="D5D8DC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def hf(color="FFFFFF", sz=10):
    return Font(name="Calibri", bold=True, color=color, size=sz)
def cf(sz=9, bold=False, color="1A1A2E"):
    return Font(name="Calibri", size=sz, bold=bold, color=color)
def ctr(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def lft(): return Alignment(horizontal="left",   vertical="center", wrap_text=True)

AZUL_OSC="0D1B2A"; AZUL_MED="1B4F72"; AZUL_CLAR="D6EAF8"; AZUL_HI="2E86C1"
VERDE_OSC="1E8449"; VERDE_CL="D5F5E3"; VERDE_MED="27AE60"
NARANJA="E67E22"; NARJ_CL="FDEBD0"; AMARILLO="FFF3CD"
GRIS_OSC="566573"; GRIS_CL="F2F3F4"; BLANCO="FFFFFF"

def dur_fmt(mins):
    if not mins: return "N/D"
    h, m = divmod(int(mins), 60)
    return f"{h}h {m:02d}m"

def badge_cancel(nivel):
    mapa = {"Baja":"🟢 Baja","Baja-Media":"🟡 Baja-Media","Media":"🟡 Media","Media-Alta":"🔴 Media-Alta","Alta":"🔴 Alta"}
    return mapa.get(nivel, nivel)

def badge_escala(esc):
    return "✈️ Directo" if esc == "Directo" else f"🔄 {esc}"

def build_url(fuente, orig, dest, fi, fv, pax=5):
    urls = {
        "Kayak":         f"https://www.kayak.com.ar/flights/{orig}-{dest}/{fi}/{fv}/{pax}adults",
        "Despegar":      f"https://www.despegar.com/vuelos/resultado/{orig}/{dest}/{fi}/{fv}/{pax}/0/0/OW/usd",
        "Almundo":       f"https://www.almundo.com.ar/vuelos/busqueda/{orig}/{dest}/{fi}/{fv}/{pax}/0/0",
        "Google Flights":f"https://www.google.com/travel/flights?q=Flights+{orig}+to+{dest}+{fi}+returning+{fv}&hl=es&curr=USD",
        "Flybondi":      f"https://www.flybondi.com/ar/vuelos/{orig.lower()}/{dest.lower()}/{fi}?pax={pax}",
        "Turismocity":   f"https://www.turismocity.com.ar/vuelos-baratos-desde-{orig}-a-{dest}?adults={pax}&date1={fi}&date2={fv}",
    }
    return urls.get(fuente, f"https://www.google.com/search?q=vuelo+{orig}+{dest}+{fi}")

# ── SOLAPA 1: Resumen ejecutivo ───────────────────────────────────────────────
def sheet_resumen(wb, vuelos, paquetes):
    ws = wb.create_sheet("📋 Resumen")
    ws.sheet_view.showGridLines = False

    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    ws.merge_cells("A1:L1")
    ws["A1"] = "✈️  TRAVEL ADVISOR — BÚZIOS / RÍO DE JANEIRO  |  ENERO 2026"
    ws["A1"].font = Font(name="Calibri", bold=True, size=16, color=BLANCO)
    ws["A1"].fill = fill(AZUL_OSC); ws["A1"].alignment = ctr()
    ws.row_dimensions[1].height = 42

    ws.merge_cells("A2:L2")
    ws["A2"] = f"EZE/AEP → GIG/SDU  |  5 pasajeras  |  6 noches  |  Generado: {now}"
    ws["A2"].font = Font(name="Calibri", size=10, color="BDC3C7", italic=True)
    ws["A2"].fill = fill("162535"); ws["A2"].alignment = ctr()
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 12

    if not vuelos:
        ws.merge_cells("A4:L4")
        ws["A4"] = "⚠️ Sin datos disponibles aún. Usá /buscar para iniciar el monitoreo."
        ws["A4"].font = Font(name="Calibri", size=11, color=NARANJA, bold=True)
        ws["A4"].alignment = ctr()
        return

    precios = [v.get("price_usd", 0) for v in vuelos if v.get("price_usd")]
    if not precios:
        return

    min_pp = min(precios); max_pp = max(precios); avg_pp = round(sum(precios)/len(precios))
    directos = sum(1 for v in vuelos if v.get("stops", 1) == 0)

    kpis = [
        ("💰 Precio Mínimo", f"USD {min_pp:.0f}", "/persona", VERDE_CL, VERDE_OSC),
        ("📊 Precio Promedio", f"USD {avg_pp}", "/persona", AZUL_CLAR, AZUL_HI),
        ("🔝 Precio Máximo", f"USD {max_pp:.0f}", "/persona", NARJ_CL, NARANJA),
        ("✈️ Vuelos Directos", str(directos), f"de {len(vuelos)} totales", "E8DAEF", "7D3C98"),
        ("🏨 Paquetes", str(len(paquetes)), "encontrados", VERDE_CL, VERDE_OSC),
        ("🔔 Alerta activa", "< USD 300", "/persona", AMARILLO, NARANJA),
    ]
    col_starts = [1, 3, 5, 7, 9, 11]
    for i, (titulo, valor, sub, bg, fg) in enumerate(kpis):
        c = col_starts[i]
        for r in [4,5,6,7]:
            ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+1)
            ws.cell(r,c).fill = fill(bg); ws.cell(r,c).border = brd(fg)
        ws.cell(4,c).value = titulo
        ws.cell(4,c).font = Font(name="Calibri", size=9, color=fg, bold=True); ws.cell(4,c).alignment=ctr()
        ws.cell(5,c).value = valor
        ws.cell(5,c).font = Font(name="Calibri", size=18, color=fg, bold=True); ws.cell(5,c).alignment=ctr()
        ws.cell(6,c).value = sub
        ws.cell(6,c).font = Font(name="Calibri", size=8, color=GRIS_OSC, italic=True); ws.cell(6,c).alignment=ctr()
    ws.row_dimensions[4].height=18; ws.row_dimensions[5].height=38
    ws.row_dimensions[6].height=16; ws.row_dimensions[7].height=8

    # Mejor opción
    mejor = min(vuelos, key=lambda v: v.get("price_usd", 9999))
    ws.row_dimensions[8].height = 8
    ws.merge_cells("A9:L9")
    ws["A9"] = "🏆  MEJOR OPCIÓN DEL DÍA"
    ws["A9"].font = Font(name="Calibri", bold=True, size=11, color=BLANCO)
    ws["A9"].fill = fill(VERDE_OSC); ws["A9"].alignment = Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[9].height = 26

    aero = mejor.get("airline","N/D"); esc = "Directo" if mejor.get("stops",1)==0 else f"Escala"
    ws.merge_cells("A10:L10")
    ws["A10"] = (f"  ✈️  {aero}  |  {mejor.get('origin','')}→{mejor.get('destination','')}  |  "
                 f"{mejor.get('departure_date','')}  |  {esc}  |  "
                 f"USD {mejor.get('price_usd',0):.0f}/pp  →  Total x5: USD {mejor.get('price_usd',0)*5:.0f}  |  "
                 f"{mejor.get('baggage','N/D')}  |  Vía {mejor.get('fuente','N/D')}")
    ws["A10"].font = Font(name="Calibri", size=10, color=VERDE_OSC, bold=True)
    ws["A10"].fill = fill(VERDE_CL); ws["A10"].alignment = lft()
    ws.row_dimensions[10].height = 22

    # Top 5
    ws.row_dimensions[11].height = 8
    ws.merge_cells("A12:L12")
    ws["A12"] = "⚡  TOP 5 VUELOS MÁS BARATOS"
    ws["A12"].font = hf(BLANCO, 11); ws["A12"].fill = fill(AZUL_HI)
    ws["A12"].alignment = Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[12].height = 26

    top5_hdrs = ["#","Aerolínea","Origen→Dest","Fecha","Escala","Duración","Precio/pp (USD)","Total x5 (USD)","Equipaje","Puntual","Fuente","🔗 Ver"]
    for ci,h in enumerate(top5_hdrs,1):
        c=ws.cell(13,ci,h); c.font=hf(BLANCO,9); c.fill=fill(AZUL_MED); c.alignment=ctr(); c.border=brd()
    ws.row_dimensions[13].height=26

    top5 = sorted(vuelos, key=lambda v: v.get("price_usd",9999))[:5]
    for ri, v in enumerate(top5, 14):
        bg = VERDE_CL if ri==14 else (GRIS_CL if ri%2==0 else BLANCO)
        info = AEROLINEAS_INFO.get(v.get("airline",""), {})
        vals = [ri-13, v.get("airline",""), f"{v.get('origin','')}→{v.get('destination','')}",
                v.get("departure_date",""), badge_escala("Directo" if v.get("stops",1)==0 else "Escala"),
                dur_fmt(v.get("duration_min",0)), v.get("price_usd",0), v.get("price_usd",0)*5,
                v.get("baggage","N/D"), info.get("puntualidad","N/D"), v.get("fuente","N/D"), "🔗 Ver"]
        for ci,val in enumerate(vals,1):
            cell=ws.cell(ri,ci,val); cell.fill=fill(bg); cell.border=brd(); cell.alignment=ctr()
            if ci in[7,8]:
                cell.number_format='#,##0 "USD"'
                cell.font=Font(name="Calibri",size=9,bold=(ri==14),color=VERDE_OSC if ri==14 else "1A1A2E")
            else:
                cell.font=cf(9,bold=(ri==14 and ci in[2,7,8]))
        url=v.get("url","https://www.google.com")
        lc=ws.cell(ri,12,"🔗 Ver"); lc.hyperlink=url
        lc.font=Font(name="Calibri",size=9,color="1155CC",underline="single")
        lc.fill=fill(bg); lc.border=brd(); lc.alignment=ctr()
        ws.row_dimensions[ri].height=16

    # Análisis aerolíneas
    ws.row_dimensions[19].height=8
    ws.merge_cells("A20:L20")
    ws["A20"]="📊  ANÁLISIS POR AEROLÍNEA"
    ws["A20"].font=hf(BLANCO,11); ws["A20"].fill=fill(AZUL_MED)
    ws["A20"].alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[20].height=26

    ahdrs=["Aerolínea","Rating","Puntualidad","Cancelaciones","Equipaje","Precio Mín (USD)","Precio Máx (USD)","Precio Prom (USD)","# Vuelos","Recomendación"]
    for ci,h in enumerate(ahdrs,1):
        c=ws.cell(21,ci,h); c.font=hf(BLANCO,9); c.fill=fill(GRIS_OSC); c.alignment=ctr(); c.border=brd()
    ws.row_dimensions[21].height=26

    ri=22
    for aero_name, info in AEROLINEAS_INFO.items():
        vv=[v for v in vuelos if v.get("airline","")==aero_name]
        if not vv: continue
        pp=[v.get("price_usd",0) for v in vv if v.get("price_usd")]
        if not pp: continue
        bg=info.get("color","FFFFFF")
        vals=[aero_name, info["estrellas"], info["puntualidad"], badge_cancel(info["cancelaciones"]),
              info["equipaje"], min(pp), max(pp), round(sum(pp)/len(pp)), len(vv), info["nota"]]
        for ci,val in enumerate(vals,1):
            cell=ws.cell(ri,ci,val); cell.fill=fill(bg); cell.border=brd()
            cell.alignment=lft() if ci in[1,5,10] else ctr()
            if ci in[6,7,8]:
                cell.number_format='#,##0 "USD"'
                cell.font=Font(name="Calibri",size=9,bold=(ci==6),color=VERDE_OSC if ci==6 else "1A1A2E")
            elif ci==2:
                cell.font=Font(name="Calibri",size=9,color=NARANJA)
            elif ci==4:
                cm={"🟢 Baja":VERDE_OSC,"🟡 Baja-Media":NARANJA,"🟡 Media":NARANJA,"🔴 Media-Alta":"C0392B"}
                cell.font=Font(name="Calibri",size=9,bold=True,color=cm.get(val,"1A1A2E"))
            else:
                cell.font=cf(9)
        ws.row_dimensions[ri].height=18; ri+=1

    widths=[4,13,10,10,10,10,13,12,10,10,10,18]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A13"

# ── SOLAPA 2: Vuelos completo ─────────────────────────────────────────────────
def sheet_vuelos(wb, vuelos):
    wv = wb.create_sheet("✈️ Vuelos")
    wv.sheet_view.showGridLines = False
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    wv.merge_cells("A1:S1")
    wv["A1"] = f"✈️  INFORME DIARIO DE VUELOS — Búzios / Río de Janeiro  |  {datetime.now().strftime('%d/%m/%Y')}"
    wv["A1"].font = Font(name="Calibri",bold=True,size=14,color=BLANCO)
    wv["A1"].fill = fill(AZUL_OSC); wv["A1"].alignment = ctr(); wv.row_dimensions[1].height=38

    wv.merge_cells("A2:S2")
    wv["A2"] = f"EZE/AEP → GIG/SDU  |  5 pasajeras  |  6 noches  |  Enero 2026  |  Generado: {now}"
    wv["A2"].font = Font(name="Calibri",size=9,color="BDC3C7",italic=True)
    wv["A2"].fill = fill("162535"); wv["A2"].alignment = ctr(); wv.row_dimensions[2].height=20
    wv.row_dimensions[3].height=6

    hdrs=[
        ("#",3),("Aerolínea",14),("Rating",8),("Origen",7),("Destino",7),("Destino Completo",24),
        ("Fecha Ida",11),("Fecha Vuelta",12),("Noches",7),("Escala",14),("Duración",10),
        ("Precio/pp\n(USD)",12),("Total x5\n(USD)",13),("Precio ARS\nest.",13),
        ("Equipaje\nbodega",21),("Puntualidad",11),("Cancelaciones",14),
        ("Fuente",13),("🔗 Comprar",22),
    ]
    for ci,(lbl,w) in enumerate(hdrs,1):
        c=wv.cell(4,ci,lbl); c.font=hf(BLANCO,9); c.fill=fill(AZUL_MED)
        c.alignment=ctr(); c.border=brd()
        wv.column_dimensions[get_column_letter(ci)].width=w
    wv.row_dimensions[4].height=32

    if not vuelos:
        wv.cell(5,1,"Sin datos — usá /buscar para iniciar").font=cf(10,color=NARANJA)
        wv.freeze_panes="A5"; return

    vuelos_ord = sorted(vuelos, key=lambda v: v.get("price_usd",9999))
    minp = min(v.get("price_usd",9999) for v in vuelos_ord)
    destinos_full = {"GIG":"Río de Janeiro (Galeão)","SDU":"Río de Janeiro (Santos Dumont)"}

    for ri, v in enumerate(vuelos_ord, 5):
        ppp = v.get("price_usd",0) or 0
        esc_raw = "Directo" if v.get("stops",1)==0 else v.get("stops","Escala")
        es_min = ppp == minp
        bg = VERDE_CL if es_min else (GRIS_CL if ri%2==0 else BLANCO)
        info = AEROLINEAS_INFO.get(v.get("airline",""), {})
        dest = v.get("destination","")
        vals = [ri-4, v.get("airline",""), info.get("estrellas",""),
                v.get("origin",""), dest, destinos_full.get(dest, dest),
                v.get("departure_date",""), v.get("return_date",""), NOCHES,
                badge_escala(str(esc_raw)), dur_fmt(v.get("duration_min",0)),
                ppp, ppp*PASAJERAS, ppp*ARS_RATE,
                v.get("baggage","N/D"), info.get("puntualidad","N/D"),
                badge_cancel(info.get("cancelaciones","N/D")),
                v.get("fuente","N/D"), None]
        for ci,val in enumerate(vals,1):
            cell=wv.cell(ri,ci,val); cell.fill=fill(bg); cell.border=brd()
            cell.alignment=lft() if ci in[2,6,15] else ctr()
            if ci in[12,13]:
                cell.number_format='#,##0 "USD"'
                cell.font=Font(name="Calibri",size=9,bold=es_min,color=VERDE_OSC if es_min else "1A1A2E")
            elif ci==14:
                cell.number_format='#,##0 "ARS"'; cell.font=cf(8,color=GRIS_OSC)
            elif ci==3:
                cell.font=Font(name="Calibri",size=9,color=NARANJA)
            elif ci==17:
                cm={"🟢 Baja":VERDE_OSC,"🟡 Baja-Media":NARANJA,"🟡 Media":NARANJA,"🔴 Media-Alta":"C0392B"}
                cell.font=Font(name="Calibri",size=9,bold=True,color=cm.get(str(val),"1A1A2E"))
            else:
                cell.font=cf(9,bold=es_min and ci in[2,12,13])
        url=v.get("url","https://www.google.com")
        lc=wv.cell(ri,19,"🔗 Ver oferta"); lc.hyperlink=url
        lc.font=Font(name="Calibri",size=9,color="1155CC",underline="single")
        lc.fill=fill(bg); lc.border=brd(); lc.alignment=ctr()
        wv.row_dimensions[ri].height=15

    lr=5+len(vuelos_ord)
    wv.merge_cells(f"A{lr+1}:F{lr+1}")
    wv[f"A{lr+1}"]="  🟢 Verde = precio más bajo del día  |  ★ Rating de confiabilidad  |  🔗 Clic para abrir con tus datos pre-cargados"
    wv[f"A{lr+1}"].font=Font(name="Calibri",size=8,italic=True,color=GRIS_OSC)
    wv[f"A{lr+1}"].fill=fill(GRIS_CL); wv[f"A{lr+1}"].alignment=lft()
    wv.freeze_panes="A5"

# ── SOLAPA 3: Paquetes ────────────────────────────────────────────────────────
def sheet_paquetes(wb, paquetes):
    wp = wb.create_sheet("🏨 Paquetes")
    wp.sheet_view.showGridLines = False
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    wp.merge_cells("A1:O1")
    wp["A1"] = f"🏨  PAQUETES COMPLETOS — Vuelo + Hotel  |  {datetime.now().strftime('%d/%m/%Y')}"
    wp["A1"].font = Font(name="Calibri",bold=True,size=14,color=BLANCO)
    wp["A1"].fill = fill(VERDE_OSC); wp["A1"].alignment = ctr(); wp.row_dimensions[1].height=38

    wp.merge_cells("A2:O2")
    wp["A2"] = "5 pasajeras  |  6 noches  |  Enero 2026  |  🔗 Clic para abrir con datos pre-cargados"
    wp["A2"].font = Font(name="Calibri",size=9,color="D5F5E3",italic=True)
    wp["A2"].fill = fill("1A5E34"); wp["A2"].alignment = ctr(); wp.row_dimensions[2].height=20
    wp.row_dimensions[3].height=6

    phdrs=[
        ("#",3),("Hotel",26),("⭐",8),("Fecha Ida",11),("Fecha Vuelta",12),
        ("Aerolínea",13),("Vuelo/pp\n(USD)",12),("Hotel/noche\n(USD)",13),
        ("Total vuelos\n(USD)",14),("Total hotel\n(USD)",14),
        ("TOTAL PAQUETE\n(USD)",16),("TOTAL/pp\n(USD)",14),
        ("Equipaje",20),("Fuente",13),("🔗 Ver paquete",22),
    ]
    for ci,(lbl,w) in enumerate(phdrs,1):
        c=wp.cell(4,ci,lbl); c.font=hf(BLANCO,9); c.fill=fill(VERDE_MED)
        c.alignment=ctr(); c.border=brd()
        wp.column_dimensions[get_column_letter(ci)].width=w
    wp.row_dimensions[4].height=32

    if not paquetes:
        wp.cell(5,1,"Sin datos — el scraper buscará paquetes automáticamente").font=cf(10,color=NARANJA)
        wp.freeze_panes="A5"; return

    paq_ord = sorted(paquetes, key=lambda p: p.get("price_usd",9999))
    minpaq = min(p.get("price_usd",9999) for p in paq_ord)
    for ri,p in enumerate(paq_ord,5):
        pusd = p.get("price_usd",0) or 0
        es_min = pusd == minpaq
        bg = VERDE_CL if es_min else (GRIS_CL if ri%2==0 else BLANCO)
        info = AEROLINEAS_INFO.get(p.get("airline",""), {})
        vals=[ri-4, p.get("hotel_name","N/D"), p.get("rating","N/D"),
              p.get("fecha_ida",""), p.get("fecha_vuelta",""),
              p.get("airline","N/D"),
              p.get("price_usd",0)//6//5 if p.get("price_usd") else 0,
              p.get("price_ars",0)//NOCHES//PASAJERAS if p.get("price_ars") else 0,
              p.get("price_usd",0)//6*1 if p.get("price_usd") else 0,
              p.get("price_ars",0)//1 if p.get("price_ars") else 0,
              pusd, round(pusd/PASAJERAS,0) if pusd else 0,
              info.get("equipaje","N/D"), p.get("fuente","N/D"), None]
        for ci,val in enumerate(vals,1):
            cell=wp.cell(ri,ci,val); cell.fill=fill(bg); cell.border=brd()
            cell.alignment=lft() if ci in[2] else ctr()
            if ci in[7,8,9,10,11,12]:
                cell.number_format='#,##0 "USD"'
                cell.font=Font(name="Calibri",size=9,bold=es_min and ci in[11,12],
                              color=VERDE_OSC if (es_min and ci in[11,12]) else "1A1A2E")
            else:
                cell.font=cf(9,bold=es_min and ci==2)
        url=p.get("url","https://www.despegar.com")
        lc=wp.cell(ri,15,"🔗 Ver paquete"); lc.hyperlink=url
        lc.font=Font(name="Calibri",size=9,color="1155CC",underline="single")
        lc.fill=fill(bg); lc.border=brd(); lc.alignment=ctr()
        wp.row_dimensions[ri].height=15
    wp.freeze_panes="A5"

# ── SOLAPA 4: Histórico + Asesoría ───────────────────────────────────────────
def sheet_historico(wb, vuelos):
    wh = wb.create_sheet("📊 Histórico & Asesoría")
    wh.sheet_view.showGridLines = False

    wh.merge_cells("A1:J1")
    wh["A1"] = "📊  HISTÓRICO DE PRECIOS + ASESORÍA  |  EZE/AEP → GIG/SDU"
    wh["A1"].font = Font(name="Calibri",bold=True,size=14,color=BLANCO)
    wh["A1"].fill = fill(NARANJA); wh["A1"].alignment = ctr(); wh.row_dimensions[1].height=38

    wh.merge_cells("A2:J2")
    wh["A2"] = "Referencia histórica 2023-2024 + comparación con precios actuales"
    wh["A2"].font = Font(name="Calibri",size=9,color=BLANCO,italic=True)
    wh["A2"].fill = fill("C0392B"); wh["A2"].alignment = ctr(); wh.row_dimensions[2].height=20
    wh.row_dimensions[3].height=8

    hhdrs=[("Año",7),("Aerolínea",16),("Ruta",20),("Mín hist.\n(USD/pp)",13),
           ("Máx hist.\n(USD/pp)",13),("Prom hist.\n(USD/pp)",14),
           ("Precio hoy\nmín (USD/pp)",14),("vs Promedio\nhistórico",13),
           ("Mejor mes",12),("Veredicto hoy",16)]
    for ci,(lbl,w) in enumerate(hhdrs,1):
        c=wh.cell(4,ci,lbl); c.font=hf(BLANCO,9); c.fill=fill(NARANJA)
        c.alignment=ctr(); c.border=brd()
        wh.column_dimensions[get_column_letter(ci)].width=w
    wh.row_dimensions[4].height=30

    hist_data={
        "LATAM":          {2023:(200,580,370),2024:(215,610,390)},
        "Copa Airlines":  {2023:(180,520,320),2024:(190,550,340)},
        "Azul":           {2023:(170,490,300),2024:(185,510,320)},
        "Gol":            {2023:(155,470,280),2024:(170,500,300)},
        "Aerolíneas Arg.":{2023:(210,600,390),2024:(225,630,410)},
    }
    meses_b={"LATAM":"Sep","Copa Airlines":"May","Azul":"Jun","Gol":"May","Aerolíneas Arg.":"Oct"}

    ri=5
    for aero,years in hist_data.items():
        vv=[v for v in vuelos if v.get("airline","")==aero] if vuelos else []
        min_hoy=min((v.get("price_usd",9999) for v in vv),default=None)
        for year,(pmin,pmax,pprom) in years.items():
            rf=GRIS_CL if ri%2==0 else BLANCO
            if min_hoy:
                diff=min_hoy-pprom
                if diff>80: veredicto="🔴 Caro vs hist."
                elif diff>0: veredicto="🟡 Normal"
                elif diff>-50: veredicto="🟢 Buena oferta"
                else: veredicto="🟢🟢 EXCELENTE oferta"
                hoy_str=f"USD {min_hoy:.0f}"
                diff_str=f"{'+' if diff>=0 else ''}{diff:.0f} USD"
            else:
                veredicto="—"; hoy_str="Sin datos"; diff_str="—"
            vals=[year,aero,"EZE/AEP → GIG/SDU",pmin,pmax,pprom,hoy_str,diff_str,meses_b.get(aero,"Jun"),veredicto]
            for ci,val in enumerate(vals,1):
                cell=wh.cell(ri,ci,val); cell.fill=fill(rf); cell.border=brd()
                cell.alignment=ctr() if ci!=3 else lft()
                if ci in[4,5,6]:
                    cell.number_format='#,##0 "USD"'; cell.font=cf(9)
                elif ci==10:
                    cm={"🟢🟢 EXCELENTE oferta":VERDE_OSC,"🟢 Buena oferta":VERDE_OSC,"🟡 Normal":NARANJA,"🔴 Caro vs hist.":"C0392B"}
                    cell.font=Font(name="Calibri",size=9,bold=True,color=cm.get(val,GRIS_OSC))
                else:
                    cell.font=cf(9)
            wh.row_dimensions[ri].height=15; ri+=1

    # Asesoría
    ri+=2
    wh.merge_cells(f"A{ri}:J{ri}")
    wh[f"A{ri}"]="💡  ASESORÍA — ¿QUÉ AEROLÍNEA TE CONVIENE PARA BÚZIOS?"
    wh[f"A{ri}"].font=Font(name="Calibri",bold=True,size=11,color=BLANCO)
    wh[f"A{ri}"].fill=fill(AZUL_OSC); wh[f"A{ri}"].alignment=Alignment(horizontal="left",vertical="center",indent=1)
    wh.row_dimensions[ri].height=26; ri+=1

    asesorias=[
        ("LATAM","★★★★★",VERDE_CL,VERDE_OSC,"La más recomendada. Mejor balance precio-puntualidad-servicio. Equipaje 23kg incluido. Si la encontrás a menos de USD 300/pp es oferta excelente."),
        ("Copa Airlines","★★★★☆","D6EAF8",AZUL_HI,"Segunda mejor opción. Muy confiable. Escala en Panamá (+2hs de viaje). Equipaje incluido. Buen servicio a bordo."),
        ("Azul","★★★★☆",VERDE_CL,VERDE_OSC,"Excelente low-cost brasileña. Tiene vuelos directos. Equipaje incluido. Buena alternativa si LATAM está caro o sin disponibilidad."),
        ("Gol","★★★☆☆","FFF3CD",NARANJA,"La más barata PERO el equipaje en bodega se paga APARTE (~USD 25 por bolso). Solo conviene si van con carry-on solamente."),
        ("Aerolíneas Arg.","★★★☆☆","FDEBD0","C0392B","Sale de Aeroparque (más práctico desde Palermo/Recoleta). Equipaje incluido. PERO peor puntualidad del grupo. Solo si el precio lo justifica."),
    ]
    for aero,stars,bg,fg,consejo in asesorias:
        vv=[v for v in vuelos if v.get("airline","")==aero] if vuelos else []
        min_pp_a=min((v.get("price_usd",9999) for v in vv),default=None)
        precio_txt=f"Desde USD {min_pp_a:.0f}/pp" if min_pp_a else "Sin datos hoy"
        wh.cell(ri,1,aero).font=Font(name="Calibri",size=10,bold=True,color=fg)
        wh.cell(ri,1).fill=fill(bg); wh.cell(ri,1).border=brd(); wh.cell(ri,1).alignment=ctr()
        wh.cell(ri,2,stars).font=Font(name="Calibri",size=10,color=NARANJA)
        wh.cell(ri,2).fill=fill(bg); wh.cell(ri,2).border=brd(); wh.cell(ri,2).alignment=ctr()
        wh.cell(ri,3,precio_txt).font=Font(name="Calibri",size=9,bold=True,color=VERDE_OSC)
        wh.cell(ri,3).fill=fill(bg); wh.cell(ri,3).border=brd(); wh.cell(ri,3).alignment=ctr()
        wh.merge_cells(start_row=ri,start_column=4,end_row=ri,end_column=10)
        wh.cell(ri,4,consejo).font=cf(9); wh.cell(ri,4).fill=fill(bg)
        wh.cell(ri,4).border=brd()
        wh.cell(ri,4).alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
        wh.row_dimensions[ri].height=26; ri+=1
    wh.freeze_panes="A5"

# ── SOLAPA 5: Análisis por fecha ──────────────────────────────────────────────
def sheet_por_fecha(wb, vuelos):
    wf = wb.create_sheet("📅 Por Fecha")
    wf.sheet_view.showGridLines = False

    wf.merge_cells("A1:H1")
    wf["A1"] = "📅  ANÁLISIS POR SEMANA — ¿Cuándo conviene más viajar?"
    wf["A1"].font = Font(name="Calibri",bold=True,size=14,color=BLANCO)
    wf["A1"].fill = fill("7D3C98"); wf["A1"].alignment = ctr(); wf.row_dimensions[1].height=38

    wf.merge_cells("A2:H2")
    wf["A2"] = "Comparación de precios por semana de enero 2026"
    wf["A2"].font = Font(name="Calibri",size=9,color=BLANCO,italic=True)
    wf["A2"].fill = fill("6C3483"); wf["A2"].alignment = ctr(); wf.row_dimensions[2].height=20
    wf.row_dimensions[3].height=8

    fhdrs=[("Semana (Ida)",14),("Vuelta",12),("# Vuelos",9),("Precio Mín\n(USD/pp)",14),
           ("Precio Prom\n(USD/pp)",14),("Precio Máx\n(USD/pp)",14),
           ("Aerolínea más\nbarata",16),("Veredicto",16)]
    for ci,(lbl,w) in enumerate(fhdrs,1):
        c=wf.cell(4,ci,lbl); c.font=hf(BLANCO,9); c.fill=fill("7D3C98")
        c.alignment=ctr(); c.border=brd()
        wf.column_dimensions[get_column_letter(ci)].width=w
    wf.row_dimensions[4].height=30

    semanas=[("2026-01-03","2026-01-09"),("2026-01-10","2026-01-16"),
             ("2026-01-17","2026-01-23"),("2026-01-24","2026-01-30")]

    mejor_semana_precio = 9999
    if vuelos:
        for fi,fv in semanas:
            vv=[v for v in vuelos if v.get("departure_date","")==fi]
            if vv:
                mp=min(v.get("price_usd",9999) for v in vv)
                if mp < mejor_semana_precio:
                    mejor_semana_precio = mp

    for ri,(fi,fv) in enumerate(semanas,5):
        vv=[v for v in vuelos if v.get("departure_date","")==fi] if vuelos else []
        fi_fmt=f"{fi[8:]}/{fi[5:7]}/{fi[:4]}"
        fv_fmt=f"{fv[8:]}/{fv[5:7]}/{fv[:4]}"
        if vv:
            pp=[v.get("price_usd",0) for v in vv if v.get("price_usd")]
            mejor_aero=min(vv,key=lambda x:x.get("price_usd",9999)).get("airline","N/D")
            min_p=min(pp); max_p=max(pp); avg_p=round(sum(pp)/len(pp))
        else:
            pp=[]; mejor_aero="Sin datos"; min_p=max_p=avg_p=0
        if min_p==0: veredicto="—"
        elif min_p==mejor_semana_precio: veredicto="🟢🟢 LA MÁS BARATA"
        elif min_p<300: veredicto="🟢 Muy barata"
        elif min_p<450: veredicto="🟡 Normal"
        else: veredicto="🔴 Cara"
        bg="E8DAEF" if (min_p==mejor_semana_precio and pp) else (GRIS_CL if ri%2==0 else BLANCO)
        vals=[fi_fmt,fv_fmt,len(vv),min_p or "—",avg_p or "—",max_p or "—",mejor_aero,veredicto]
        for ci,val in enumerate(vals,1):
            cell=wf.cell(ri,ci,val); cell.fill=fill(bg); cell.border=brd(); cell.alignment=ctr()
            if ci in[4,5,6] and isinstance(val,int):
                cell.number_format='#,##0 "USD"'
                cell.font=Font(name="Calibri",size=10,bold=(ci==4 and min_p==mejor_semana_precio),
                              color=VERDE_OSC if (ci==4 and min_p==mejor_semana_precio) else "1A1A2E")
            elif ci==8:
                cm={"🟢🟢 LA MÁS BARATA":VERDE_OSC,"🟢 Muy barata":VERDE_OSC,"🟡 Normal":NARANJA,"🔴 Cara":"C0392B"}
                cell.font=Font(name="Calibri",size=10,bold=True,color=cm.get(val,GRIS_OSC))
            else:
                cell.font=cf(10,bold=(min_p==mejor_semana_precio and pp))
        wf.row_dimensions[ri].height=20
    wf.freeze_panes="A5"

# ── FUNCIÓN PRINCIPAL ─────────────────────────────────────────────────────────
def generate_daily_report(filename=OUTPUT_FILE):
    vuelos   = get_flights_today()
    paquetes = get_packages_today()

    if not vuelos and not paquetes:
        log.warning("Sin datos en DB para generar reporte.")
        return None, None

    wb = Workbook()
    wb.remove(wb.active)

    sheet_resumen(wb, vuelos, paquetes)
    sheet_vuelos(wb, vuelos)
    sheet_paquetes(wb, paquetes)
    sheet_historico(wb, vuelos)
    sheet_por_fecha(wb, vuelos)

    wb.save(filename)
    log.info(f"Excel PRO generado: {filename}")

    # Generar mensaje Telegram
    mensaje = generar_mensaje_telegram(vuelos, paquetes)
    return filename, mensaje

def generar_mensaje_telegram(vuelos, paquetes):
    """Genera el resumen ejecutivo para Telegram."""
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    if not vuelos:
        return (
            f"📊 <b>Reporte diario — {now}</b>\n"
            f"⚠️ Sin vuelos encontrados hoy. Reintentando en el próximo ciclo.\n"
            f"Usá /buscar para una búsqueda manual."
        )

    precios = [v.get("price_usd",0) for v in vuelos if v.get("price_usd")]
    min_pp = min(precios); max_pp = max(precios); avg_pp = round(sum(precios)/len(precios))
    mejor = min(vuelos, key=lambda v: v.get("price_usd",9999))
    directos = sum(1 for v in vuelos if v.get("stops",1)==0)

    # Top 3
    top3 = sorted(vuelos, key=lambda v: v.get("price_usd",9999))[:3]
    top3_lines = ""
    for i,v in enumerate(top3,1):
        esc = "Directo ✈️" if v.get("stops",1)==0 else f"Escala 🔄"
        top3_lines += (f"  <b>{i}.</b> {v.get('airline','')} | {v.get('departure_date','')} | "
                      f"{esc} | <b>USD {v.get('price_usd',0):.0f}/pp</b> (x5: USD {v.get('price_usd',0)*5:.0f})\n")

    # Semana más barata
    semanas={"2026-01-03":"3-9 ene","2026-01-10":"10-16 ene","2026-01-17":"17-23 ene","2026-01-24":"24-30 ene"}
    mejor_semana=""
    mejor_precio_sem=9999
    for fi,label in semanas.items():
        vv=[v for v in vuelos if v.get("departure_date","")==fi]
        if vv:
            mp=min(v.get("price_usd",9999) for v in vv)
            if mp<mejor_precio_sem:
                mejor_precio_sem=mp; mejor_semana=label

    # Veredicto vs histórico
    if min_pp < 200: veredicto="🟢🟢 PRECIO HISTÓRICO MUY BAJO — ¡COMPRÁ YA!"
    elif min_pp < 280: veredicto="🟢 Precio muy bueno — debería comprarse"
    elif min_pp < 380: veredicto="🟡 Precio normal para la ruta"
    else: veredicto="🔴 Precio alto — esperá una baja"

    msg = (
        f"✈️ <b>REPORTE DIARIO DE VUELOS</b>\n"
        f"📅 {now} | 🇧🇷 Búzios/Río | 5 pasajeras | 6 noches\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"💰 <b>PRECIOS HOY</b>\n"
        f"  🟢 Mínimo: <b>USD {min_pp:.0f}/pp</b> → Total x5: <b>USD {min_pp*5:.0f}</b>\n"
        f"  📊 Promedio: USD {avg_pp}/pp\n"
        f"  🔴 Máximo: USD {max_pp:.0f}/pp\n\n"
        f"🏆 <b>MEJOR OPCIÓN</b>\n"
        f"  {mejor.get('airline','')} | {mejor.get('departure_date','')} | "
        f"{'Directo ✈️' if mejor.get('stops',1)==0 else 'Escala 🔄'}\n"
        f"  <b>USD {mejor.get('price_usd',0):.0f}/pp — Total x5: USD {mejor.get('price_usd',0)*5:.0f}</b>\n"
        f"  Equipaje: {mejor.get('baggage','N/D')} | Vía {mejor.get('fuente','N/D')}\n\n"
        f"⚡ <b>TOP 3 MÁS BARATOS</b>\n{top3_lines}\n"
        f"📊 <b>ESTADÍSTICAS</b>\n"
        f"  📋 {len(vuelos)} vuelos analizados | ✈️ {directos} directos\n"
        f"  📅 Semana más barata: <b>{mejor_semana}</b> (desde USD {mejor_precio_sem:.0f}/pp)\n\n"
        f"🎯 <b>VEREDICTO</b>\n  {veredicto}\n\n"
        f"📎 Excel adjunto con detalles completos, histórico y links para comprar."
    )
    return msg
