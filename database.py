import sqlite3, logging
DB_PATH = "travel_advisor.db"
log = logging.getLogger(__name__)

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute('''CREATE TABLE IF NOT EXISTS flights (id INTEGER PRIMARY KEY AUTOINCREMENT, search_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP, origin TEXT, destination TEXT, departure_date TEXT, return_date TEXT, airline TEXT, price_ars REAL, price_usd REAL, stops INTEGER, baggage TEXT, puntualidad TEXT, cancelaciones TEXT, fuente TEXT, url TEXT, duration_min INTEGER)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS packages (id INTEGER PRIMARY KEY AUTOINCREMENT, search_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP, hotel_name TEXT, rating TEXT, price_ars REAL, price_usd REAL, duration TEXT, includes_flight BOOLEAN, fecha_ida TEXT, fecha_vuelta TEXT, fuente TEXT, url TEXT, notes TEXT, airline TEXT)''')
    conn.commit(); conn.close()
    log.info("DB OK")

def save_flight(f):
    conn = sqlite3.connect(DB_PATH)
    conn.execute('INSERT INTO flights (origin,destination,departure_date,return_date,airline,price_ars,price_usd,stops,baggage,puntualidad,cancelaciones,fuente,url,duration_min) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
        (f.get('origin'), f.get('destination'), f.get('departure_date'), f.get('return_date'), f.get('airline'), f.get('price_ars',0), f.get('price_usd',0), f.get('stops',1), f.get('baggage',''), f.get('puntualidad',''), f.get('cancelaciones',''), f.get('fuente',''), f.get('url',''), f.get('duration_min',0)))
    conn.commit(); conn.close()

def save_package(p):
    conn = sqlite3.connect(DB_PATH)
    conn.execute('INSERT INTO packages (hotel_name,rating,price_ars,price_usd,duration,includes_flight,fecha_ida,fecha_vuelta,fuente,url,notes,airline) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
        (p.get('hotel_name',''), p.get('rating',''), p.get('price_ars',0), p.get('price_usd',0), p.get('duration','6 noches'), p.get('includes_flight',True), p.get('fecha_ida',''), p.get('fecha_vuelta',''), p.get('fuente',''), p.get('url',''), p.get('notes',''), p.get('airline','')))
    conn.commit(); conn.close()

def get_flights_today():
    conn = sqlite3.connect(DB_PATH); conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM flights WHERE date(search_date)=date('now') ORDER BY price_usd ASC").fetchall()
    conn.close(); return [dict(r) for r in rows]

def get_packages_today():
    conn = sqlite3.connect(DB_PATH); conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM packages WHERE date(search_date)=date('now') ORDER BY price_usd ASC").fetchall()
    conn.close(); return [dict(r) for r in rows]

def get_historical():
    conn = sqlite3.connect(DB_PATH); conn.row_factory = sqlite3.Row
    rows = conn.execute('SELECT airline, MIN(price_usd) as min_usd, MAX(price_usd) as max_usd, AVG(price_usd) as avg_usd, COUNT(*) as total FROM flights GROUP BY airline ORDER BY avg_usd ASC').fetchall()
    conn.close(); return [dict(r) for r in rows]

def get_avg_price(destination, type='flight'):
    conn = sqlite3.connect(DB_PATH)
    avg = conn.execute('SELECT AVG(price_usd) FROM flights WHERE destination=?', (destination,)).fetchone()[0] if type=='flight' else conn.execute('SELECT AVG(price_usd) FROM packages').fetchone()[0]
    conn.close(); return avg or float('inf')
