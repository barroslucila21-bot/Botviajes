import asyncio

async def scrape_viajes():
    print("Iniciando scraper (Modo Simulación Nube)...")
    resultados_vuelos = []
    resultados_paquetes = []
    resultados_alojamientos = []
    
    # Simulamos el tiempo de navegación que tomará en Railway
    print("Navegando Kayak y Almundo...")
    await asyncio.sleep(5)
    
    return resultados_vuelos, resultados_paquetes, resultados_alojamientos

if __name__ == "__main__":
    asyncio.run(scrape_viajes())
