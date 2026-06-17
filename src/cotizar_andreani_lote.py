from playwright.sync_api import sync_playwright
import csv
import time
import re


URL = "https://www.andreani.com/?tab=cotizar-envio"

INPUT_CSV = "via_cargo_input.csv"
OUTPUT_CSV = "andreani_resultados_ok.csv"

# Andreani no encontró 1303. Usamos 1431 como origen CABA.
ORIGEN_CP_ANDREANI = "1431"

MAX_INTENTOS_POR_FILA = 2


def calcular_dimensiones_sin_aforo(kg):
    volumen_cm3_max = kg * 1_000_000 / 350
    volumen_objetivo = volumen_cm3_max * 0.85
    lado = int(volumen_objetivo ** (1 / 3))
    lado = max(lado, 5)
    return lado, lado, lado


def normalizar_precio_argentino(valor_txt):
    """
    "$16.410,62" -> 16410.62
    """
    valor_txt = str(valor_txt).strip()
    valor_txt = valor_txt.replace("$", "").replace(" ", "")
    valor_txt = valor_txt.replace(".", "").replace(",", ".")
    return float(valor_txt)


def extraer_precios_andreani(texto):
    """
    Busca precios en formato:
    A sucursal
    $37.865,96

    A domicilio
    $38.632,41
    """
    precio_sucursal = ""
    precio_domicilio = ""

    match_sucursal = re.search(
        r"A\s+sucursal\s*\n+\s*\$?\s*([0-9\.\,]+)",
        texto,
        flags=re.IGNORECASE
    )

    match_domicilio = re.search(
        r"A\s+domicilio\s*\n+\s*\$?\s*([0-9\.\,]+)",
        texto,
        flags=re.IGNORECASE
    )

    if match_sucursal:
        precio_sucursal = normalizar_precio_argentino(match_sucursal.group(1))

    if match_domicilio:
        precio_domicilio = normalizar_precio_argentino(match_domicilio.group(1))

    return precio_sucursal, precio_domicilio


def seleccionar_autocomplete(page, input_locator, texto_busqueda, descripcion):
    print(f"Completando {descripcion}: {texto_busqueda}")

    input_locator.click()
    input_locator.fill("")
    time.sleep(0.5)
    input_locator.type(str(texto_busqueda), delay=120)

    # Espera opciones del autocomplete
    time.sleep(4)

    opciones = page.locator("[role='option']").all()
    opcion_elegida = None

    print(f"Opciones encontradas para {descripcion}:")
    for i, opcion in enumerate(opciones):
        try:
            texto = opcion.inner_text(timeout=3000).strip()
            print(i, texto)

            if str(texto_busqueda) in texto:
                opcion_elegida = opcion
                break
        except Exception as e:
            print(i, "No pude leer opción:", e)

    if opcion_elegida is None:
        texto_body = page.locator("body").inner_text(timeout=5000)
        print("Texto visible al fallar autocomplete:")
        print(texto_body[:2000])
        raise Exception(f"No encontré opción para {descripcion}: {texto_busqueda}")

    opcion_elegida.click()
    time.sleep(1)


def esperar_resultado(page, timeout_segundos=35):
    """
    Espera hasta que aparezca el resultado de Andreani.
    """
    inicio = time.time()

    while time.time() - inicio < timeout_segundos:
        texto = page.locator("body").inner_text(timeout=10000)

        if "A sucursal" in texto and "$" in texto:
            return texto

        if "Tu cotización" in texto and "$" in texto:
            return texto

        time.sleep(2)

    texto = page.locator("body").inner_text(timeout=10000)
    return texto


def cotizar_una_fila(browser, fila):
    destino = str(fila["destino"])
    destino_cp = str(fila["destino_cp"])
    kg = float(fila["kg"])

    print("\n======================================")
    print(f"Cotizando Andreani: {destino} - CP {destino_cp} - {kg} kg")
    print("======================================")

    page = browser.new_page()

    try:
        page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        time.sleep(8)

        peso_gr = int(kg * 1000)
        alto, ancho, largo = calcular_dimensiones_sin_aforo(kg)

        inputs_cp = page.locator("input[placeholder='CP o Localidad']")
        input_desde = inputs_cp.nth(0)
        input_hasta = inputs_cp.nth(1)

        seleccionar_autocomplete(page, input_desde, ORIGEN_CP_ANDREANI, "origen")
        seleccionar_autocomplete(page, input_hasta, destino_cp, "destino")

        print("Completando medidas...")
        page.locator("input[name='Alto']").fill(str(alto))
        page.locator("input[name='Ancho']").fill(str(ancho))
        page.locator("input[name='Largo']").fill(str(largo))
        page.locator("input[name='Peso']").fill(str(peso_gr))

        print(f"Dimensiones usadas: {alto} x {ancho} x {largo} cm")
        print(f"Peso usado: {peso_gr} gr")

        print("Haciendo click en Cotizar...")
        page.get_by_role("button", name="Cotizar").click()

        print("Esperando resultado...")
        texto = esperar_resultado(page, timeout_segundos=40)

        precio_sucursal, precio_domicilio = extraer_precios_andreani(texto)

        if precio_sucursal == "":
            print("Texto completo donde no pude extraer precio:")
            print(texto[:4000])
            raise Exception("No pude extraer el precio A sucursal")

        resultado = {
            "destino": destino,
            "destino_cp": destino_cp,
            "kg": int(kg) if kg.is_integer() else kg,
            "alto_cm": alto,
            "ancho_cm": ancho,
            "largo_cm": largo,
            "peso_gr": peso_gr,
            "precio_andreani_sucursal": round(precio_sucursal, 2),
            "precio_andreani_domicilio": round(precio_domicilio, 2) if precio_domicilio != "" else "",
            "estado": "OK",
            "observacion": "",
        }

        print("Resultado extraído:")
        print(resultado)

        return resultado

    finally:
        page.close()


def leer_input():
    with open(INPUT_CSV, newline="", encoding="utf-8-sig") as archivo:
        return list(csv.DictReader(archivo))


def guardar_resultados(resultados):
    columnas = [
        "destino",
        "destino_cp",
        "kg",
        "alto_cm",
        "ancho_cm",
        "largo_cm",
        "peso_gr",
        "precio_andreani_sucursal",
        "precio_andreani_domicilio",
        "estado",
        "observacion",
    ]

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as archivo:
        writer = csv.DictWriter(archivo, fieldnames=columnas, delimiter=";")
        writer.writeheader()
        writer.writerows(resultados)


def main():
    filas = leer_input()
    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=350)

        for fila in filas:
            ultimo_error = None

            for intento in range(1, MAX_INTENTOS_POR_FILA + 1):
                try:
                    print(f"\nIntento {intento} de {MAX_INTENTOS_POR_FILA}")
                    resultado = cotizar_una_fila(browser, fila)
                    resultados.append(resultado)
                    guardar_resultados(resultados)
                    break

                except Exception as e:
                    ultimo_error = e
                    print("ERROR en intento:", intento)
                    print(e)

                    if intento < MAX_INTENTOS_POR_FILA:
                        print("Reintentando la misma fila...")
                        time.sleep(5)

            else:
                print("ERROR definitivo en fila:", fila)
                print(ultimo_error)

                resultados.append({
                    "destino": fila.get("destino", ""),
                    "destino_cp": fila.get("destino_cp", ""),
                    "kg": fila.get("kg", ""),
                    "alto_cm": "",
                    "ancho_cm": "",
                    "largo_cm": "",
                    "peso_gr": "",
                    "precio_andreani_sucursal": "",
                    "precio_andreani_domicilio": "",
                    "estado": "ERROR",
                    "observacion": str(ultimo_error),
                })

                guardar_resultados(resultados)

        input("\nProceso terminado. Presioná ENTER para cerrar navegador...")
        browser.close()


if __name__ == "__main__":
    main()