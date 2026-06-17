from playwright.sync_api import sync_playwright
import csv
import time
import re

URL = "https://viacargo.com.ar/cotizar-envio/"

INPUT_CSV = "via_cargo_input.csv"
OUTPUT_CSV = "via_cargo_resultados_ok.csv"


def calcular_dimensiones_sin_aforo(kg):
    """
    Calcula dimensiones en cm para que el peso volumétrico quede por debajo del peso real.

    Fórmula:
    kg_aforado = alto * ancho * profundidad * 350 / 1.000.000
    """
    volumen_cm3_max = kg * 1_000_000 / 350
    volumen_objetivo = volumen_cm3_max * 0.85
    lado = int(volumen_objetivo ** (1 / 3))
    lado = max(lado, 5)
    return lado, lado, lado


def obtener_frame_cotizador(page):
    """
    Busca el iframe donde está embebido el cotizador de Vía Cargo.
    """
    page.wait_for_selector("iframe", state="attached", timeout=60000)

    for intento in range(60):
        for frame in page.frames:
            if "formularios.viacargo.com.ar" in frame.url:
                return frame

        time.sleep(1)

    raise Exception("No pude encontrar el iframe de formularios.viacargo.com.ar")


def seleccionar_autocomplete_por_cp(frame, selector_input, cp, descripcion):
    """
    Escribe el CP y selecciona la opción del autocompletado que contiene ese CP.
    """
    print(f"Completando {descripcion} con CP {cp}...")

    campo = frame.locator(selector_input)
    campo.click()
    campo.fill("")
    time.sleep(0.5)
    campo.type(str(cp), delay=120)

    time.sleep(4)

    opciones = frame.locator("[role='option']").all()
    opcion_elegida = None

    print(f"Opciones encontradas para {descripcion}:")
    for i, opcion in enumerate(opciones):
        texto = opcion.inner_text(timeout=3000).strip()
        print(i, texto)

        if f"({cp})" in texto or str(cp) in texto:
            opcion_elegida = opcion
            break

    if opcion_elegida is None:
        texto_iframe = frame.locator("body").inner_text(timeout=5000)
        print("Texto visible del iframe:")
        print(texto_iframe)
        raise Exception(f"No encontré una opción que contenga el CP {cp} para {descripcion}")

    opcion_elegida.click()
    time.sleep(1)


def normalizar_precio(valor_txt):
    """
    Normaliza precios con formato argentino o web.

    Ejemplos:
    "14.999,99" -> "14999.99"
    "14999.99"  -> "14999.99"
    "14000"     -> "14000"
    """
    valor_txt = valor_txt.strip()

    if "," in valor_txt:
        valor_num = valor_txt.replace(".", "").replace(",", ".")
    else:
        valor_num = valor_txt

    return valor_num


def extraer_valores(texto):
    """
    Extrae los precios por tipo de producto desde el texto del resultado.
    """

    productos = {
        "ViaCargo - ENTREGA A DOMICILIO": "via_cargo_entrega_domicilio",
        "ViaCargo - DESPACHO AGENCIA - ENTREGA DOMICILIO": "via_cargo_agencia_domicilio",
        "ViaCargo - RETIRO DOMICILIO - ENTREGA AGENCIA": "via_cargo_domicilio_agencia",
        "ViaCargo - DESPACHO AGENCIA - ENTREGA AGENCIA": "via_cargo_agencia_agencia",
    }

    resultado = {
        "via_cargo_entrega_domicilio": "",
        "via_cargo_agencia_domicilio": "",
        "via_cargo_domicilio_agencia": "",
        "via_cargo_agencia_agencia": "",
    }

    for nombre_producto, columna in productos.items():
        patron = re.escape(nombre_producto) + r".*?Valor\s*\n\s*\$?([0-9\.\,]+)"
        match = re.search(patron, texto, flags=re.DOTALL)

        if match:
            valor_txt = match.group(1)
            resultado[columna] = normalizar_precio(valor_txt)

    return resultado


def cotizar_una_fila(page, fila):
    """
    Cotiza una fila del CSV de entrada.
    """
    print("\n======================================")
    print(f"Cotizando destino: {fila['destino']} - CP {fila['destino_cp']}")
    print("======================================")

    page.goto(URL, wait_until="domcontentloaded", timeout=60000)
    time.sleep(8)

    frame = obtener_frame_cotizador(page)
    frame.wait_for_selector("#mat-input-0", state="visible", timeout=60000)

    origen_cp = fila["origen_cp"]
    destino_cp = fila["destino_cp"]
    destino = fila["destino"]
    bultos = int(fila["bultos"])
    kg = float(fila["kg"])
    valor_declarado = int(float(fila["valor_declarado"]))

    alto, ancho, profundidad = calcular_dimensiones_sin_aforo(kg)

    seleccionar_autocomplete_por_cp(frame, "#mat-input-0", origen_cp, "origen")
    seleccionar_autocomplete_por_cp(frame, "#mat-input-1", destino_cp, "destino")

    print("Completando paquete...")
    frame.locator("#mat-input-2").fill(str(bultos))
    frame.locator("#mat-input-3").fill(str(int(kg)))
    frame.locator("#mat-input-4").fill(str(alto))
    frame.locator("#mat-input-5").fill(str(ancho))
    frame.locator("#mat-input-6").fill(str(profundidad))
    frame.locator("#mat-input-7").fill(str(valor_declarado))

    print(f"Dimensiones usadas: {alto} x {ancho} x {profundidad} cm")

    print("Seleccionando pago en origen...")
    frame.locator("#mat-radio-2-input").check(force=True)

    print("Haciendo click en Cotizá...")
    frame.get_by_role("button", name="Cotizá").click()

    print("Esperando resultado...")
    time.sleep(12)

    texto = frame.locator("body").inner_text(timeout=10000)
    valores = extraer_valores(texto)

    resultado = {
        "origen_cp": origen_cp,
        "destino_cp": destino_cp,
        "destino": destino,
        "bultos": bultos,
        "kg": int(kg),
        "alto_cm": alto,
        "ancho_cm": ancho,
        "profundidad_cm": profundidad,
        "valor_declarado": valor_declarado,
        **valores,
        "estado": "OK",
        "observacion": "",
    }

    print("Resultado extraído:")
    print(resultado)

    return resultado


def leer_input():
    """
    Lee el CSV de entrada separado por coma.
    """
    with open(INPUT_CSV, newline="", encoding="utf-8-sig") as archivo:
        return list(csv.DictReader(archivo))


def guardar_resultados(resultados):
    """
    Guarda el CSV de salida separado por punto y coma para que Excel lo abra bien.
    """
    columnas = [
        "origen_cp",
        "destino_cp",
        "destino",
        "bultos",
        "kg",
        "alto_cm",
        "ancho_cm",
        "profundidad_cm",
        "valor_declarado",
        "via_cargo_entrega_domicilio",
        "via_cargo_agencia_domicilio",
        "via_cargo_domicilio_agencia",
        "via_cargo_agencia_agencia",
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
        browser = p.chromium.launch(headless=False, slow_mo=400)
        page = browser.new_page()

        for fila in filas:
            try:
                resultado = cotizar_una_fila(page, fila)
                resultados.append(resultado)
                guardar_resultados(resultados)

            except Exception as e:
                print("ERROR en fila:", fila)
                print(e)

                resultados.append({
                    "origen_cp": fila.get("origen_cp", ""),
                    "destino_cp": fila.get("destino_cp", ""),
                    "destino": fila.get("destino", ""),
                    "bultos": fila.get("bultos", ""),
                    "kg": fila.get("kg", ""),
                    "alto_cm": "",
                    "ancho_cm": "",
                    "profundidad_cm": "",
                    "valor_declarado": fila.get("valor_declarado", ""),
                    "via_cargo_entrega_domicilio": "",
                    "via_cargo_agencia_domicilio": "",
                    "via_cargo_domicilio_agencia": "",
                    "via_cargo_agencia_agencia": "",
                    "estado": "ERROR",
                    "observacion": str(e),
                })

                guardar_resultados(resultados)

        input("\nProceso terminado. Presioná ENTER para cerrar navegador...")
        browser.close()


if __name__ == "__main__":
    main()