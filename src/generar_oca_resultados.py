import csv
import unicodedata
import re


INPUT_CSV = "via_cargo_input.csv"
ZONAS_CSV = "oca_zonas.csv"
OUTPUT_CSV = "oca_resultados_ok.csv"


# Tarifas OCA vigentes del PDF 02/2026 - Vigencia 08/04/2026
# Sección comparable: Entrega en sucursal.
# Columnas: Local, Regional, Nacional 1, Nacional 2.
TARIFAS_OCA_SUCURSAL = [
    {"desde": 0, "hasta": 1, "Local": 13080, "Regional": 13580, "Nacional 1": 18260, "Nacional 2": 18900},
    {"desde": 1, "hasta": 2, "Local": 13580, "Regional": 15290, "Nacional 1": 18660, "Nacional 2": 20520},
    {"desde": 2, "hasta": 5, "Local": 14030, "Regional": 16970, "Nacional 1": 19060, "Nacional 2": 22110},
    {"desde": 5, "hasta": 10, "Local": 16520, "Regional": 22640, "Nacional 1": 23410, "Nacional 2": 27520},
    {"desde": 10, "hasta": 15, "Local": 20400, "Regional": 29580, "Nacional 1": 36020, "Nacional 2": 42610},
    {"desde": 15, "hasta": 20, "Local": 23290, "Regional": 33880, "Nacional 1": 37290, "Nacional 2": 44080},
    {"desde": 20, "hasta": 25, "Local": 26110, "Regional": 37910, "Nacional 1": 50930, "Nacional 2": 60550},
]

PRECIO_KG_EXCEDENTE = {
    "Local": 640,
    "Regional": 1030,
    "Nacional 1": 1200,
    "Nacional 2": 1950,
}


def normalizar_texto(valor):
    """
    Convierte 'Córdoba', 'Cordoba', 'CORDOBA' en una misma clave: 'cordoba'.
    También normaliza espacios.
    """
    texto = str(valor).strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = texto.lower()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def leer_zonas():
    zonas = {}

    with open(ZONAS_CSV, newline="", encoding="utf-8-sig") as archivo:
        reader = csv.DictReader(archivo)

        for row in reader:
            destino = row["destino"].strip()
            zona = row["zona_oca"].strip()
            zonas[normalizar_texto(destino)] = zona

    return zonas


def calcular_precio_oca(kg, zona):
    """
    Calcula precio OCA de Entrega en Sucursal.

    Para kg hasta 25, toma el rango tarifario.
    Para kg mayor a 25, toma el precio de 20 a 25 kg
    y suma kg excedente por cada kg adicional.
    """
    kg = float(kg)

    if kg <= 25:
        for fila in TARIFAS_OCA_SUCURSAL:
            desde = fila["desde"]
            hasta = fila["hasta"]

            if kg > desde and kg <= hasta:
                return fila[zona]

            if kg == 0 and desde == 0:
                return fila[zona]

    tarifa_base_25 = TARIFAS_OCA_SUCURSAL[-1][zona]
    excedente = kg - 25
    precio_excedente = excedente * PRECIO_KG_EXCEDENTE[zona]

    return tarifa_base_25 + precio_excedente


def leer_input():
    with open(INPUT_CSV, newline="", encoding="utf-8-sig") as archivo:
        return list(csv.DictReader(archivo))


def guardar_resultados(resultados):
    columnas = [
        "destino",
        "kg",
        "zona_oca",
        "precio_oca_sucursal",
        "estado",
        "observacion",
    ]

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as archivo:
        writer = csv.DictWriter(archivo, fieldnames=columnas, delimiter=";")
        writer.writeheader()
        writer.writerows(resultados)


def main():
    zonas = leer_zonas()
    filas = leer_input()

    resultados = []

    for fila in filas:
        destino = fila["destino"].strip()
        destino_key = normalizar_texto(destino)
        kg = float(fila["kg"])

        if destino_key not in zonas:
            resultados.append({
                "destino": destino,
                "kg": int(kg) if kg.is_integer() else kg,
                "zona_oca": "",
                "precio_oca_sucursal": "",
                "estado": "ERROR",
                "observacion": "Destino sin zona OCA mapeada",
            })
            continue

        zona = zonas[destino_key]
        precio = calcular_precio_oca(kg, zona)

        resultados.append({
            "destino": destino,
            "kg": int(kg) if kg.is_integer() else kg,
            "zona_oca": zona,
            "precio_oca_sucursal": round(precio, 2),
            "estado": "OK",
            "observacion": "",
        })

    guardar_resultados(resultados)

    errores = [r for r in resultados if r["estado"] != "OK"]

    print("Proceso terminado.")
    print(f"Filas generadas: {len(resultados)}")
    print(f"Errores: {len(errores)}")
    print(f"Archivo generado: {OUTPUT_CSV}")

    if errores:
        print("\nFilas con error:")
        for error in errores:
            print(error)


if __name__ == "__main__":
    main()