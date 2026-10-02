"""
Genera datos.js a partir de productos.xlsx.

Uso:  python generar_catalogo.py

Si encuentra errores NO escribe nada y termina con codigo 1, para que
actualizar.bat no publique una pagina rota.
"""

import json
import math
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent
XLSX = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else RAIZ / "productos.xlsx"
SALIDA = RAIZ / "datos.js"

errores = []
avisos = []


def err(mensaje):
    errores.append(mensaje)


def aviso(mensaje):
    avisos.append(mensaje)


def limpiar(texto):
    """Quita tildes, pasa a minusculas y deja solo letras y digitos."""
    texto = str(texto).strip().lower()
    texto = (
        texto.replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ü", "u")
        .replace("ñ", "n")
    )
    texto = re.sub(r"[^a-z0-9]+", "-", texto)
    return texto.strip("-")


def texto(valor):
    """Convierte a texto y trata celdas vacias (NaN/None) como cadena vacia."""
    if valor is None:
        return ""
    if isinstance(valor, float) and math.isnan(valor):
        return ""
    return str(valor).strip()


def mostrar(valor):
    """Como texto pero las celdas vacias se ven asi en los mensajes."""
    return texto(valor) if texto(valor) else "(vacio)"


def a_precio(valor):
    """Acepta 88000, 88.000, $88.000 o '88 000'."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return None
    if isinstance(valor, (int, float)):
        return int(valor)
    digitos = re.sub(r"[^0-9]", "", str(valor))
    if not digitos:
        return None
    return int(digitos)


def a_lista(valor):
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return None
    partes = [p.strip() for p in str(valor).replace(";", ",").split(",")]
    partes = [p for p in partes if p]
    return partes or None


# ---------------------------------------------------------------- leer Excel
if not XLSX.exists():
    print(f"ERROR: no existe {XLSX.name} junto a este script.")
    sys.exit(1)

try:
    catalogo = pd.read_excel(XLSX, sheet_name="Catalogo")
    categorias = pd.read_excel(XLSX, sheet_name="Categorias")
except Exception as exc:  # noqa: BLE001
    print(f"ERROR: no se pudo abrir {XLSX.name}: {exc}")
    print("Revisa que el archivo no este abierto en Excel.")
    sys.exit(1)

faltan_columnas = {"categoria", "referencia", "foto", "precio", "stock"} - set(catalogo.columns)
if faltan_columnas:
    print("ERROR: en la hoja Catalogo faltan columnas:", ", ".join(sorted(faltan_columnas)))
    sys.exit(1)

catalogo = catalogo.where(catalogo.notna(), None)
categorias = categorias.where(categorias.notna(), None)

# --------------------------------------------------------------- categorias
ajustes = {}
for _, fila in categorias.iterrows():
    nombre = texto(fila.get("categoria"))
    if not nombre:
        continue
    orden = a_precio(fila.get("orden"))
    if orden is None:
        aviso(f"La categoria '{nombre}' no tiene orden; se pondra al final.")
        orden = 10_000
    ajustes[nombre] = {"orden": int(orden), "tallas": a_lista(fila.get("tallas"))}

if not ajustes:
    err("La hoja Categorias esta vacia.")

# ---------------------------------------------------------------- productos
vistos = {}
productos_por_categoria = {}

for num, fila in catalogo.iterrows():
    donde = f"Fila {num + 2}"
    categoria = texto(fila.get("categoria"))
    referencia = texto(fila.get("referencia"))
    foto = texto(fila.get("foto"))
    cantidad = a_precio(fila.get("fotos"))
    cantidad = 1 if cantidad is None else int(cantidad)
    precio = a_precio(fila.get("precio"))
    precio_base = a_precio(fila.get("precio_base"))
    stock = a_precio(fila.get("stock"))
    tallas_propias = a_lista(fila.get("tallas"))

    if not categoria:
        err(f"{donde}: falta la categoria.")
        continue
    if categoria not in ajustes:
        err(f"{donde}: la categoria '{categoria}' no esta en la hoja Categorias.")
        continue
    if not referencia:
        err(f"{donde} (categoria '{categoria}'): falta la referencia (nombre).")
        continue
    if not foto:
        err(f"{donde} ({referencia}): falta la columna foto.")
        continue
    if precio is None or precio <= 0:
        err(f"{donde} ({referencia}): el precio {mostrar(fila.get('precio'))} no es un numero valido.")
        continue
    if stock is None or stock < 0:
        err(f"{donde} ({referencia}): el stock {mostrar(fila.get('stock'))} no es un numero valido.")
        continue
    if cantidad < 1:
        err(f"{donde} ({referencia}): fotos debe ser 1 o mas.")
        continue

    slug = limpiar(foto)
    if not slug:
        err(f"{donde} ({referencia}): la foto '{foto}' no deja un nombre valido.")
        continue

    clave = (categoria, slug)
    if clave in vistos:
        err(f"{donde} ({referencia}): '{categoria}/{slug}' ya lo usa la fila {vistos[clave] + 2}.")
        continue
    vistos[clave] = num

    # nombres de archivo: 1 foto -> foto.jpg ; varias -> foto-1.jpg, foto-2.jpg
    if cantidad == 1:
        imagenes = [f"img/{limpiar(categoria)}/{slug}.jpg"]
    else:
        imagenes = [f"img/{limpiar(categoria)}/{slug}-{i}.jpg" for i in range(1, cantidad + 1)]

    tallas = tallas_propias or ajustes[categoria]["tallas"]

    producto = {
        "id": f"{limpiar(categoria)}-{len(productos_por_categoria.get(categoria, []))}",
        "nombre": referencia,
        "precio": precio,
        "precioBase": precio_base,
        "stock": int(stock),
        "tallas": tallas,
        "imagenes": imagenes,
    }
    productos_por_categoria.setdefault(categoria, []).append(producto)

    if precio_base is not None and precio_base >= precio:
        aviso(f"{referencia}: el precio base ({precio_base}) es mayor o igual al de venta ({precio}).")

if not productos_por_categoria:
    errores.append("No hay productos validos en la hoja Catalogo.")

# ------------------------------------------------- comprobar que existan fotos
faltantes = []
for categoria, productos in productos_por_categoria.items():
    for producto in productos:
        for imagen in producto["imagenes"]:
            if not (RAIZ / imagen).exists():
                faltantes.append(imagen)

for imagen in sorted(set(faltantes)):
    err(f"Falta la foto: {imagen}")

# ----------------------------------------------------- fotos sin usar
usadas = set()
for productos in productos_por_categoria.values():
    for producto in productos:
        usadas.update(producto["imagenes"])

sobrantes = []
for carpeta in sorted((RAIZ / "img").iterdir()):
    if not carpeta.is_dir():
        continue
    for archivo in carpeta.iterdir():
        ruta = f"img/{carpeta.name}/{archivo.name}"
        if ruta not in usadas and archivo.name != "logo.jpeg":
            sobrantes.append(ruta)

for ruta in sobrantes:
    aviso(f"La foto {ruta} no esta en el Excel, no se va a publicar.")

# ------------------------------------------------------------------ resumen
print()
print("=" * 62)
print("  CATALOGO DE JC DISTRICT")
print("=" * 62)

for categoria, productos in sorted(
    productos_por_categoria.items(), key=lambda x: (ajustes[x[0]]["orden"], x[0])
):
    precios = sorted({p["precio"] for p in productos})
    unidades = sum(p["stock"] for p in productos)
    rango = f"{precios[0]:,}".replace(",", ".")
    if len(precios) > 1:
        rango += f" a {precios[-1]:,}".replace(",", ".")
    fotos = sum(len(p["imagenes"]) for p in productos)
    print(
        f"  {categoria:<12} {len(productos):>3} referencias  {fotos:>3} fotos  "
        f"{rango:>15}  stock: {unidades}"
    )

print()
print(f"  Total: {sum(len(p) for p in productos_por_categoria.values())} productos, {len(usadas)} fotos")

if avisos:
    print()
    print("  AVISOS:")
    for texto in avisos:
        print("   -", texto)

if errores:
    print()
    print("  ERRORES (no se publico nada):")
    for texto in errores:
        print("   -", texto)
    print()
    print("=" * 62)
    sys.exit(1)

# --------------------------------------------------------------- escribir js
salida = {
    "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "categorias": [
        {
            "id": limpiar(categoria),
            "titulo": categoria,
            "tallas": ajustes[categoria]["tallas"],
            "productos": productos,
        }
        for categoria, productos in sorted(
            productos_por_categoria.items(), key=lambda x: (ajustes[x[0]]["orden"], x[0])
        )
    ],
}


def envolver(datos):
    return (
        "/* Generado por generar_catalogo.py desde productos.xlsx.\n"
        "   No editar a mano: tus cambios se pierden al volver a generar. */\n"
        "window.CATALOGO = " + json.dumps(datos, ensure_ascii=False, indent=2) + ";\n"
    )


contenido = envolver(salida)

# La hora de generacion no debe obligar a reescribir el archivo: si el
# catalogo es identico se conserva el datos.js que ya existe.
sin_cambios = False
if SALIDA.exists():
    anterior = SALIDA.read_text(encoding="utf-8")
    if anterior == contenido:
        sin_cambios = True
    else:
        try:
            previo = json.loads(anterior.split("window.CATALOGO = ", 1)[1].rstrip().rstrip(";"))
            previo.pop("generado", None)
            actual = dict(salida)
            actual.pop("generado", None)
            sin_cambios = previo == actual
        except (ValueError, IndexError, AttributeError):
            sin_cambios = False

if sin_cambios:
    print()
    print("  No habia cambios en el Excel, datos.js quedo igual.")
else:
    SALIDA.write_text(contenido, encoding="utf-8", newline="")
    print()
    print(f"  datos.js escrito ({len(contenido.splitlines())} lineas).")

print("=" * 62)
print()
