#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Baja la planilla de conflictos territoriales y escribe conflictos.json.

Lo corre una GitHub Action una vez por día. El index.html lee ese JSON del mismo
origen, así que no hay problema de CORS: el CSV publicado de Google redirige a
googleusercontent.com, que no manda cabecera Access-Control-Allow-Origin, y por
eso el navegador nunca puede leerlo directo.

Sólo librería estándar, a propósito: la Action arranca en segundos y no depende
de nada que se pueda romper.

FORMATO DE SALIDA (desde esta versión):
  El JSON respeta las columnas del Excel tal cual están: cada fila es un objeto
  cuyas claves son los encabezados textuales de la planilla, en el MISMO orden
  de las columnas. No se renombra ni se adivina nada. La lista "columnas"
  repite los encabezados en orden para que el consumidor no dependa del orden
  de claves del JSON. El emparejamiento con los campos internos del mapa
  (provincia, lat, lon, etc.) lo hace el propio index.html al leer el archivo.

Qué NO publica:
  · La columna de contacto del vocero (y variantes: contacto, teléfono, mail,
    whatsapp). Se detecta por el encabezado y se excluye entera. Son referentes
    de conflictos con represión y desalojos, y esto termina en un archivo
    público.
  · Teléfonos, mails y links de WhatsApp que aparezcan en CUALQUIER campo. Es
    una segunda línea de defensa: si un encabezado con coma quedó mal
    entrecomillado, las columnas se corren y el teléfono puede terminar en
    "Observaciones".

La geolocalización (centroide del departamento, dispersión de los que comparten
uno) la hace el JS del mapa con la misma geometría que dibuja, así que acá sólo
se sanea texto.
"""

import csv
import io
import json
import os
import re
import sys
import unicodedata
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

SHEET_CSV_URL = os.environ.get("SHEET_CSV_URL", "").strip() or (
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTMoLBFDGIitw7vbPQ23M6FkxWqk8PmZIjyGDbIBGBoqkLiD5KATD-ueREzbp-9pp1vdg4sZ_V4D6Am/pub?gid=1240607916&single=true&output=csv"
)
SALIDA = Path(os.environ.get("SALIDA", "conflictos.json"))

# Si la fila de encabezados no es detectable (o se quiere fijar a mano), se
# puede pasar FILA_ENCABEZADO=2 (numeración humana, empezando en 1).
FILA_ENCABEZADO = os.environ.get("FILA_ENCABEZADO", "").strip()

# Columnas que NUNCA se publican. Se excluye la columna entera si su encabezado
# normalizado contiene alguna de estas palabras como palabra completa.
PALABRAS_EXCLUIDAS = ('contacto', 'vocero', 'telefono', 'celular', 'email',
                      'mail', 'whatsapp')

# Columnas de coordenadas: no pasan por la redacción de PII porque el patrón de
# teléfono se comería los decimales. Se reconocen por el encabezado normalizado
# EXACTO (acá sí exacto: "lat" suelto dentro de otra palabra sería lotería).
ENCABEZADOS_COORDENADA = ('lat', 'latitud', 'latitude', 'y',
                          'lon', 'lng', 'long', 'longitud', 'longitude', 'x')

# El patrón de teléfono no puede empezar ni terminar pegado a un dígito, punto,
# coma o guion. Sin eso se come las coordenadas: -26.0938761 entra como si fuera
# un teléfono y el conflicto pierde su ubicación exacta.
PATRONES_PII = [
    (re.compile(r'\bhttps?://(?:wa\.me|api\.whatsapp)\S*', re.I), '[dato de contacto omitido]'),
    (re.compile(r'\b(?:wa\.me|api\.whatsapp)\S*', re.I), '[dato de contacto omitido]'),
    (re.compile(r'[\w.+-]+@[\w-]+\.[\w.]+'), '[dato de contacto omitido]'),
    (re.compile(r'(^|[^\d.,\-])'
                r'((?:\+?54[\s\-]?)?(?:9[\s\-]?)?(?:\(?0?\d{2,4}\)?[\s\-]?)(?:15[\s\-]?)?\d{3,4}[\s\-]?\d{4})'
                r'(?![\d]|[.,\-]\d)'), r'\1[dato de contacto omitido]'),
]


def normalizar(texto):
    if not isinstance(texto, str):
        return ""
    sin_acentos = ''.join(c for c in unicodedata.normalize('NFD', texto)
                          if unicodedata.category(c) != 'Mn')
    return sin_acentos.lower().strip()


def redactar_pii(valor):
    for pat, reemplazo in PATRONES_PII:
        valor = pat.sub(reemplazo, valor)
    return valor


def es_columna_excluida(encabezado_norm):
    return any(re.search(r'(?<![a-z0-9])' + p + r'(?![a-z0-9])', encabezado_norm)
               for p in PALABRAS_EXCLUIDAS)


def es_columna_coordenada(encabezado_norm):
    return encabezado_norm in ENCABEZADOS_COORDENADA


def elegir_fila_encabezado(filas, max_filas=5):
    """Devuelve el índice de la fila de encabezados.

    La fila 1 de la planilla suele traer títulos de sección fusionados ("DATOS
    VISIBLES EN EL MAPA", "GEORREFERENCIACION") y los encabezados reales están
    en la 2. Regla determinística, sin adivinar alias: es encabezado la primera
    fila que tiene una celda cuyo texto normalizado es exactamente "provincia",
    o que la contiene como palabra completa. Si ninguna cumple, se usa la
    primera fila. FILA_ENCABEZADO en el entorno pisa todo.
    """
    if FILA_ENCABEZADO:
        return max(0, int(FILA_ENCABEZADO) - 1)

    tope = min(max_filas, len(filas))
    for i in range(tope):
        if any(normalizar(c) == 'provincia' for c in filas[i]):
            return i
    pat = re.compile(r'(?<![a-z0-9])provincia(?![a-z0-9])')
    for i in range(tope):
        if any(pat.search(normalizar(c)) for c in filas[i]):
            return i
    return 0


def encabezados_unicos(fila_encabezado):
    """Devuelve los encabezados tal cual el Excel, en orden, sin repetidos.

    · Se respeta el texto original (con acentos, mayúsculas y aclaraciones),
      sólo recortando espacios en los bordes.
    · Una celda vacía se nombra "columna_N" (N = posición, desde 1) para no
      perder la columna ni romper el JSON.
    · Un encabezado repetido recibe sufijo " (2)", " (3)"… porque las claves de
      un objeto JSON no pueden repetirse.
    """
    vistos, salida = {}, []
    for i, crudo in enumerate(fila_encabezado):
        nombre = (crudo or '').strip() or f'columna_{i + 1}'
        base = nombre
        n = vistos.get(base, 0) + 1
        vistos[base] = n
        if n > 1:
            nombre = f'{base} ({n})'
        salida.append(nombre)
    return salida


def bajar_csv(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (PRIHA bot)'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read().decode('utf-8-sig', errors='replace')


def sanear(texto_csv):
    """Devuelve (columnas_publicadas, filas) respetando el Excel.

    Las claves de cada fila son los encabezados textuales de la planilla, en el
    mismo orden de las columnas. Sólo se excluyen las columnas sensibles y se
    redacta PII en los valores.
    """
    filas = [f for f in csv.reader(io.StringIO(texto_csv)) if any(str(c).strip() for c in f)]
    if len(filas) < 2:
        raise SystemExit("El sheet vino vacío o sin filas de datos.")

    fh = elegir_fila_encabezado(filas)
    if fh:
        print(f"  encabezados leídos de la fila {fh + 1} del CSV")

    encabezados = encabezados_unicos(filas[fh])
    norm = [normalizar(h) for h in encabezados]

    publicables = []   # [(índice de columna, encabezado original)]
    excluidas = []
    for i, h in enumerate(encabezados):
        if es_columna_excluida(norm[i]):
            excluidas.append(h)
        else:
            publicables.append((i, h))

    if excluidas:
        print(f"  columnas sensibles ignoradas (no se publican): {excluidas}")

    if not any(normalizar(h) == 'provincia' or 'provincia' in normalizar(h)
               for _, h in publicables):
        raise SystemExit(f"No encontré la columna Provincia. Encabezados: {encabezados}")

    print("  columnas publicadas (en el orden del Excel):")
    for _, h in publicables:
        print(f"    {h!r}")

    salida, redactadas = [], 0
    for fila in filas[fh + 1:]:
        reg = {}
        for i, h in publicables:
            crudo = fila[i].strip() if i < len(fila) else ''
            if es_columna_coordenada(norm[i]):
                reg[h] = crudo
                continue
            limpio = redactar_pii(crudo)
            if limpio != crudo:
                redactadas += 1
            reg[h] = limpio
        if any(reg.values()):
            salida.append(reg)

    print(f"  {len(salida)} filas saneadas sobre {len(publicables)} columnas publicadas")
    if redactadas:
        print(f"  {redactadas} valor(es) con teléfono o email redactados. "
              f"Suele indicar columnas corridas en la planilla.")
    return [h for _, h in publicables], salida


def main():
    print(f"Bajando {SHEET_CSV_URL[:80]}…")
    columnas, filas = sanear(bajar_csv(SHEET_CSV_URL))

    # No pisar datos buenos con una respuesta vacía por un problema pasajero.
    if not filas:
        if SALIDA.exists():
            raise SystemExit("El sheet no devolvió filas; dejo el conflictos.json anterior.")
        raise SystemExit("El sheet no devolvió filas y no hay archivo previo.")

    nuevo = {
        "actualizado": datetime.now(timezone.utc).isoformat(timespec='seconds').replace('+00:00', 'Z'),
        "columnas": columnas,
        "filas": filas,
    }

    # Si sólo cambió la marca de tiempo, no commitear: evita un commit por día
    # sin contenido real.
    if SALIDA.exists():
        try:
            previo = json.loads(SALIDA.read_text(encoding='utf-8'))
            if previo.get("filas") == filas and previo.get("columnas") == columnas:
                print("Sin cambios en los datos; no reescribo el archivo.")
                return
        except Exception:
            pass

    SALIDA.write_text(json.dumps(nuevo, ensure_ascii=False, indent=1), encoding='utf-8')
    print(f"Escrito {SALIDA} ({len(filas)} filas)")


if __name__ == "__main__":
    sys.exit(main())
