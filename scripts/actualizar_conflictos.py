#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Baja la planilla de conflictos territoriales y escribe conflictos.json.
Lo corre una GitHub Action una vez por día. El index.html lee ese JSON del mismo
origen, así que no hay problema de CORS: el CSV publicado de Google redirige a
googleusercontent.com, que no manda cabecera Access-Control-Allow-Origin, y por
eso el navegador nunca puede leerlo directo.
Sólo librería estándar, a propósito: la Action arranca en segundos y no depende
de nada que se pueda romper.
Qué NO publica:
  · La columna de contacto del vocero (y variantes). No está en SHEET_COLS, así
    que ni se lee. Son referentes de conflictos con represión y desalojos, y
    esto termina en un archivo público.
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
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQX-bXajP4517XhNPSYbGiEWIGzEJTBB"
    "_8CYe-kq7Th1giq6osVVsV2VghWoedipxwTUNMozyCdwv3P/pub"
    "?gid=1400968737&single=true&output=csv"
)
SALIDA = Path(os.environ.get("SALIDA", "conflictos.json"))

# Fila de encabezados dentro de la planilla (1-indexado, como se ve en el Sheet).
# La fila 1 de "Hoja 1" son títulos de sección fusionados ("DATOS VISIBLES EN EL
# MAPA", "DATOS INVISIBLES", "GEORREFERENCIACION"); los encabezados reales de
# columna (Provincia, Departamento, ..., provincia_norm ... georef_fecha) están
# en la fila 2, y los datos arrancan en la fila 3. Mismo ajuste que se hizo en
# el Apps Script de georreferenciación (ver claude/georef-apps-script.md).
FILA_ENCABEZADO = int(os.environ.get("FILA_ENCABEZADO", "2"))
IDX_ENCABEZADO = FILA_ENCABEZADO - 1  # índice 0-based dentro de las filas del CSV

# Alias tolerantes de encabezados: se comparan sin acentos y en minúscula.
SHEET_COLS = {
    'provincia':     ('provincia',),
    'departamento':  ('departamento', 'depto', 'partido'),
    'localidad':     ('localidad', 'paraje', 'lugar'),
    'lat':           ('lat', 'latitud', 'latitude', 'y'),
    'lon':           ('lon', 'lng', 'long', 'longitud', 'longitude', 'x'),
    'precision':     ('precision', 'precision_geo', 'geo_precision'),
    'comunidad':     ('comunidad / caso', 'comunidad/caso', 'comunidad', 'caso'),
    'inicio':        ('inicio del conflicto', 'inicio'),
    'familias':      ('cantidad de famlias afectadas',      # typo en la planilla
                      'cantidad de familias afectadas', 'familias'),
    'hectareas':     ('hectareas afectadas', 'hectareas', 'superficie'),
    'motivo':        ('motivo del conflicto', 'motivo'),
    'grupo':         ('conflicto con: (grupo economico / empresario)',
                      'conflicto con (grupo economico / empresario)',
                      'grupo economico/empresario', 'grupo economico', 'empresa'),
    'observaciones': ('observaciones (represion, orden desalojo)', 'observaciones'),
    'estado':        ('estado (activo / latente / resuelto)', 'estado'),
    'fuente':        ('fuente de informacion', 'fuente'),
    'enlace':        ('link', 'enlace', 'url'),
    'juzgado':       ('juzgado/ dependencia judicial', 'juzgado/dependencia judicial', 'juzgado'),
    'otra_info':     ('otra info del conflicto (del caso/judicial)', 'otra info del conflicto'),
    'inai':          ('situacion del relevamiento de inai (finalizado, en tramite, sin relevar)',
                      'situacion del relevamiento de inai', 'inai'),
}
SHEET_COLS_EXCLUIDAS = ('contacto con vocero del conflicto', 'contacto con vocero',
                        'contacto', 'vocero', 'telefono', 'teléfono', 'email', 'mail')

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
# Columnas que nunca se redactan: son coordenadas, no texto libre.
COLS_SIN_REDACTAR = ('lat', 'lon')


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


def bajar_csv(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (PRIHA bot)'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read().decode('utf-8-sig', errors='replace')


def sanear(texto_csv):
    filas = [f for f in csv.reader(io.StringIO(texto_csv)) if any(str(c).strip() for c in f)]
    if len(filas) <= IDX_ENCABEZADO:
        raise SystemExit(
            f"El sheet no tiene ni la fila de encabezados (fila {FILA_ENCABEZADO}); "
            f"llegaron {len(filas)} fila(s) no vacías.")

    headers_crudos = filas[IDX_ENCABEZADO]
    headers = [normalizar(h) for h in headers_crudos]
    filas_datos = filas[IDX_ENCABEZADO + 1:]
    if not filas_datos:
        raise SystemExit("El sheet no tiene filas de datos debajo del encabezado.")

    sensibles = [headers_crudos[i] for i, h in enumerate(headers) if h in SHEET_COLS_EXCLUIDAS]
    if sensibles:
        print(f"  columnas sensibles ignoradas (no se publican): {sensibles}")

    idx = {}
    for destino, alias in SHEET_COLS.items():
        for a in alias:
            if a in headers:
                idx[destino] = headers.index(a)
                break
    if 'provincia' not in idx:
        raise SystemExit(
            f"No encontré la columna Provincia en la fila {FILA_ENCABEZADO}. "
            f"Encabezados: {headers_crudos}")

    salida, redactadas = [], 0
    for fila in filas_datos:
        reg = {}
        for destino, i in idx.items():
            crudo = fila[i].strip() if i < len(fila) else ''
            if destino in COLS_SIN_REDACTAR:
                reg[destino] = crudo
                continue
            limpio = redactar_pii(crudo)
            if limpio != crudo:
                redactadas += 1
            reg[destino] = limpio
        if any(reg.values()):
            salida.append(reg)

    print(f"  {len(salida)} filas saneadas sobre {len(idx)} columnas útiles")
    if redactadas:
        print(f"  {redactadas} valor(es) con teléfono o email redactados. "
              f"Suele indicar columnas corridas en la planilla.")
    return salida


def main():
    print(f"Bajando {SHEET_CSV_URL[:80]}…")
    filas = sanear(bajar_csv(SHEET_CSV_URL))
    # No pisar datos buenos con una respuesta vacía por un problema pasajero.
    if not filas:
        if SALIDA.exists():
            raise SystemExit("El sheet no devolvió filas; dejo el conflictos.json anterior.")
        raise SystemExit("El sheet no devolvió filas y no hay archivo previo.")
    nuevo = {
        "actualizado": datetime.now(timezone.utc).isoformat(timespec='seconds').replace('+00:00', 'Z'),
        "filas": filas,
    }
    # Si sólo cambió la marca de tiempo, no commitear: evita un commit por día
    # sin contenido real.
    if SALIDA.exists():
        try:
            previo = json.loads(SALIDA.read_text(encoding='utf-8'))
            if previo.get("filas") == filas:
                print("Sin cambios en los datos; no reescribo el archivo.")
                return
        except Exception:
            pass
    SALIDA.write_text(json.dumps(nuevo, ensure_ascii=False, indent=1), encoding='utf-8')
    print(f"Escrito {SALIDA} ({len(filas)} filas)")


if __name__ == "__main__":
    sys.exit(main())
