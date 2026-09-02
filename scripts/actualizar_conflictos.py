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
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTMoLBFDGIitw7vbPQ23M6FkxWqk8PmZIjyGDbIBGBoqkLiD5KATD-ueREzbp-9pp1vdg4sZ_V4D6Am/pub?gid=1240607916&single=true&output=csv"
)
SALIDA = Path(os.environ.get("SALIDA", "conflictos.json"))

# Alias tolerantes de encabezados: se comparan sin acentos y en minúscula.
SHEET_COLS = {
    'provincia':     ('provincia',),
    'departamento':  ('departamento', 'depto', 'partido'),
    'localidad':     ('localidad', 'paraje', 'lugar'),
    # columnas que escribe el Apps Script de Georef en la planilla
    'provincia_norm':    ('provincia_norm',),
    'departamento_norm': ('departamento_norm',),
    'localidad_norm':    ('localidad_norm',),
    'lat':           ('lat', 'latitud', 'latitude', 'y'),
    'lon':           ('lon', 'lng', 'long', 'longitud', 'longitude', 'x'),
    'precision':     ('precision', 'precision_geo', 'geo_precision'),
    'comunidad':     ('comunidad',),
    'inicio':        ('inicio del conflicto', 'inicio'),
    'familias':      ('cantidad de famlias afectadas',      # typo en la planilla
                      'cantidad de familias afectadas', 'familias'),
    'hectareas':     ('hectareas afectadas', 'hectareas', 'superficie'),
    'motivo':        ('motivo del conflicto', 'motivo'),
    'grupo':         ('grupo economico/empresario', 'grupo economico', 'empresa'),
    'observaciones': ('observaciones (represion, orden desalojo)', 'observaciones'),
    'estado':        ('estado',),
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


def mapear_encabezados(headers):
    """Empareja los encabezados del sheet con los campos internos.

    No alcanza con comparar por igualdad: la planilla se edita a mano y los
    títulos vienen con el nombre de la sección pegado adelante ("DATOS VISIBLES
    EN EL MAPA Provincia") o con aclaraciones atrás ("ESTADO (activo / latente /
    resuelto)"). Primero se busca igualdad exacta y después que el encabezado
    CONTENGA el alias como palabra entera. El segundo paso sólo vale para alias
    de 5 letras o más: con 'lat' o 'x' la coincidencia por contenido sería una
    lotería. Una columna ya asignada no se vuelve a ofrecer.

    `headers` viene ya normalizado (sin acentos, en minúscula, sin espacios
    alrededor). Devuelve {campo: índice de columna}.
    """
    idx, usados = {}, set()

    for destino, alias in SHEET_COLS.items():
        for a in alias:
            for i, h in enumerate(headers):
                if i not in usados and h == a:
                    idx[destino] = i
                    usados.add(i)
                    break
            if destino in idx:
                break

    for destino, alias in SHEET_COLS.items():
        if destino in idx:
            continue
        mejor = None
        for a in alias:
            if len(a) < 5:
                continue
            pat = re.compile(r'(?<![a-z0-9])' + re.escape(a) + r'(?![a-z0-9])')
            for i, h in enumerate(headers):
                if i in usados or not pat.search(h):
                    continue
                # ante varios, gana el encabezado más corto: es el más específico
                if mejor is None or len(h) < len(headers[mejor]):
                    mejor = i
        if mejor is not None:
            idx[destino] = mejor
            usados.add(mejor)

    return idx

def elegir_fila_encabezado(filas, max_filas=5):
    """Devuelve el índice de la fila que mejor funciona como encabezado.

    La fila 1 de la planilla son títulos de sección fusionados ("DATOS VISIBLES
    EN EL MAPA", "GEORREFERENCIACION") y los encabezados reales están en la 2.
    Fijar el número se rompe en cuanto alguien agrega o saca una fila, así que
    se prueban las primeras y gana la que reconoce más columnas.
    """
    mejor, mejor_puntaje = 0, -1
    for i in range(min(max_filas, len(filas))):
        idx = mapear_encabezados([normalizar(h) for h in filas[i]])
        puntaje = len(idx) + (5 if 'provincia' in idx else 0)
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = i, puntaje
    return mejor

def bajar_csv(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (PRIHA bot)'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read().decode('utf-8-sig', errors='replace')


def sanear(texto_csv):
    filas = [f for f in csv.reader(io.StringIO(texto_csv)) if any(str(c).strip() for c in f)]
    if len(filas) < 2:
        raise SystemExit("El sheet vino vacío o sin filas de datos.")

    fh = elegir_fila_encabezado(filas)
    if fh:
        print(f"  encabezados leídos de la fila {fh + 1} del CSV")
    headers = [normalizar(h) for h in filas[fh]]

    sensibles = [filas[fh][i] for i, h in enumerate(headers) if h in SHEET_COLS_EXCLUIDAS]
    if sensibles:
        print(f"  columnas sensibles ignoradas (no se publican): {sensibles}")

    idx = mapear_encabezados(headers)
    if 'provincia' not in idx:
        raise SystemExit(f"No encontré la columna Provincia. Encabezados: {filas[fh]}")

    print("  columnas reconocidas:")
    for destino in SHEET_COLS:
        if destino in idx:
            print(f"    {destino:<14} <- {filas[fh][idx[destino]].strip()!r}")
    ignoradas = [filas[fh][i].strip() for i in range(len(filas[fh])) if i not in set(idx.values())]
    if ignoradas:
        print(f"  encabezados sin usar: {ignoradas}")

    salida, redactadas = [], 0
    for fila in filas[fh + 1:]:
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
