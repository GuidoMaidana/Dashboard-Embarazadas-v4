# -*- coding: utf-8 -*-
"""
generar_data_js.py — arma data.js (window.DASHBOARD_DATA) para el dashboard.

Entradas (en esta misma carpeta):
  - datos_salud_optimizados.csv   base limpia (consultas + establecimientos)
  - zonas.geojson                 polígonos de las áreas programáticas

Salida:
  - ../data.js                    datos del dashboard, en formato legible
                                  (se puede leer y editar como texto)

Uso:  python generar_data_js.py
"""
from __future__ import annotations

import csv
import json
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
CSV_IN = HERE / "datos_salud_optimizados.csv"
ZONAS_IN = HERE / "zonas.geojson"
JS_OUT = HERE.parent / "data.js"

SCALARS = (str, int, float, bool, type(None))
NL = "\n"


def to_num(value):
    """'' / NULL -> None · enteros sin decimales · el resto redondeado a 7 decimales."""
    if value is None or str(value).strip() in ("", "NULL"):
        return None
    try:
        f = float(value)
    except ValueError:
        return str(value)
    return int(f) if f.is_integer() else round(f, 7)


def prune(d):
    """Quita campos vacíos para que data.js no arrastre ruido (conserva 0)."""
    return {k: v for k, v in d.items() if v not in ("", None)}


def dump(value, ind=0):
    """Serializa a JSON legible: un campo por línea; listas de números en una línea."""
    pad, pad2 = "  " * ind, "  " * (ind + 1)
    if isinstance(value, dict):
        if not value:
            return "{}"
        parts = [f"{pad2}{json.dumps(k, ensure_ascii=False)}: {dump(v, ind + 1)}"
                 for k, v in value.items()]
        return "{" + NL + ("," + NL).join(parts) + NL + pad + "}"
    if isinstance(value, list):
        if not value:
            return "[]"
        if all(isinstance(x, SCALARS) for x in value):
            inline = "[" + ", ".join(json.dumps(x, ensure_ascii=False) for x in value) + "]"
            if len(inline) <= 100:
                return inline
        parts = [pad2 + dump(x, ind + 1) for x in value]
        return "[" + NL + ("," + NL).join(parts) + NL + pad + "]"
    return json.dumps(value, ensure_ascii=False)


def consulta_row(r):
    return prune({
        "id": r["id_registro"],
        "id_centro": r["id_centro"],
        "saps": r["nombre_centro"],
        "saps_original": r["nombre_original"],
        "zona": r["zona"],
        "barrio": r["barrio"],
        "periodo": r["periodo"],
        "mes_nombre": r["mes_nombre"],
        "fecha_iso": r["fecha_inicio_periodo"],
        "etapa_pandemia": r["etapa_pandemia"],
        "rango_etario": r["grupo_etario"],
        "edad_min": to_num(r["edad_min"]),
        "edad_max": to_num(r["edad_max"]),
        "consulta_cantidad": to_num(r["cantidad_consultas_mes"]),
        "observaciones": r["observaciones"],
    })


def saps_row(r):
    return prune({
        "id": r["id_registro"],
        "nombre": r["nombre_centro"],
        "numero_saps": r["numero_saps"],
        "barrio": r["barrio"],
        "zona": r["zona"],
        "direccion": r["direccion"],
        "lat": to_num(r["latitud"]),
        "lng": to_num(r["longitud"]),
        "consultas": to_num(r["consultas_totales_2020"]),
        "meses_con_datos": to_num(r["meses_con_registros_2020"]),
        "observaciones": r["observaciones"],
    })


def simple_row(r):
    return prune({
        "id": r["id_registro"],
        "nombre": r["nombre_centro"],
        "barrio": r["barrio"],
        "direccion": r["direccion"],
        "zona": r["zona"],
        "lat": to_num(r["latitud"]),
        "lng": to_num(r["longitud"]),
        "observaciones": r["observaciones"],
    })


def main():
    with open(CSV_IN, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))

    consultas = [consulta_row(r) for r in rows if r["tipo_registro"] == "consulta"]
    saps = [saps_row(r) for r in rows if r["tipo_registro"] == "centro"
            and r["tipo_centro"] == "SAPS"]
    caps = [simple_row(r) for r in rows if r["tipo_registro"] == "centro"
            and r["tipo_centro"] == "CAPS"]
    hospitales = [simple_row(r) for r in rows if r["tipo_registro"] == "centro"
                  and r["tipo_centro"] == "Hospital público"]
    centros = [simple_row(r) for r in rows if r["tipo_registro"] == "centro"
               and r["tipo_centro"] == "Centro privado"]

    zonas_fc = json.loads(ZONAS_IN.read_text(encoding="utf-8"))
    zonas = [{"zona": f["properties"]["zona"], "geometry": f["geometry"]}
             for f in zonas_fc["features"]]

    meta = {
        "titulo": "Salud materna · Ciudad de Corrientes",
        "generado": date.today().isoformat(),
        "generador": "generacion/generar_data_js.py",
        "base": "generacion/datos_salud_optimizados.csv",
        "totales": {
            "consultas": sum(c["consulta_cantidad"] for c in consultas),
            "filas_consulta": len(consultas),
            "establecimientos": len(saps) + len(caps) + len(hospitales) + len(centros),
            "saps_catalogo": len(saps),
            "saps_con_consultas": sum(1 for s in saps if s.get("consultas")),
        },
    }

    data = {"meta": meta, "consultas": consultas, "saps": saps, "caps": caps,
            "hospitales": hospitales, "centros": centros, "zonas": zonas}

    header = (
        "/* ============================================================" + NL +
        "   data.js — Datos del Dashboard de Salud Materna" + NL +
        "   (Ciudad de Corrientes) — generado por" + NL +
        "   generacion/generar_data_js.py a partir de" + NL +
        "   generacion/datos_salud_optimizados.csv." + NL +
        "   No editar a mano: volver a correr el script." + NL +
        "   ============================================================ */" + NL
    )
    js = header + "window.DASHBOARD_DATA = " + dump(data) + ";" + NL
    JS_OUT.write_text(js, encoding="utf-8", newline="")

    print(f"data.js -> {JS_OUT}")
    print(f"  {len(js.encode('utf-8'))} bytes | lineas: {js.count(chr(10))}")
    print(f"  consultas={len(consultas)} (total {meta['totales']['consultas']}) | "
          f"saps={len(saps)} | caps={len(caps)} | hospitales={len(hospitales)} | "
          f"centros={len(centros)} | zonas={len(zonas)}")


if __name__ == "__main__":
    main()
