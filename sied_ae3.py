#!/usr/bin/env python3
"""Tablero de la AE3 por aula, a partir del informe SQL «Resultados_AE3».

Genera un index.html autocontenido para GitHub Pages: los datos van embebidos,
no hay pedidos de red ni dependencias.

    python3 sied_ae3.py --datos AE3_por_aula.xlsx --salida sitio_ae3

Paleta validada con scripts/validate_palette.js del skill dataviz:
  claro   #2E8B57 / #E09112 / #C8352B  — CVD protan 10.8, visión normal 20.8
  oscuro  #1A8A51 / #A89300 / #A74639  — CVD protan  8.6, visión normal 16.2
Los dos modos tienen un aviso de contraste por debajo de 3:1, que se compensa
con etiquetas directas en cada segmento y la tabla de detalle, como pide la guía.
«Nadie rindió» y «Sin estudiantes» van en gris: no son estados de falla y no
deben competir por un tono cromático.
"""
from __future__ import annotations

import argparse
import json
import re
import unicodedata
from datetime import date
from pathlib import Path

import pandas as pd

# Aulas que no pertenecen al relevamiento: respaldos viejos y plantillas.
FUERA = re.compile(r"respaldo_eabc|aulas base", re.I)

# Módulos que se llaman AE3 pero no son la evaluación: simulacros y prácticas.
# Contarlos infla todo, y fueron la mayor parte del error de sobreconteo.
NO_ES_AE3 = re.compile(r"simulacro|simulaci|prueba|pr[áa]ctica|practica", re.I)

ESTADOS = [
    ("corregida", "Corregida", "ok"),
    ("parcial", "Corregida parcialmente", "warn"),
    ("sin_nota", "Rindieron, pendientes de calificar", "bad"),
    ("nadie", "No se registran estudiantes que rindieron", "gris"),
    ("vacia", "Aula sin estudiantes", "gris2"),
]


# ── Unidad académica, carrera y sede ────────────────────────────────
# El árbol de categorías no está armado igual en las dos ramas.
#   Presencial  /998/<sede>/<facultad>/<carrera>/<año>/<período>
#   Digital     /1000/<facultad>/<carrera>/…   o   /1000/<carrera>/…
# Y los mismos nombres vienen escritos de varias formas: «Contador
# Publico» con y sin tilde y con espacio al final, la facultad de
# Ingeniería con y sin coma, «Lic.» contra «Licenciatura». Normalizar
# no es cosmético: sin esto, una carrera aparece como tres.

UA_OFICIALES = [
    "Facultad de Ciencias Sociales y Humanidades",
    "Facultad de Ciencias Económicas y Ambientales",
    "Facultad de Ciencias Jurídicas y Políticas",
    "Facultad de Ciencias de la Salud y Bienestar",
    "Facultad de Ingeniería, Tecnología y Arquitectura",
    "Facultad de Arte, Diseño y Comunicación",
]

# Carrera → unidad académica, para la rama digital, donde la facultad
# no siempre está en el árbol.
UA_POR_CARRERA = {
    "abogacia": "Facultad de Ciencias Jurídicas y Políticas",
    "escribania": "Facultad de Ciencias Jurídicas y Políticas",
    "criminalistica": "Facultad de Ciencias Jurídicas y Políticas",
    "contador publico": "Facultad de Ciencias Económicas y Ambientales",
    "licenciatura en administracion de empresas": "Facultad de Ciencias Económicas y Ambientales",
    "licenciatura en comercio internacional": "Facultad de Ciencias Económicas y Ambientales",
    "licenciatura en psicologia": "Facultad de Ciencias Sociales y Humanidades",
    "licenciatura en psicopedagogia": "Facultad de Ciencias Sociales y Humanidades",
    "licenciatura en ciencias de la educacion": "Facultad de Ciencias Sociales y Humanidades",
    "profesorado universitario": "Facultad de Ciencias Sociales y Humanidades",
    "licenciatura en publicidad": "Facultad de Ciencias Sociales y Humanidades",
    "licenciatura en acompanamiento terapeutico": "Facultad de Ciencias de la Salud y Bienestar",
    "tecnicatura universitaria en acompanante terapeutico": "Facultad de Ciencias de la Salud y Bienestar",
    "licenciatura en nutricion": "Facultad de Ciencias de la Salud y Bienestar",
    "licenciatura en fonoaudiologia": "Facultad de Ciencias de la Salud y Bienestar",
    "licenciatura en farmacia": "Facultad de Ciencias de la Salud y Bienestar",
    "tecnicatura universitaria en diseno y programacion de videojuegos":
        "Facultad de Ingeniería, Tecnología y Arquitectura",
    "ingenieria en sistemas de informacion": "Facultad de Ingeniería, Tecnología y Arquitectura",
    "licenciatura en sistemas de informacion": "Facultad de Ingeniería, Tecnología y Arquitectura",
    "arquitectura": "Facultad de Ingeniería, Tecnología y Arquitectura",
    "licenciatura en diseno grafico y multimedia": "Facultad de Arte, Diseño y Comunicación",
    "licenciatura en diseno de indumentaria y textil": "Facultad de Arte, Diseño y Comunicación",
    "licenciatura en diseno de muebles e interiores": "Facultad de Arte, Diseño y Comunicación",
}

ALIAS_CARRERA = {
    "tecnicatura universitaria en videojuegos":
        "tecnicatura universitaria en diseno y programacion de videojuegos",
    "criminalistica": "licenciatura en criminalistica",
    "licenciatura en psicopedagogia": "licenciatura en psicopedagogia",
}

_SIN_SENTIDO = re.compile(
    r"\s*\((aulas base|base)\)|\s*-\s*oculto\b|\s*plan\s*\d+/\d+\b|"
    r"\s*\bconeau\b.*$", re.I)


def _clave(t: str) -> str:
    """Nombre comparable: sin tildes, sin abreviaturas, sin ruido."""
    t = str(t or "").strip()
    if t.lower() in ("nan", "none", "-"):
        return ""
    t = _SIN_SENTIDO.sub("", t).strip()
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    t = re.sub(r"\bccc\.?\s*", "", t)
    t = re.sub(r"\blic\.?\s+(en\s+)?", "licenciatura en ", t)
    t = re.sub(r"\btec\.\s*(univ\.?)?\s*(en\s+)?", "tecnicatura universitaria en ", t)
    t = re.sub(r"\bing\.?\s+(en\s+)?", "ingenieria en ", t)
    t = re.sub(r"\bcs\.?\s+", "ciencias ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return ALIAS_CARRERA.get(t, t)


def _titulo(clave: str) -> str:
    """De la clave normalizada al nombre que se muestra."""
    if not clave:
        return "Sin asignar"
    tildes = {
        "abogacia": "Abogacía", "escribania": "Escribanía",
        "criminalistica": "Criminalística", "arquitectura": "Arquitectura",
        "contador publico": "Contador Público",
        "psicologia": "Psicología", "psicopedagogia": "Psicopedagogía",
        "educacion": "Educación", "nutricion": "Nutrición",
        "fonoaudiologia": "Fonoaudiología", "administracion": "Administración",
        "informacion": "Información", "diseno": "Diseño",
        "programacion": "Programación", "terapeutico": "Terapéutico",
        "acompanamiento": "Acompañamiento", "acompanante": "Acompañante",
        "grafico": "Gráfico", "indumentaria": "Indumentaria",
        "tecnicatura": "Tecnicatura", "ingenieria": "Ingeniería",
        "universitaria": "Universitaria", "universitario": "Universitario",
        "licenciatura": "Licenciatura", "profesorado": "Profesorado",
        "publico": "Público", "publica": "Pública", "tecnico": "Técnico",
        "matematica": "Matemática", "practica": "Práctica",
    }
    palabras = []
    for w in clave.split():
        if w in tildes:
            palabras.append(tildes[w])
        elif w in ("en", "de", "y", "la", "el", "del", "e"):
            palabras.append(w)
        else:
            palabras.append(w.capitalize())
    t = " ".join(palabras)
    return t[0].upper() + t[1:] if t else t


def _ua(nombre) -> str:
    """Devuelve la facultad oficial si el texto nombra una."""
    k = _clave(nombre)
    if not k.startswith("facultad"):
        return ""
    for of in UA_OFICIALES:
        if _clave(of) == k:
            return of
    # Ingeniería viene con y sin coma; comparamos sin puntuación.
    sin = lambda x: re.sub(r"[^a-z ]", "", _clave(x))
    for of in UA_OFICIALES:
        if sin(of) == sin(nombre):
            return of
    return str(nombre).strip()


def ubicar(r) -> tuple[str, str, str]:
    """(sede, unidad académica, carrera) para una fila del informe."""
    n2, n3, n4 = (str(r.get(c) or "").strip() for c in ("nivel_2", "nivel_3", "nivel_4"))
    if r["rama"] == "Presencial":
        sede = n2 or "Sin sede"
        ua = _ua(n3) or "Sin asignar"
        carrera = _titulo(_clave(n4)) if n4 else "Sin asignar"
    else:
        sede = "Campus digital"
        ua_directa = _ua(n2)
        bruto = n3 if ua_directa else n2
        carrera = _titulo(_clave(bruto)) if bruto else "Sin asignar"
        ua = ua_directa or UA_POR_CARRERA.get(_clave(bruto), "Sin asignar")
    if not _clave(carrera):
        carrera = "Sin asignar"
    return sede, ua, carrera


def _norm(t) -> str:
    if not isinstance(t, str):
        return ""
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip()


def leer(ruta: Path, conteos: pd.DataFrame | None = None) -> pd.DataFrame:
    d = pd.read_excel(ruta) if ruta.suffix == ".xlsx" else pd.read_csv(ruta)
    d.columns = [str(c).strip() for c in d.columns]
    faltan = {"rama", "categoria_moodle", "curso_id", "aula"} - set(d.columns)
    if faltan:
        raise SystemExit(f"Al informe le faltan columnas: {sorted(faltan)}")

    d = d[~d["modulo_nombre"].fillna("").str.contains(NO_ES_AE3)].copy()
    d["historica"] = (d["categoria_moodle"].fillna("").str.contains(FUERA)
                      | d.get("nivel_3", pd.Series("", index=d.index)).fillna("").str.contains(FUERA)
                      | d.get("nivel_2", pd.Series("", index=d.index)).fillna("").str.contains(FUERA))
    d = d[~d["historica"]].copy()

    if "nivel_2" in d.columns:
        ub = d.apply(ubicar, axis=1, result_type="expand")
        d["sede"], d["unidad"], d["carrera"] = ub[0], ub[1], ub[2]
    else:
        d["sede"] = d["unidad"] = d["carrera"] = "Sin asignar"

    for c in ("matriculados", "intentos", "estudiantes_que_rindieron", "con_nota",
              "aprobados", "desaprobados", "excepciones_usuario", "excepciones_grupo"):
        d[c] = pd.to_numeric(d.get(c), errors="coerce").fillna(0).astype(int)

    # Una fila por aula. Un aula puede tener los cinco casos de la AE3
    # montados; contarlos por módulo haría pasar por maquetación fallida a los
    # casos que ningún estudiante eligió.
    a = (d.groupby(["rama", "sede", "unidad", "carrera", "categoria_moodle",
                     "curso_id", "aula"], dropna=False)
           .agg(modulos=("modulo_id", "size"),
                matriculados=("matriculados", "max"),
                intentos=("intentos", "sum"),
                rindieron=("estudiantes_que_rindieron", "sum"),
                con_nota=("con_nota", "sum"),
                aprobados=("aprobados", "sum"),
                desaprobados=("desaprobados", "sum"),
                exc_usuario=("excepciones_usuario", "sum"),
                exc_grupo=("excepciones_grupo", "sum"),
                nota_maxima=("nota_maxima", "max"),
                nota_aprobacion=("nota_aprobacion", "max"))
           .reset_index())
    # Los conteos por módulo se reemplazan por los de estudiantes distintos
    # apenas estén disponibles: son los únicos que no duplican.
    if conteos is not None:
        for destino, origen in (("rindieron", "estudiantes_distintos"),
                                ("intentos", "intentos_totales"),
                                ("con_nota", "estudiantes_con_nota"),
                                ("aprobados", "estudiantes_aprobados")):
            if origen in conteos.columns:
                real = a["curso_id"].map(conteos[origen])
                a[destino] = real.fillna(a[destino]).astype(int)
        if "modulos_usados" in conteos.columns:
            a["modulos_usados"] = (a["curso_id"].map(conteos["modulos_usados"])
                                     .fillna(0).astype(int))
        if "nota_promedio_sobre_10" in conteos.columns:
            a["nota_promedio"] = a["curso_id"].map(conteos["nota_promedio_sobre_10"])
        # Aprobados y desaprobados tienen que cerrar contra los calificados.
        a["desaprobados"] = (a["con_nota"] - a["aprobados"]).clip(lower=0)

    a["sin_corregir"] = (a["rindieron"] - a["con_nota"]).clip(lower=0)

    def estado(r):
        if r["matriculados"] == 0:
            return "vacia"
        if r["intentos"] == 0:
            return "nadie"
        if r["con_nota"] == 0:
            return "sin_nota"
        if r["sin_corregir"] > 0:
            return "parcial"
        return "corregida"

    a["estado"] = a.apply(estado, axis=1)
    # Dos situaciones distintas, que antes iban juntas bajo «fuera de norma»:
    # un aula puede no tener cargada la nota de aprobación, o estar puntuada
    # sobre una escala que no es la de 10. No son lo mismo y se arreglan
    # distinto, así que se informan por separado.
    a["sin_aprobacion"] = a["nota_aprobacion"].isna()
    a["escala_distinta"] = a["nota_maxima"].notna() & (a["nota_maxima"] != 10)
    return a.sort_values(["rama", "unidad", "carrera", "aula"]).reset_index(drop=True)


def leer_conteos(distintos: Path | None, califs: Path | None) -> pd.DataFrame | None:
    """Conteos de estudiantes DISTINTOS por aula.

    Sin esto hay que sumar los módulos AE3 de cada aula, y un estudiante que
    hizo más de un caso se cuenta una vez por cada uno. En las 71 aulas con
    varios módulos usados eso infla el total; en las 852 restantes da igual.
    """
    partes = []
    for ruta, cols in ((distintos, ["estudiantes_distintos", "intentos_totales",
                                    "modulos_usados"]),
                       (califs, ["estudiantes_con_nota", "estudiantes_aprobados",
                                 "nota_promedio_sobre_10"])):
        if not ruta:
            continue
        d = pd.read_excel(ruta) if ruta.suffix == ".xlsx" else pd.read_csv(ruta)
        d.columns = [str(c).strip() for c in d.columns]
        if "curso_id" not in d.columns:
            raise SystemExit(f"{ruta.name}: falta la columna «curso_id».")
        usar = ["curso_id"] + [c for c in cols if c in d.columns]
        partes.append(d[usar].set_index("curso_id"))
    if not partes:
        return None
    return pd.concat(partes, axis=1)


def leer_evolucion(ruta: Path | None) -> list[dict]:
    """Informe «Evolución AE3»: una fila por día y rama."""
    if not ruta:
        return []
    d = pd.read_excel(ruta) if ruta.suffix == ".xlsx" else pd.read_csv(ruta)
    d.columns = [str(c).strip() for c in d.columns]
    if "dia" not in d.columns:
        raise SystemExit("El informe de evolución no tiene columna «dia».")
    d["dia"] = pd.to_datetime(d["dia"], errors="coerce")
    d = d.dropna(subset=["dia"]).sort_values("dia")
    for c in ("intentos_iniciados", "intentos_terminados", "abandonados", "estudiantes"):
        d[c] = pd.to_numeric(d.get(c), errors="coerce").fillna(0).astype(int)
    return [{"r": r["rama"], "d": r["dia"].strftime("%Y-%m-%d"),
             "i": int(r["intentos_iniciados"]), "t": int(r["intentos_terminados"]),
             "e": int(r["estudiantes"]), "ab": int(r["abandonados"])}
            for _, r in d.iterrows()]


def compactar(a: pd.DataFrame, evo: list[dict] | None = None) -> dict:
    filas = [{
        "r": r["rama"], "sed": r["sede"], "ua": r["unidad"], "car": r["carrera"],
        "cat": r["categoria_moodle"], "id": int(r["curso_id"]),
        "aula": r["aula"], "mod": int(r["modulos"]), "mat": int(r["matriculados"]),
        "rin": int(r["rindieron"]), "cn": int(r["con_nota"]),
        "ap": int(r["aprobados"]), "des": int(r["desaprobados"]),
        "sc": int(r["sin_corregir"]), "eu": int(r["exc_usuario"]),
        "eg": int(r["exc_grupo"]), "est": r["estado"],
        "sap": bool(r["sin_aprobacion"]),
        "esc": bool(r["escala_distinta"]),
        "nmax": (None if pd.isna(r["nota_maxima"]) else float(r["nota_maxima"])),
        "nap": (None if pd.isna(r["nota_aprobacion"]) else float(r["nota_aprobacion"])),
        "prom": (None if pd.isna(r.get("nota_promedio")) else round(float(r["nota_promedio"]), 2)),
    } for _, r in a.iterrows()]
    return {
        "generado": date.today().isoformat(),
        "estados": [{"k": k, "t": t, "c": c} for k, t, c in ESTADOS],
        "aulas": filas,
        "evo": evo or [],
    }


PAGINA = r"""<!doctype html>
<html lang="es"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>__TITULO__</title>
<style>
:root{
  --bg:#F4F6FA; --card:#FFFFFF; --ink:#1A1F2B; --mute:#5F6878; --line:#E3E7EE;
  --pri:#025EE4; --head:#0F5FDB; --track:#EDF1F7;
  --ok:#2E8B57; --warn:#E09112; --bad:#C8352B; --gris:#8894A6; --gris2:#C7CFDA;
  --cal-si:#025EE4; --cal-no:#C8352B;
  --sombra:0 1px 2px rgba(16,24,40,.05), 0 1px 3px rgba(16,24,40,.04);
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#10151D; --card:#171E29; --ink:#E7ECF2; --mute:#9AA7B8; --line:#2A3544;
  --pri:#4D8DF5; --head:#7FB0FF; --track:#222C3A;
  --ok:#1A8A51; --warn:#A89300; --bad:#A74639; --gris:#7C8798; --gris2:#495566;
  --cal-si:#4D8DF5; --cal-no:#C8564A;
  --sombra:0 1px 3px rgba(0,0,0,.4);
}}
:root[data-theme="dark"]{
  --bg:#10151D; --card:#171E29; --ink:#E7ECF2; --mute:#9AA7B8; --line:#2A3544;
  --pri:#4D8DF5; --head:#7FB0FF; --track:#222C3A;
  --ok:#1A8A51; --warn:#A89300; --bad:#A74639; --gris:#7C8798; --gris2:#495566;
  --cal-si:#4D8DF5; --cal-no:#C8564A;
  --sombra:0 1px 3px rgba(0,0,0,.4);
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font:400 15px/1.5 "Segoe UI",Roboto,system-ui,-apple-system,Arial,sans-serif;
  padding:26px 18px 70px;-webkit-text-size-adjust:100%}
.wrap{max-width:1180px;margin:0 auto}
h1{margin:0 0 6px;font-size:27px;font-weight:600;color:var(--head);line-height:1.2}
#sub{color:var(--mute);font-size:14px;margin:0 0 20px}
.barra{display:flex;align-items:flex-start;gap:16px;flex-wrap:wrap}
.barra h1{flex:1 1 320px}
button{font:inherit;cursor:pointer;border:1px solid var(--line);border-radius:8px;
  padding:8px 14px;background:var(--card);color:var(--ink);font-size:13.5px}
button:hover{border-color:var(--pri);color:var(--pri)}

/* filtros en una sola fila, arriba de los gráficos */
#filtros{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));
  gap:12px;margin:0 0 8px;align-items:end}
.f label{display:block;font-size:11px;letter-spacing:.07em;text-transform:uppercase;
  color:var(--mute);font-weight:700;margin-bottom:5px}
.f select{width:100%;font:inherit;font-size:14px;padding:8px 10px;border-radius:8px;
  border:1px solid var(--line);background:var(--card);color:var(--ink);cursor:pointer}
.f select:focus{outline:2px solid var(--pri);outline-offset:1px}
#cuenta{color:var(--mute);font-size:13.5px;margin:0 0 18px}

/* KPIs */
#kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(330px,1fr));
  gap:13px;margin-bottom:26px}
.kpi{background:var(--card);border-radius:10px;padding:17px 19px;box-shadow:var(--sombra);
  border-left:4px solid var(--pri)}
.kpi.k-ok{border-left-color:var(--ok)} .kpi.k-bad{border-left-color:var(--bad)}
.kpi.k-warn{border-left-color:var(--warn)}
.kpi b{display:block;font-size:31px;font-weight:600;line-height:1.1;
  font-variant-numeric:tabular-nums}
.kpi span{display:block;font-size:13.5px;margin-top:5px}
.kpi small{display:block;color:var(--mute);font-size:12.5px;margin-top:5px;line-height:1.45}

h2{font-size:17px;font-weight:600;color:var(--head);margin:0 0 4px;
  padding-left:10px;border-left:3px solid var(--pri)}
.secsub{color:var(--mute);font-size:13.5px;margin:0 0 14px;padding-left:13px}
.box{background:var(--card);border-radius:10px;padding:21px 23px;box-shadow:var(--sombra);
  margin-bottom:26px}

/* leyenda: siempre presente, la identidad nunca depende solo del color */
.leyenda{display:flex;gap:17px;flex-wrap:wrap;margin-bottom:16px;font-size:13px}
.leyenda span{display:flex;align-items:center;gap:7px;color:var(--mute)}
.leyenda i{width:11px;height:11px;border-radius:3px;display:inline-block;flex:none}

/* barra segmentada: 2px de hueco entre segmentos, extremos redondeados */
.seg-fila{margin-bottom:19px}
.seg-fila:last-child{margin-bottom:0}
.seg-cab{display:flex;justify-content:space-between;gap:12px;margin-bottom:7px;
  font-size:14px;align-items:baseline}
.seg-cab b{font-weight:600}
.seg-cab em{font-style:normal;color:var(--mute);font-size:13px;
  font-variant-numeric:tabular-nums}
.seg{display:flex;height:30px;border-radius:5px;overflow:hidden;background:var(--track);
  gap:2px}
.seg div{display:flex;align-items:center;justify-content:center;min-width:2px;
  transition:width .5s cubic-bezier(.4,0,.2,1);position:relative}
.seg div:first-child{border-radius:5px 0 0 5px}
.seg div:last-child{border-radius:0 5px 5px 0}
.seg div span{font-size:11.5px;font-weight:600;color:#fff;padding:0 5px;
  white-space:nowrap;font-variant-numeric:tabular-nums;
  text-shadow:0 1px 2px rgba(0,0,0,.35)}

/* barras horizontales de magnitud */
.mag{margin-bottom:15px}
.mag:last-child{margin-bottom:0}
.mag-cab{display:flex;justify-content:space-between;gap:12px;font-size:13.5px;
  margin-bottom:5px}
.mag-cab em{font-style:normal;color:var(--mute);font-variant-numeric:tabular-nums}
.pista{height:11px;background:var(--track);border-radius:4px;overflow:hidden}
.pista i{display:block;height:100%;border-radius:4px;background:var(--pri);
  transition:width .5s cubic-bezier(.4,0,.2,1)}
.pista i.rojo{background:var(--bad)} .pista i.ambar{background:var(--warn)}

/* tabla */
.cab-tabla{display:flex;justify-content:space-between;align-items:center;gap:12px;
  margin-bottom:12px;flex-wrap:wrap}
.cab-tabla h3{margin:0;font-size:15.5px;font-weight:600}
.scroll{overflow-x:auto;-webkit-overflow-scrolling:touch}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:8px 10px;text-align:right;border-bottom:1px solid var(--line);
  white-space:nowrap}
th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){text-align:left;
  white-space:normal;min-width:150px}
th{color:var(--mute);font-weight:600;font-size:11.5px;letter-spacing:.05em;
  text-transform:uppercase;cursor:pointer;user-select:none;position:sticky;top:0;
  background:var(--card)}
th:hover{color:var(--pri)}
td{font-variant-numeric:tabular-nums}
tbody tr:hover{background:var(--track)}
.pastilla{display:inline-block;padding:2px 9px;border-radius:99px;font-size:11.5px;
  font-weight:600;color:#fff;white-space:nowrap}
.p-corregida{background:var(--ok)} .p-parcial{background:var(--warn)}
.p-sin_nota{background:var(--bad)} .p-nadie{background:var(--gris)}
.p-vacia{background:var(--gris2);color:var(--ink)}
.aviso{display:inline-block;margin-left:6px;font-size:11px;color:var(--bad);
  font-weight:600}

/* evolución: paneles pequeños, uno por rama */
.panel{margin-bottom:26px}
.panel:last-child{margin-bottom:0}
.panel-cab{display:flex;justify-content:space-between;align-items:baseline;gap:12px;
  margin-bottom:9px;flex-wrap:wrap}
.panel-cab b{font-size:14.5px;font-weight:600}
.panel-cab em{font-style:normal;color:var(--mute);font-size:13px;
  font-variant-numeric:tabular-nums}
.panel svg{display:block;width:100%;height:150px;overflow:visible}
.panel .eje{stroke:var(--line);stroke-width:1}
.panel .rot{fill:var(--mute);font-size:11px}
.panel .area{fill:var(--pri);opacity:.13}
.panel .linea{fill:none;stroke:var(--pri);stroke-width:2;
  stroke-linejoin:round;stroke-linecap:round}
.panel .cruz{stroke:var(--mute);stroke-width:1;stroke-dasharray:3 3;opacity:0}
.panel .punto{fill:var(--pri);stroke:var(--card);stroke-width:2;r:0;
  transition:r .1s}
.panel:hover .cruz{opacity:.6}

/* tooltip */
#tip{position:fixed;pointer-events:none;opacity:0;transition:opacity .12s;
  background:var(--ink);color:var(--bg);padding:7px 11px;border-radius:7px;
  font-size:12.5px;line-height:1.5;z-index:60;max-width:290px;
  box-shadow:0 4px 14px rgba(0,0,0,.2)}
#tip.on{opacity:1}
#tip b{color:var(--bg)}
footer{color:var(--mute);font-size:12.5px;text-align:center;margin-top:26px;line-height:1.7}
@media(max-width:600px){
  body{padding:18px 13px 50px} h1{font-size:22px} .kpi b{font-size:26px}
  .box{padding:16px 15px}
}
</style>
</head><body>
<div class="wrap">

<div class="barra">
  <h1>__TITULO__</h1>
  <button id="tema" type="button">Modo oscuro</button>
</div>
<p id="sub"></p>

<div id="filtros">
  <div class="f"><label for="fRama">Rama</label><select id="fRama"></select></div>
  <div class="f"><label for="fUA">Unidad académica</label><select id="fUA"></select></div>
  <div class="f"><label for="fCar">Carrera</label><select id="fCar"></select></div>
  <div class="f"><label for="fSede">Sede</label><select id="fSede"></select></div>
  <div class="f"><label for="fEst">Estado</label><select id="fEst"></select></div>
  <div class="f"><label for="fNorma">Configuración</label><select id="fNorma"></select></div>
  <div class="f"><button id="limpiar" type="button">Limpiar filtros</button></div>
</div>
<p id="cuenta"></p>

<div id="kpis"></div>

<h2>Estado de las aulas</h2>
<p class="secsub">Cada aula se clasifica por lo que falta hacer en ella. La AE3 usa
   preguntas de ensayo, que Moodle no corrige solo: por eso «rindieron, pendientes de
   calificar» mide trabajo pendiente, no incumplimiento.</p>
<div class="box">
  <div class="leyenda" id="leyenda"></div>
  <div id="segmentos"></div>
</div>

<h2>Calificación de la AE3</h2>
<p class="secsub">Por carrera: cuántos rindieron, y de esos cuántos ya tienen
   calificación cargada y cuántos siguen esperándola.</p>
<div class="box">
  <div class="leyenda">
    <span><i style="background:var(--cal-si)"></i>Con calificación</span>
    <span><i style="background:var(--cal-no)"></i>Sin calificación</span>
  </div>
  <div id="porCat"></div>
</div>

<section id="secEvo" hidden>
<h2>Evolución de los intentos</h2>
<p class="secsub">Cuándo se desarrolló la instancia. Cada rama tiene su propia escala
   —presencial mueve diez veces más volumen—, así que van en paneles separados: un
   solo eje compartido aplastaría la curva de digital contra el piso.</p>
<div class="box" id="evo"></div>
</section>

<h2>Excepciones de usuario</h2>
<p class="secsub">Aulas con más excepciones individuales cargadas. Cuando superan a
   buena parte del curso, lo que falló fue la configuración del cuestionario y se
   reparó a mano, de a uno.</p>
<div class="box"><div id="exc"></div></div>

<div class="box">
  <div class="cab-tabla">
    <h3>Detalle por aula</h3>
    <button id="bajar" type="button">Descargar CSV</button>
  </div>
  <div class="scroll"><table id="tabla"></table></div>
</div>

<footer><p id="pie"></p></footer>
</div>
<div id="tip" role="status" aria-live="polite"></div>

<script>
const D = __DATOS__;
const $ = i => document.getElementById(i);
const F = {rama:"", ua:"", car:"", sed:"", est:"", norma:""};
const COLOR = {}; D.estados.forEach(e => COLOR[e.k] = e.c);
const TITULO = {}; D.estados.forEach(e => TITULO[e.k] = e.t);

const n0 = v => Math.round(v).toLocaleString("es-AR");
const pc = (a,b) => b > 0 ? Math.round(100*a/b) + "%" : "—";
const esc = s => String(s==null?"":s).replace(/[&<>"]/g,
  c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));

const filtra = () => D.aulas.filter(a =>
  (!F.rama  || a.r === F.rama) &&
  (!F.ua    || a.ua === F.ua) &&
  (!F.car   || a.car === F.car) &&
  (!F.sed   || a.sed === F.sed) &&
  (!F.est   || a.est === F.est) &&
  (!F.norma
     || (F.norma === "sap" && a.sap)
     || (F.norma === "esc" && a.esc)
     || (F.norma === "ok"  && !a.sap && !a.esc)));

const suma = (l,k) => l.reduce((s,a) => s + (a[k]||0), 0);

/* ── tooltip ───────────────────────────────────────────────── */
const tip = $("tip");
function mostrarTip(e, html){
  tip.innerHTML = html; tip.classList.add("on");
  const r = tip.getBoundingClientRect();
  let x = e.clientX + 14, y = e.clientY + 16;
  if (x + r.width > innerWidth - 10) x = e.clientX - r.width - 14;
  if (y + r.height > innerHeight - 10) y = e.clientY - r.height - 16;
  tip.style.left = x + "px"; tip.style.top = y + "px";
}
const ocultarTip = () => tip.classList.remove("on");
function conTip(el, html){
  el.addEventListener("mousemove", e => mostrarTip(e, html));
  el.addEventListener("mouseleave", ocultarTip);
}

/* ── filtros ───────────────────────────────────────────────── */
function opciones(sel, valores, etiqueta, todas){
  sel.innerHTML = `<option value="">${todas} (${valores.length})</option>` +
    valores.map(v => `<option value="${esc(v)}">${esc(etiqueta(v))}</option>`).join("");
}
function armarFiltros(){
  opciones($("fRama"), [...new Set(D.aulas.map(a=>a.r))].sort(), v=>v, "Todas");
  opciones($("fUA"),   [...new Set(D.aulas.map(a=>a.ua))].sort(),  v=>v, "Todas");
  opciones($("fCar"),  [...new Set(D.aulas.map(a=>a.car))].sort(), v=>v, "Todas");
  opciones($("fSede"), [...new Set(D.aulas.map(a=>a.sed))].sort(), v=>v, "Todas");
  $("fEst").innerHTML = `<option value="">Todos (${D.estados.length})</option>` +
    D.estados.map(e => `<option value="${e.k}">${esc(e.t)}</option>`).join("");
  $("fNorma").innerHTML = `<option value="">Todas</option>` +
    `<option value="ok">Escala de 10 · aprueba con 6</option>` +
    `<option value="sap">Sin nota de aprobación configurada</option>` +
    `<option value="esc">Con escala distinta de 10</option>`;
}

/* ── KPIs ──────────────────────────────────────────────────── */
function kpis(l){
  const rin = suma(l,"rin"), cn = suma(l,"cn"), sc = suma(l,"sc");
  const ap = suma(l,"ap"), des = suma(l,"des"), eu = suma(l,"eu");
  const sap = l.filter(a=>a.sap).length, esc = l.filter(a=>a.esc).length;
  $("kpis").innerHTML =
    `<div class="kpi"><b>${n0(l.length)}</b><span>Aulas con la AE3 montada</span>
       <small>${n0(suma(l,"mod"))} módulos en total · un aula puede tener los cinco casos</small></div>
     <div class="kpi k-ok"><b>${n0(rin)}</b><span>Estudiantes que rindieron</span>
       <small>sobre ${n0(suma(l,"mat"))} matriculados</small></div>
     <div class="kpi k-bad"><b>${n0(sc)}</b><span>Esperando corrección · ${pc(sc,rin)}</span>
       <small>rindieron y todavía no tienen nota cargada</small></div>
     <div class="kpi k-warn"><b>${pc(ap, ap+des)}</b><span>Aprobación sobre lo corregido</span>
       <small>${n0(ap)} aprobados y ${n0(des)} desaprobados de ${n0(cn)} con nota</small></div>
     <div class="kpi"><b>${n0(eu)}</b><span>Excepciones de usuario</span>
       <small>en ${n0(l.filter(a=>a.eu>0).length)} aulas</small></div>
     <div class="kpi k-warn"><b>${n0(sap)}</b><span>Sin nota de aprobación configurada</span>
       <small>el cuestionario no tiene cargado desde qué nota se aprueba${
         esc ? ` · ${n0(esc)} además usan una escala distinta de 10` : ""}</small></div>`;
}

/* ── barra segmentada por rama ─────────────────────────────── */
function segmentos(l){
  $("leyenda").innerHTML = D.estados.map(e =>
    `<span><i style="background:var(--${e.c})"></i>${esc(e.t)}</span>`).join("");
  const ramas = [...new Set(l.map(a=>a.r))].sort();
  $("segmentos").innerHTML = ramas.map(r => {
    const sub = l.filter(a=>a.r===r), tot = sub.length;
    const trozos = D.estados.map(e => {
      const n = sub.filter(a=>a.est===e.k).length;
      if (!n) return "";
      const w = n/tot*100;
      return `<div data-est="${e.k}" data-rama="${esc(r)}" data-n="${n}"
                style="width:${w}%;background:var(--${e.c})">${
                w > 7 ? `<span>${n0(n)}</span>` : ""}</div>`;
    }).join("");
    return `<div class="seg-fila">
        <div class="seg-cab"><b>${esc(r)}</b><em>${n0(tot)} aulas</em></div>
        <div class="seg">${trozos}</div></div>`;
  }).join("");
  $("segmentos").querySelectorAll(".seg div").forEach(d => {
    const e = d.dataset.est, n = +d.dataset.n;
    const tot = l.filter(a=>a.r===d.dataset.rama).length;
    conTip(d, `<b>${esc(TITULO[e])}</b><br>${n0(n)} de ${n0(tot)} aulas · ${pc(n,tot)}`);
  });
}

/* ── magnitud: pendientes por categoría ────────────────────── */
function barras(cont, filas, clase, unidad){
  if (!filas.length){ cont.innerHTML = `<p style="color:var(--mute);margin:0">Sin datos para este filtro.</p>`; return; }
  const max = Math.max(...filas.map(f=>f.v)) || 1;
  cont.innerHTML = filas.map(f => `
    <div class="mag" data-tip="${esc(f.tip||"")}">
      <div class="mag-cab"><span>${esc(f.t)}</span><em>${n0(f.v)} ${unidad}</em></div>
      <div class="pista"><i class="${clase}" style="width:${f.v/max*100}%"></i></div>
    </div>`).join("");
  cont.querySelectorAll(".mag").forEach(el => {
    if (el.dataset.tip) conTip(el, el.dataset.tip);
  });
}

function porCategoria(l){
  const m = {};
  l.forEach(a => {
    m[a.car] = m[a.car] || {rin:0, cn:0, sc:0, aulas:0};
    m[a.car].rin += a.rin; m[a.car].cn += a.cn;
    m[a.car].sc += a.sc; m[a.car].aulas++;
  });
  const filas = Object.entries(m).filter(([,v]) => v.rin > 0)
    .sort((a,b) => b[1].rin - a[1].rin).slice(0, 16);
  if (!filas.length){
    $("porCat").innerHTML = `<p style="color:var(--mute);margin:0">Sin estudiantes que hayan rendido para este filtro.</p>`;
    return;
  }
  const tope = Math.max(...filas.map(([,v]) => v.rin));
  $("porCat").innerHTML = filas.map(([k,v]) => {
    // el ancho total es proporcional a los que rindieron, para que no
    // parezca lo mismo una carrera de 3.000 y una de 30
    const ancho = v.rin/tope*100;
    const pSi = v.rin ? v.cn/v.rin*100 : 0;
    const trozo = (val, pct, color, clase) => val
      ? `<div class="${clase}" style="width:${pct}%;background:var(--${color})">${
          pct > 9 ? `<span>${n0(val)}</span>` : ""}</div>` : "";
    return `<div class="mag" data-car="${esc(k)}">
      <div class="mag-cab"><span>${esc(k)}</span>
        <em>${n0(v.rin)} rindieron · ${pc(v.cn, v.rin)} con calificación</em></div>
      <div class="seg" style="width:${ancho}%;height:19px">
        ${trozo(v.cn, pSi, "cal-si", "s1")}${trozo(v.sc, 100-pSi, "cal-no", "s2")}
      </div></div>`;
  }).join("");
  $("porCat").querySelectorAll(".mag").forEach(el => {
    const v = m[el.dataset.car];
    conTip(el, `<b>${esc(el.dataset.car)}</b><br>` +
      `${n0(v.rin)} rindieron en ${n0(v.aulas)} aulas<br>` +
      `${n0(v.cn)} con calificación · ${pc(v.cn, v.rin)}<br>` +
      `${n0(v.sc)} sin calificación · ${pc(v.sc, v.rin)}`);
  });
}

function excepciones(l){
  const filas = l.filter(a=>a.eu>0).sort((a,b)=>b.eu-a.eu).slice(0,12)
    .map(a => ({t:a.aula, v:a.eu,
      tip:`<b>${esc(a.aula)}</b><br>${n0(a.eu)} excepciones de usuario<br>` +
          `${n0(a.mat)} matriculados · ${n0(a.rin)} rindieron`}));
  barras($("exc"), filas, "ambar", "excepciones");
}

/* ── evolución: un panel por rama, escalas propias ───────── */
function evolucion(){
  if (!D.evo || !D.evo.length) return;
  $("secEvo").hidden = false;
  const ramas = [...new Set(D.evo.map(e=>e.r))];
  const dias = [...new Set(D.evo.map(e=>e.d))].sort();
  const fcorta = iso => { const [,m,d] = iso.split("-"); return `${+d}/${+m}`; };
  const W = 1000, H = 150, P = {t:10, r:10, b:22, l:46};

  $("evo").innerHTML = ramas.map(r => {
    const serie = dias.map(d => {
      const f = D.evo.find(e => e.r === r && e.d === d);
      return {d, i: f ? f.i : 0, t: f ? f.t : 0, e: f ? f.e : 0};
    });
    const max = Math.max(...serie.map(p=>p.i)) || 1;
    const total = serie.reduce((s,p)=>s+p.i, 0);
    const x = k => P.l + k*(W-P.l-P.r)/Math.max(1,serie.length-1);
    const y = v => P.t + (1 - v/max)*(H-P.t-P.b);
    const pts = serie.map((p,k)=>`${x(k).toFixed(1)},${y(p.i).toFixed(1)}`).join(" ");
    const area = `${P.l},${y(0)} ${pts} ${x(serie.length-1)},${y(0)}`;
    const pico = serie.reduce((a,b)=> b.i>a.i ? b : a, serie[0]);
    // una marca cada ~6 días, para que las fechas no choquen
    const paso = Math.max(1, Math.round(serie.length/7));
    const rotulos = serie.map((p,k)=> k%paso===0 || k===serie.length-1
      ? `<text class="rot" x="${x(k)}" y="${H-6}" text-anchor="middle">${fcorta(p.d)}</text>` : "").join("");
    return `<div class="panel" data-rama="${esc(r)}">
      <div class="panel-cab"><b>${esc(r)}</b>
        <em>${n0(total)} intentos · pico el ${fcorta(pico.d)} con ${n0(pico.i)}</em></div>
      <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" role="img"
           aria-label="Intentos por día en ${esc(r)}">
        <line class="eje" x1="${P.l}" y1="${y(0)}" x2="${W-P.r}" y2="${y(0)}"></line>
        <line class="eje" x1="${P.l}" y1="${P.t}" x2="${P.l}" y2="${y(0)}"></line>
        <text class="rot" x="${P.l-8}" y="${P.t+4}" text-anchor="end">${n0(max)}</text>
        <text class="rot" x="${P.l-8}" y="${y(0)}" text-anchor="end">0</text>
        <polygon class="area" points="${area}"></polygon>
        <polyline class="linea" points="${pts}"></polyline>
        <line class="cruz" x1="0" y1="${P.t}" x2="0" y2="${y(0)}"></line>
        <circle class="punto" cx="0" cy="0"></circle>
        ${rotulos}
      </svg></div>`;
  }).join("");

  // crosshair + tooltip sobre el punto más cercano
  $("evo").querySelectorAll(".panel").forEach(pan => {
    const svg = pan.querySelector("svg");
    const cruz = pan.querySelector(".cruz"), punto = pan.querySelector(".punto");
    const r = pan.dataset.rama;
    const serie = dias.map(d => {
      const f = D.evo.find(e => e.r === r && e.d === d);
      return {d, i: f?f.i:0, t: f?f.t:0, e: f?f.e:0};
    });
    const max = Math.max(...serie.map(p=>p.i)) || 1;
    svg.addEventListener("mousemove", ev => {
      const caja = svg.getBoundingClientRect();
      const u = (ev.clientX - caja.left) / caja.width * W;
      const k = Math.max(0, Math.min(serie.length-1,
        Math.round((u - P.l) / ((W-P.l-P.r)/Math.max(1,serie.length-1)))));
      const px = P.l + k*(W-P.l-P.r)/Math.max(1,serie.length-1);
      const py = P.t + (1 - serie[k].i/max)*(H-P.t-P.b);
      cruz.setAttribute("x1", px); cruz.setAttribute("x2", px);
      punto.setAttribute("cx", px); punto.setAttribute("cy", py);
      punto.style.r = "5";
      const p = serie[k];
      mostrarTip(ev, `<b>${fcorta(p.d)}</b> · ${esc(r)}<br>` +
        `${n0(p.i)} intentos iniciados<br>${n0(p.t)} terminados · ${n0(p.e)} estudiantes`);
    });
    svg.addEventListener("mouseleave", () => { punto.style.r = "0"; ocultarTip(); });
  });
}

/* ── tabla ─────────────────────────────────────────────────── */
const COLS = [
  ["aula","Aula",a=>esc(a.aula)],
  ["ua","Unidad académica",a=>esc(a.ua.replace(/^Facultad de /,""))],
  ["car","Carrera",a=>esc(a.car)],
  ["mat","Matric.",a=>n0(a.mat)],
  ["rin","Rindieron",a=>n0(a.rin)],
  ["cn","Con nota",a=>n0(a.cn)],
  ["sc","Sin corregir",a=>n0(a.sc)],
  ["ap","Aprob.",a=>n0(a.ap)],
  ["des","Desaprob.",a=>n0(a.des)],
  ["eu","Excep.",a=>n0(a.eu)],
  ["est","Estado",a=>`<span class="pastilla p-${a.est}">${esc(TITULO[a.est])}</span>` +
      (a.sap ? `<span class="aviso" title="El cuestionario no tiene cargada la nota de aprobación">sin nota de aprobación</span>` : "") +
      (a.esc ? `<span class="aviso" title="El cuestionario no está puntuado sobre 10">escala ${a.nmax ?? "?"}</span>` : "")],
];
let orden = {col:"sc", desc:true};
function tabla(l){
  const d = [...l].sort((x,y) => {
    const a = x[orden.col], b = y[orden.col];
    const r = (typeof a === "number" && typeof b === "number")
      ? a - b : String(a).localeCompare(String(b), "es");
    return orden.desc ? -r : r;
  }).slice(0, 400);
  $("tabla").innerHTML =
    `<thead><tr>${COLS.map(([k,t]) =>
      `<th data-col="${k}">${t}${orden.col===k ? (orden.desc?" ▼":" ▲") : ""}</th>`).join("")}</tr></thead>` +
    `<tbody>${d.map(a => `<tr>${COLS.map(([,,f]) => `<td>${f(a)}</td>`).join("")}</tr>`).join("")}</tbody>`;
  $("tabla").querySelectorAll("th").forEach(th => th.onclick = () => {
    const c = th.dataset.col;
    orden = {col:c, desc: orden.col === c ? !orden.desc : true};
    tabla(filtra());
  });
}

/* ── pintar ────────────────────────────────────────────────── */
function pintar(){
  const l = filtra();
  const partes = [];
  if (F.rama) partes.push(F.rama);
  if (F.ua) partes.push(F.ua);
  if (F.car) partes.push(F.car);
  if (F.sed) partes.push(F.sed);
  if (F.est) partes.push(TITULO[F.est]);
  if (F.norma) partes.push({sap:"sin nota de aprobación configurada",
                            esc:"con escala distinta de 10",
                            ok:"escala de 10 y aprueba con 6"}[F.norma]);
  $("cuenta").textContent =
    `Mostrando ${n0(l.length)} aulas de ${n0(D.aulas.length)}` +
    (partes.length ? ` · ${partes.join(" · ")}` : " · sin filtros");
  kpis(l); segmentos(l); porCategoria(l); excepciones(l); tabla(l);
}

["fRama","fUA","fCar","fSede","fEst","fNorma"].forEach((id,i) => {
  $(id).onchange = e => {
    F[["rama","ua","car","sed","est","norma"][i]] = e.target.value;
    pintar();
  };
});
$("limpiar").onclick = () => {
  F.rama = F.ua = F.car = F.sed = F.est = F.norma = "";
  ["fRama","fUA","fCar","fSede","fEst","fNorma"].forEach(id => $(id).value = "");
  pintar();
};

$("bajar").onclick = () => {
  const l = filtra();
  const cab = ["rama","sede","unidad_academica","carrera","categoria","curso_id","aula","modulos_ae3","matriculados",
               "rindieron","con_nota","sin_corregir","aprobados","desaprobados",
               "excepciones_usuario","excepciones_grupo","nota_maxima",
               "nota_aprobacion","estado","sin_nota_de_aprobacion","escala_distinta_de_10"];
  const filas = l.map(a => [a.r,a.sed,a.ua,a.car,a.cat,a.id,a.aula,a.mod,a.mat,a.rin,a.cn,a.sc,
    a.ap,a.des,a.eu,a.eg,a.nmax ?? "",a.nap ?? "",TITULO[a.est],
    a.sap?"sí":"no",a.esc?"sí":"no"]);
  const csv = "﻿" + [cab,...filas].map(f =>
    f.map(v => `"${String(v).replace(/"/g,'""')}"`).join(",")).join("\r\n");
  const u = URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8"}));
  const a = document.createElement("a");
  a.href = u; a.download = "AE3_por_aula.csv";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(u), 1500);
};

/* ── tema ──────────────────────────────────────────────────── */
const raiz = document.documentElement;
const esOscuro = () => raiz.dataset.theme === "dark" ||
  (!raiz.dataset.theme && matchMedia("(prefers-color-scheme:dark)").matches);
function tema(){ $("tema").textContent = esOscuro() ? "Modo claro" : "Modo oscuro"; }
$("tema").onclick = () => {
  raiz.dataset.theme = esOscuro() ? "light" : "dark";
  tema(); pintar();
};

$("sub").textContent = "Dónde está montada la Actividad de Evaluación 3 y en qué " +
  "estado quedó cada aula, en las ramas presencial y digital del campus.";
$("pie").textContent = `Datos procesados el ${D.generado.split("-").reverse().join("/")}. ` +
  `Departamento Pedagógico SIED · Universidad de la Cuenca del Plata.`;

armarFiltros(); tema(); pintar(); evolucion();
</script>
</body></html>
"""


def main():
    ap = argparse.ArgumentParser(description="Tablero de la AE3 por aula")
    ap.add_argument("--datos", required=True, type=Path)
    ap.add_argument("--salida", default="sitio_ae3", type=Path)
    ap.add_argument("--distintos", type=Path, default=None,
                    help="Informe «Estudiantes por aula AE3»")
    ap.add_argument("--calificaciones", type=Path, default=None,
                    help="Informe «Calificaciones por aula AE3»")
    ap.add_argument("--evolucion", type=Path, default=None,
                    help="Informe «Evolución AE3», opcional")
    ap.add_argument("--titulo", default="Actividad de Evaluación 3 · relevamiento por aula")
    args = ap.parse_args()

    a = leer(args.datos, leer_conteos(args.distintos, args.calificaciones))
    datos = compactar(a, leer_evolucion(args.evolucion))
    html = (PAGINA.replace("__TITULO__", args.titulo)
                  .replace("__DATOS__", json.dumps(datos, ensure_ascii=False,
                                                   separators=(",", ":"))))
    args.salida.mkdir(parents=True, exist_ok=True)
    destino = args.salida / "index.html"
    destino.write_text(html, encoding="utf-8")

    por = a["estado"].value_counts()
    print(f"  {destino}  ({len(html)//1024} KB · {len(a)} aulas)")
    for k, t, _ in ESTADOS:
        print(f"    {t:34} {int(por.get(k, 0)):5}")
    print(f"    {'esperando corrección':34} {int(a['sin_corregir'].sum()):5} estudiantes")


if __name__ == "__main__":
    main()
