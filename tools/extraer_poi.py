#!/usr/bin/env python3
"""Extrae OEI -> AEI -> AOI de cada facultad del PDF "POI - GENERAL" (Anexo B-5)
y genera apps-script/Datos.gs.

Uso:  python3 tools/extraer_poi.py "POI - GENERAL.pdf" apps-script/Datos.gs
Requiere: pip install pymupdf
"""
import collections
import json
import re
import sys

import pymupdf as fitz

# sigla -> (formulario, centro de costo del PDF).  Igual que el Anexo 1 del Sheet.
FACULTADES = [
    ("FM", "F01", "01.30"), ("FDCP", "F02", "01.22"), ("FLCH", "F03", "01.29"),
    ("FFB", "F04", "01.24"), ("FO", "F05", "01.32"), ("FE", "F06", "01.23"),
    ("FQIQ", "F07", "01.34"), ("FMV", "F08", "01.31"), ("FCA", "F09", "01.15"),
    ("FCB", "F10", "01.16"), ("FCC", "F11", "01.17"), ("FCE", "F12", "01.18"),
    ("FCF", "F13", "01.19"), ("FCM", "F14", "01.20"), ("FCCSS", "F15", "01.21"),
    ("FIGMMG", "F16", "01.27"), ("FII", "F20", "01.28"), ("FPSIC", None, "01.33"),
    ("FIEE", None, "01.26"), ("FISI", None, "01.25"),
]


def lineas(words, ytol=2.0):
    out = []
    for w in sorted(words, key=lambda w: (w[1], w[0])):
        if out and abs(out[-1][0] - w[1]) <= ytol:
            out[-1][1].append(w)
        else:
            out.append([w[1], [w]])
    for _, l in out:
        l.sort(key=lambda w: w[0])
    return out


def extraer(pdf):
    doc = fitz.open(pdf)
    cc = oei = aei = None
    filas = []
    for pi, page in enumerate(doc):
        if pi < 5:  # portada, resolución, etc.
            continue
        words = [w for w in page.get_text("words") if not (w[1] < 22 or w[1] > 818)]
        events, codes = [], []
        for y, l in lineas(words):
            txt = " ".join(w[4] for w in l)
            if l[0][4] == "Centro" and len(l) > 2 and l[1][4].startswith("de"):
                events.append((y, "CC", txt))
            elif re.fullmatch(r"OEI\.\d+", l[0][4]):
                events.append((y, "OEI", txt))
            elif re.fullmatch(r"AEI\.\d+\.\d+", l[0][4]):
                events.append((y, "AEI", txt))
        for w in words:
            if re.fullmatch(r"AOI\d{11}", w[4]):
                codes.append(w)
                events.append((w[1], "AOI", w))
        events.sort(key=lambda e: e[0])
        cy = {id(c): (c[1] + c[3]) / 2 for c in codes}
        # El ancho de las columnas cambia de una tabla a otra: se acota la
        # descripción de cada fila por la columna "Ubigeo" de esa misma fila.
        bound = {}
        for c in codes:
            cand = [w for w in words if re.fullmatch(r"\d{6}", w[4]) and w[0] > c[2]
                    and abs((w[1] + w[3]) / 2 - cy[id(c)]) < 6]
            bound[id(c)] = min([w[0] for w in cand], default=c[2] + 170)
        dm = collections.defaultdict(list)
        for w in words:
            if re.fullmatch(r"AOI\d{11}", w[4]):
                continue
            wy = (w[1] + w[3]) / 2
            best, bd = None, 1e9
            for c in codes:
                if w[0] < c[2] - 1 or w[2] > bound[id(c)] - 1:
                    continue
                dd = abs(wy - cy[id(c)])
                if dd < bd:
                    bd, best = dd, c
            if best is not None and bd < 13:
                dm[id(best)].append(w)
        for _, k, p in events:
            if k == "CC":
                m = re.match(r"Centro de Costo:\s*([\d.]+)\s*-\s*(.+)", p)
                cc = (m.group(1), m.group(2).strip())
                oei = aei = None
            elif k == "OEI":
                oei = p
            elif k == "AEI":
                aei = p
            else:
                ws = sorted(dm[id(p)], key=lambda w: (w[1] // 4, w[0]))
                filas.append(dict(cc=cc[0], nombre=cc[1], oei=oei, aei=aei, aoi=p[4],
                                  desc=" ".join(w[4] for w in ws), page=pi + 1))
    return filas


def main(pdf, salida):
    filas = extraer(pdf)
    por_cc = collections.defaultdict(list)
    for f in filas:
        por_cc[f["cc"]].append(f)
    out = []
    for sigla, form, cc in FACULTADES:
        items = por_cc[cc]
        assert items, cc
        assert all(i["oei"] and i["aei"] and i["desc"] for i in items), cc
        out.append({
            "sigla": sigla, "formulario": form, "cc": cc, "nombre": items[0]["nombre"],
            "items": [[i["oei"], i["aei"], i["aoi"], i["desc"]] for i in items],
        })
    # Catalogos de OEI y AEI (texto completo) para no repetirlos en cada fila.
    oeis, aeis = {}, {}
    for o in out:
        for oei, aei, _, _ in o["items"]:
            co, ca = oei.split()[0], aei.split()[0]
            assert ca.replace("AEI", "OEI").rsplit(".", 1)[0] == co, (oei, aei)
            assert oeis.setdefault(co, oei) == oei and aeis.setdefault(ca, aei) == aei
    with open(salida, "w", encoding="utf-8") as fh:
        fh.write("/**\n * DATOS GENERADOS con tools/extraer_poi.py a partir del PDF POI 2026 (Anexo B-5,\n"
                 " * paginas 20-49). No editar a mano: volver a ejecutar el extractor.\n"
                 " * items: [codigo AEI, codigo AOI, denominacion AOI]; el OEI se deduce del AEI.\n */\n")
        fh.write("const POI_OEI = " + json.dumps(oeis, ensure_ascii=False, indent=1) + ";\n\n")
        fh.write("const POI_AEI = " + json.dumps(aeis, ensure_ascii=False, indent=1) + ";\n\n")
        fh.write("const POI_FACULTADES = [\n")
        for o in out:
            cab = {k: o[k] for k in ("sigla", "formulario", "cc", "nombre")}
            fh.write(" {\n  " + json.dumps(cab, ensure_ascii=False)[1:-1] + ",\n  \"items\": [\n")
            fh.write(",\n".join("   " + json.dumps([i[1].split()[0], i[2], i[3]], ensure_ascii=False)
                               for i in o["items"]))
            fh.write("\n  ]\n },\n")
        fh.write("];\n")
    print("facultades:", len(out), "actividades:", sum(len(o["items"]) for o in out))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
