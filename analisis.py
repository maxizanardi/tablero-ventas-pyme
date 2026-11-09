#!/usr/bin/env python3
"""
Tablero de ventas para una pyme de comercio online.
Dataset público: "Online Retail" (UCI Machine Learning Repository, CC BY 4.0),
transacciones reales de una tienda online del Reino Unido, dic-2010 a dic-2011.

Uso:
    python3 xlsx_a_csv.py "Online Retail.xlsx" datos/online_retail.csv
    python3 analisis.py datos/online_retail.csv

Genera:
    Tablero_Ventas_Pyme.xlsx   (tablero + tablas + registro de limpieza)
    datos/resumen.json         (cifras para la vista previa y el README)
"""
import csv, json, sys, datetime as dt, pathlib
from collections import defaultdict, Counter

BASE = pathlib.Path(__file__).parent
SRC = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "datos" / "online_retail.csv"
OUT = BASE / "Tablero_Ventas_Pyme.xlsx"

# Códigos que no son productos (portes, comisiones, ajustes manuales...)
NO_PRODUCTO = {"POST", "D", "M", "BANK CHARGES", "AMAZONFEE", "DOT", "CRUK", "S", "PADS", "B", "C2", "m"}
INICIO, FIN = dt.datetime(2010, 12, 1), dt.datetime(2011, 12, 1)   # 12 meses completos

PAISES_ES = {"United Kingdom": "Reino Unido", "Germany": "Alemania", "France": "Francia", "EIRE": "Irlanda",
             "Spain": "España", "Netherlands": "Países Bajos", "Belgium": "Bélgica", "Switzerland": "Suiza",
             "Portugal": "Portugal", "Australia": "Australia", "Norway": "Noruega", "Italy": "Italia",
             "Channel Islands": "Islas del Canal", "Finland": "Finlandia", "Cyprus": "Chipre", "Sweden": "Suecia",
             "Austria": "Austria", "Denmark": "Dinamarca", "Japan": "Japón", "Poland": "Polonia"}
MESES = "ene feb mar abr may jun jul ago sep oct nov dic".split()
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

# ---------------------------------------------------------------- 1. Carga y limpieza
log = Counter()
vistos = set()
ventas, devol = [], []
with open(SRC, encoding="utf-8") as f:
    for r in csv.DictReader(f):
        log["Filas originales"] += 1
        clave = tuple(r.values())
        if clave in vistos:
            log["Duplicados exactos (eliminados)"] += 1; continue
        vistos.add(clave)
        fecha = dt.datetime.fromisoformat(r["InvoiceDate"])
        if not (INICIO <= fecha < FIN):
            log["Fuera del período dic-2010 a nov-2011 (dic-2011 está incompleto)"] += 1; continue
        code = r["StockCode"].strip()
        if code.upper() in NO_PRODUCTO or code.startswith("gift_"):
            log["Códigos que no son producto: portes, comisiones, ajustes"] += 1; continue
        qty, precio = int(float(r["Quantity"])), float(r["UnitPrice"])
        fila = dict(fac=r["InvoiceNo"], code=code, desc=(r["Description"] or "").strip().capitalize(),
                    qty=qty, fecha=fecha, precio=precio, cli=(r["CustomerID"] or "").split(".")[0],
                    pais=r["Country"], imp=round(qty * precio, 2))
        if r["InvoiceNo"].startswith("C"):
            log["Devoluciones / anulaciones (separadas para su propio análisis)"] += 1
            devol.append(fila); continue
        if precio <= 0 or qty <= 0:
            log["Precio o cantidad en cero o negativos (eliminados)"] += 1; continue
        ventas.append(fila)
log["Filas de venta válidas"] = len(ventas)
sin_cli = sum(1 for v in ventas if not v["cli"])
log["Ventas sin cliente identificado (se usan en importes, no en clientes)"] = sin_cli

# Nombre único por código (el más frecuente)
nombres = defaultdict(Counter)
for v in ventas: nombres[v["code"]][v["desc"]] += 1
nombre = {c: n.most_common(1)[0][0] for c, n in nombres.items()}

# ---------------------------------------------------------------- 2. KPIs
fact = sum(v["imp"] for v in ventas)
pedidos = defaultdict(lambda: dict(imp=0.0, items=0, fecha=None, cli="", pais=""))
for v in ventas:
    p = pedidos[v["fac"]]; p["imp"] += v["imp"]; p["items"] += v["qty"]
    p["fecha"], p["cli"], p["pais"] = v["fecha"], v["cli"], v["pais"]
n_ped = len(pedidos)
ticket = fact / n_ped
ped_cli = defaultdict(set)
for fac, p in pedidos.items():
    if p["cli"]: ped_cli[p["cli"]].add(fac)
n_cli = len(ped_cli)
recurrentes = sum(1 for s in ped_cli.values() if len(s) >= 2)
fact_devol = -sum(d["imp"] for d in devol)
tasa_devol = fact_devol / fact

# ---------------------------------------------------------------- 3. Tablas
# Mensual
mes = defaultdict(lambda: [0.0, set(), set()])
for v in ventas:
    k = (v["fecha"].year, v["fecha"].month)
    mes[k][0] += v["imp"]; mes[k][1].add(v["fac"])
    if v["cli"]: mes[k][2].add(v["cli"])
t_mes = [(f"{MESES[m-1]}-{str(y)[2:]}", round(a, 2), len(b), len(c), round(a / len(b), 2))
         for (y, m), (a, b, c) in sorted(mes.items())]

# Productos
prod = defaultdict(lambda: [0.0, 0, set()])
for v in ventas:
    prod[v["code"]][0] += v["imp"]; prod[v["code"]][1] += v["qty"]; prod[v["code"]][2].add(v["fac"])
dev_prod = defaultdict(float)
for d in devol: dev_prod[d["code"]] += -d["imp"]
t_prod = sorted(((c, nombre[c], round(a, 2), q, len(f), round(dev_prod.get(c, 0), 2)) for c, (a, q, f) in prod.items()),
                key=lambda x: -x[2])
acum, n80 = 0.0, 0
for row in t_prod:
    acum += row[2]; n80 += 1
    if acum >= 0.8 * fact: break
pct_prod_80 = n80 / len(t_prod)

# Países
pais = defaultdict(lambda: [0.0, set(), set()])
for v in ventas:
    pais[v["pais"]][0] += v["imp"]; pais[v["pais"]][1].add(v["fac"])
    if v["cli"]: pais[v["pais"]][2].add(v["cli"])
t_pais = sorted(((PAISES_ES.get(p, p), round(a, 2), len(b), len(c)) for p, (a, b, c) in pais.items()), key=lambda x: -x[1])
uk_share = pais["United Kingdom"][0] / fact

# Día y hora
dh = defaultdict(float)
for v in ventas: dh[(v["fecha"].weekday(), v["fecha"].hour)] += v["imp"]
horas = sorted({h for _, h in dh})
dia_tot = defaultdict(float)
for (d, h), a in dh.items(): dia_tot[d] += a

# RFM simple
ref = FIN
rfm = {}
for c, facs in ped_cli.items():
    ult = max(pedidos[f]["fecha"] for f in facs)
    rfm[c] = dict(rec=(ref - ult).days, frec=len(facs), monto=sum(pedidos[f]["imp"] for f in facs))
def segmento(x):
    if x["rec"] <= 60 and x["frec"] >= 5: return "Campeones"
    if x["rec"] <= 90 and x["frec"] >= 2: return "Leales"
    if x["rec"] <= 60 and x["frec"] == 1: return "Nuevos"
    if x["rec"] <= 180: return "En riesgo"
    return "Perdidos"
seg = defaultdict(lambda: [0, 0.0])
for c, x in rfm.items():
    s = segmento(x); x["seg"] = s; seg[s][0] += 1; seg[s][1] += x["monto"]
ORDEN_SEG = ["Campeones", "Leales", "Nuevos", "En riesgo", "Perdidos"]
ACCION = {"Campeones": "Cuidarlos: acceso anticipado a novedades y atención prioritaria",
          "Leales": "Subir el ticket: packs y venta cruzada",
          "Nuevos": "Segunda compra: email de bienvenida con oferta a 30 días",
          "En riesgo": "Recuperar: campaña de reactivación antes de perderlos",
          "Perdidos": "Encuesta corta y oferta de vuelta; no invertir de más"}
t_seg = [(s, seg[s][0], round(seg[s][1], 2), ACCION[s]) for s in ORDEN_SEG]
top10_cli = sorted(rfm.values(), key=lambda x: -x["monto"])
share_top10pct = sum(x["monto"] for x in top10_cli[: max(1, n_cli // 10)]) / sum(x["monto"] for x in top10_cli)

# Mes pico y estacionalidad
mejor_mes = max(t_mes, key=lambda x: x[1])
peor_mes = min(t_mes, key=lambda x: x[1])
q4 = sum(x[1] for x in t_mes if x[0][:3] in ("sep", "oct", "nov"))
mejor_dia = max(dia_tot, key=dia_tot.get)
top_dev = sorted((r for r in t_prod if r[2] > 2000), key=lambda r: -(r[5] / r[2]))[:10]

def gbp(x): return "£" + f"{x:,.0f}".replace(",", ".")
pa = {r[0]: r for r in t_pais}
devol_txt = f"{tasa_devol*100:.1f}".replace(".", ",")
hallazgos = [
    f"Estacionalidad fuerte: septiembre a noviembre concentran el {q4/fact:.0%} de la facturación del año. {mejor_mes[0]} es el mejor mes "
    f"({gbp(mejor_mes[1])}) y {peor_mes[0]} el más flojo ({gbp(peor_mes[1])}). El stock y las campañas se preparan desde agosto.",
    f"Regla 80/20: el {pct_prod_80:.0%} de los productos ({n80} de {len(t_prod)}) genera el 80% de las ventas. "
    f"Ese grupo no se puede quedar sin stock; el resto se revisa o se liquida.",
    f"Clientes: el {recurrentes/n_cli:.0%} compró más de una vez y el 10% que más compra aporta el "
    f"{share_top10pct:.0%} de la facturación. Perder a unos pocos clientes grandes pesa mucho.",
    f"Hay {seg['En riesgo'][0]} clientes en riesgo (sin comprar hace 2 a 6 meses) que facturaron {gbp(seg['En riesgo'][1])}. "
    f"Reactivarlos cuesta menos que conseguir clientes nuevos.",
    f"Reino Unido es el {uk_share:.0%} de las ventas. Afuera, Países Bajos factura {gbp(pa['Países Bajos'][1])} con solo "
    f"{pa['Países Bajos'][3]} clientes (riesgo de concentración), mientras Alemania y Francia tienen más base "
    f"({pa['Alemania'][3]} y {pa['Francia'][3]} clientes) para crecer.",
    f"Devoluciones y anulaciones equivalen al {devol_txt}% de la facturación. La hoja Devoluciones muestra qué productos revisar primero.",
    f"Martes y jueves son los días más fuertes. Los sábados no hay pedidos y el domingo vende la mitad que un día de semana: "
    f"los emails y promociones rinden más de lunes a jueves.",
]

# ---------------------------------------------------------------- 4. Excel
import xlsxwriter
wb = xlsxwriter.Workbook(OUT)
AZUL, AZUL2, ORO, GRIS = "#0F2A4A", "#2F6DB5", "#F2C94C", "#5B6B7F"
F = lambda **k: wb.add_format({"font_name": "Calibri", **k})
f_tit = F(bold=True, font_size=22, font_color=AZUL)
f_sub = F(font_size=11, font_color=GRIS, italic=True)
f_kl = F(font_size=10, font_color="#FFFFFF", bg_color=AZUL, align="center", valign="vcenter")
f_kv = F(bold=True, font_size=20, font_color="#FFFFFF", bg_color=AZUL, align="center", valign="vcenter")
f_kv_gbp = F(bold=True, font_size=20, font_color="#FFFFFF", bg_color=AZUL, align="center", valign="vcenter", num_format="£#,##0")
f_kv_dec = F(bold=True, font_size=20, font_color="#FFFFFF", bg_color=AZUL, align="center", valign="vcenter", num_format="£#,##0.00")
f_kv_pct = F(bold=True, font_size=20, font_color="#FFFFFF", bg_color=AZUL, align="center", valign="vcenter", num_format="0.0%")
f_kv_int = F(bold=True, font_size=20, font_color="#FFFFFF", bg_color=AZUL, align="center", valign="vcenter", num_format="#,##0")
f_h = F(bold=True, font_color="#FFFFFF", bg_color=AZUL2, border=1, text_wrap=True, valign="vcenter")
f_gbp = F(num_format="£#,##0.00"); f_int = F(num_format="#,##0"); f_pct = F(num_format="0.0%")
f_txt = F(text_wrap=True, valign="top"); f_b = F(bold=True, font_color=AZUL, font_size=13)
f_hall = F(text_wrap=True, valign="top", font_size=11, bg_color="#F4F7FB", border=1, border_color="#D5DEEA")

def tabla(ws, r0, c0, cab, filas, formatos, nombre_tabla):
    ws.add_table(r0, c0, r0 + len(filas), c0 + len(cab) - 1,
                 {"name": nombre_tabla, "style": "Table Style Medium 2",
                  "columns": [{"header": h, "format": fm} for h, fm in zip(cab, formatos)], "data": [list(f) for f in filas]})

# --- Hojas de datos
ws_mes = wb.add_worksheet("Ventas_mes")
tabla(ws_mes, 0, 0, ["Mes", "Facturación (£)", "Pedidos", "Clientes", "Ticket medio (£)"], t_mes,
      [None, f_gbp, f_int, f_int, f_gbp], "tMes")
ws_mes.set_column(0, 0, 10); ws_mes.set_column(1, 4, 16)

ws_p = wb.add_worksheet("Productos")
tabla(ws_p, 0, 0, ["Código", "Producto", "Facturación (£)", "Unidades", "Pedidos", "Devuelto (£)"], t_prod,
      [None, None, f_gbp, f_int, f_int, f_gbp], "tProductos")
ws_p.set_column(0, 0, 10); ws_p.set_column(1, 1, 42); ws_p.set_column(2, 5, 15)

ws_pa = wb.add_worksheet("Paises")
tabla(ws_pa, 0, 0, ["País", "Facturación (£)", "Pedidos", "Clientes"], t_pais, [None, f_gbp, f_int, f_int], "tPaises")
ws_pa.set_column(0, 0, 18); ws_pa.set_column(1, 3, 15)

ws_s = wb.add_worksheet("Clientes_RFM")
tabla(ws_s, 0, 0, ["Segmento", "Clientes", "Facturación (£)", "Acción sugerida"], t_seg, [None, f_int, f_gbp, f_txt], "tSegmentos")
ws_s.set_column(0, 0, 14); ws_s.set_column(1, 2, 15); ws_s.set_column(3, 3, 60)
ws_s.write(8, 0, "Reglas del segmento (días desde la última compra al 01-12-2011 y número de pedidos):", f_b)
for i, t in enumerate(["Campeones: compró en los últimos 60 días y tiene 5 pedidos o más",
                       "Leales: compró en los últimos 90 días y tiene 2 pedidos o más",
                       "Nuevos: un solo pedido, en los últimos 60 días",
                       "En riesgo: última compra hace 61 a 180 días (y no entra en los anteriores)",
                       "Perdidos: más de 180 días sin comprar"]):
    ws_s.write(9 + i, 0, t)

ws_dh = wb.add_worksheet("Dia_hora")
ws_dh.write_row(0, 0, ["Día \\ Hora"] + [f"{h}h" for h in horas], f_h)
for d in range(7):
    ws_dh.write(d + 1, 0, DIAS[d], f_h)
    for j, h in enumerate(horas):
        ws_dh.write_number(d + 1, j + 1, round(dh.get((d, h), 0), 2), f_int)
ws_dh.conditional_format(1, 1, 7, len(horas), {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": AZUL2})
ws_dh.write_row(9, 0, ["Día", "Facturación (£)"], f_h)
for d in range(7):
    ws_dh.write(10 + d, 0, DIAS[d]); ws_dh.write_number(10 + d, 1, round(dia_tot.get(d, 0), 2), f_gbp)
ws_dh.set_column(0, 0, 11); ws_dh.set_column(1, len(horas), 9)

ws_d = wb.add_worksheet("Devoluciones")
tabla(ws_d, 0, 0, ["Código", "Producto", "Facturación (£)", "Unidades", "Pedidos", "Devuelto (£)"],
      top_dev, [None, None, f_gbp, f_int, f_int, f_gbp], "tDevol")
for i in range(len(top_dev)):
    ws_d.write_formula(i + 1, 6, f"=F{i+2}/C{i+2}", f_pct)
ws_d.write(0, 6, "% devuelto", f_h)
ws_d.set_column(0, 0, 10); ws_d.set_column(1, 1, 42); ws_d.set_column(2, 6, 14)
ws_d.write(len(top_dev) + 2, 0, "Productos con más de £2.000 de venta, ordenados por % de importe devuelto.", f_sub)

ws_ped = wb.add_worksheet("Pedidos")
filas_ped = sorted(((fac, p["fecha"].date().isoformat(), p["cli"] or "(sin dato)", PAISES_ES.get(p["pais"], p["pais"]),
                     p["items"], round(p["imp"], 2)) for fac, p in pedidos.items()), key=lambda x: x[1])
tabla(ws_ped, 0, 0, ["Factura", "Fecha", "Cliente", "País", "Unidades", "Importe (£)"], filas_ped,
      [None, None, None, None, f_int, f_gbp], "tPedidos")
ws_ped.set_column(0, 3, 14); ws_ped.set_column(4, 5, 13)

ws_l = wb.add_worksheet("Limpieza")
ws_l.write(0, 0, "Registro de limpieza de datos", f_b)
ws_l.write_row(2, 0, ["Paso", "Filas"], f_h)
for i, (k, n) in enumerate(log.items()):
    ws_l.write(3 + i, 0, k); ws_l.write_number(3 + i, 1, n, f_int)
ws_l.set_column(0, 0, 70); ws_l.set_column(1, 1, 12)

ws_f = wb.add_worksheet("Fuente")
for i, t in enumerate([
    "Fuente: Online Retail, UCI Machine Learning Repository (Chen, D., 2015). Licencia CC BY 4.0.",
    "https://archive.ics.uci.edu/dataset/352/online+retail",
    "Datos reales y anonimizados de una tienda online del Reino Unido que vende regalos y artículos de decoración,",
    "con muchos clientes mayoristas. Importes en libras esterlinas (£).",
    "Período analizado: 01-12-2010 a 30-11-2011 (12 meses completos). Diciembre 2011 está incompleto en el original y se excluye.",
    "Proyecto de portfolio de Maximiliano Zanardi. No es un trabajo para un cliente.",
]):
    ws_f.write(i, 0, t)
ws_f.set_column(0, 0, 120)

# --- Tablero
ws = wb.add_worksheet("Tablero")
ws.activate(); ws.set_first_sheet(); ws.hide_gridlines(2); ws.set_zoom(90)
ws.set_column(0, 0, 2); ws.set_column(1, 18, 11)
ws.write(0, 1, "Tablero de ventas · Tienda online (pyme)", f_tit)
ws.write(1, 1, "12 meses: dic-2010 a nov-2011 · Datos reales públicos (UCI Online Retail) · Importes en £", f_sub)
kpis = [("Facturación (£)", round(fact, 2), f_kv_int),
        ("Pedidos", n_ped, f_kv_int),
        ("Ticket medio (£)", round(ticket, 2), f_kv_int),
        ("Clientes", n_cli, f_kv_int),
        ("Recurrentes", recurrentes / n_cli, f_kv_pct),
        ("Devoluciones", tasa_devol, f_kv_pct)]
for i, (lab, val, fm) in enumerate(kpis):
    c = 1 + i * 3
    ws.merge_range(3, c, 3, c + 1, lab, f_kl)
    if isinstance(val, str): ws.merge_range(4, c, 4, c + 1, "", fm); ws.write_formula(4, c, val, fm)
    else: ws.merge_range(4, c, 4, c + 1, val, fm)
ws.set_row(3, 20); ws.set_row(4, 36)

def chart(tipo, titulo, series, sub=None, **kw):
    ch = wb.add_chart({"type": tipo, **({"subtype": sub} if sub else {})})
    for s in series: ch.add_series(s)
    ch.set_title({"name": titulo, "name_font": {"size": 12, "color": AZUL}})
    ch.set_legend({"none": True}); ch.set_chartarea({"border": {"none": True}})
    for k, v in kw.items(): getattr(ch, k)(v)
    return ch

n = len(t_mes)
ws.insert_chart(6, 1, chart("column", "Facturación por mes (£)", [{
    "categories": ["Ventas_mes", 1, 0, n, 0], "values": ["Ventas_mes", 1, 1, n, 1],
    "fill": {"color": AZUL2}, "gap": 60}], set_y_axis={"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#E6EBF2"}}}),
    {"x_scale": 1.45, "y_scale": 1.0})
ws.insert_chart(6, 10, chart("bar", "Top 10 productos por facturación (£)", [{
    "categories": ["Productos", 1, 1, 10, 1], "values": ["Productos", 1, 2, 10, 2], "fill": {"color": ORO}, "gap": 40}],
    set_y_axis={"reverse": True}, set_x_axis={"num_format": "#,##0"}), {"x_scale": 1.3, "y_scale": 1.0})
ws.insert_chart(22, 1, chart("bar", "Países fuera de Reino Unido (£)", [{
    "categories": ["Paises", 2, 0, 11, 0], "values": ["Paises", 2, 1, 11, 1], "fill": {"color": AZUL2}, "gap": 40}],
    set_y_axis={"reverse": True}, set_x_axis={"num_format": "#,##0"}), {"x_scale": 1.0, "y_scale": 1.0})
ws.insert_chart(22, 7, chart("column", "Clientes por segmento", [{
    "categories": ["Clientes_RFM", 1, 0, 5, 0], "values": ["Clientes_RFM", 1, 1, 5, 1],
    "points": [{"fill": {"color": c}} for c in ["#1E8E5A", "#2F6DB5", "#7FB8FF", ORO, "#C0504D"]],
    "data_labels": {"value": True}, "gap": 50}]), {"x_scale": 0.9, "y_scale": 1.0})
ws.insert_chart(22, 13, chart("column", "Facturación por día de la semana (£)", [{
    "categories": ["Dia_hora", 10, 0, 16, 0], "values": ["Dia_hora", 10, 1, 16, 1], "fill": {"color": AZUL}, "gap": 50}],
    set_y_axis={"num_format": "#,##0"}), {"x_scale": 0.85, "y_scale": 1.0})

ws.write(42, 1, "Qué dicen los datos y qué haría el dueño", f_b)
for i, h in enumerate(hallazgos):
    ws.merge_range(43 + i, 1, 43 + i, 18, f"{i+1}. {h}", f_hall); ws.set_row(43 + i, 32)

# Orden de pestañas: Tablero primero
wb.worksheets_objs.sort(key=lambda s: 0 if s.name == "Tablero" else 1)
wb.close()

# ---------------------------------------------------------------- 5. Resumen para README y vista previa
res = dict(facturacion=round(fact, 2), pedidos=n_ped, ticket=round(ticket, 2), clientes=n_cli,
           recurrentes=round(recurrentes / n_cli, 4), devoluciones=round(tasa_devol, 4),
           meses=t_mes, top10=[(r[1], r[2]) for r in t_prod[:10]], paises=t_pais[:11],
           segmentos=t_seg, dias=[(DIAS[d], round(dia_tot.get(d, 0), 2)) for d in range(7)],
           hallazgos=hallazgos, limpieza=dict(log), pct_prod_80=pct_prod_80, n80=n80, n_prod=len(t_prod))
(BASE / "datos").mkdir(exist_ok=True)
(BASE / "datos" / "resumen.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
print(json.dumps({k: v for k, v in res.items() if k in ("facturacion", "pedidos", "ticket", "clientes", "recurrentes", "devoluciones")}, indent=1))
print(*hallazgos, sep="\n"); print(dict(log))
