#!/usr/bin/env python3
"""
make_sample_data.py — generate a synthetic Hacienda-format export.

Produces a 'Reporte de ventas general'-shaped .xlsx with FAKE names and FAKE
cédulas but realistic structure and patterns (repeat clients, companies, a
foreigner, and a few unpaid / reversed invoices) so fotorentas_risk.py can be
demonstrated publicly without exposing any real customer PII.

Output is NOT real data. Any resemblance to a real person/company is accidental.
"""
import random
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

random.seed(42)  # reproducible

ACT = ("7730.0-Alquiler  y arrendamiento de otros tipo de maquinaria, "
       "equipo y bienes tangibles")
MONTHS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
          "agosto", "setiembre", "octubre", "noviembre", "diciembre"]
AMOUNTS = [15000, 20000, 25000, 30000, 36000, 40000, 45000, 50000,
           55000, 60000, 70000, 75000, 80000, 100000, 120000, 153000]
AMT_W = [3, 5, 5, 6, 6, 7, 6, 7, 4, 4, 3, 3, 2, 2, 1, 1]

FIRST = ["Andrés", "María", "José", "Carolina", "Luis", "Daniela", "Carlos",
         "Gabriela", "Roberto", "Valeria", "Diego", "Natalia", "Fernando",
         "Paula", "Mauricio", "Andrea", "Esteban", "Sofía", "Alejandro",
         "Laura", "Ricardo", "Melissa", "Óscar", "Jimena", "Sebastián"]
LAST = ["Rojas", "Vargas", "Jiménez", "Mora", "Solano", "Castro", "Herrera",
        "Núñez", "Araya", "Sánchez", "Quesada", "Campos", "Ramírez", "Chaves",
        "Brenes", "Soto", "Vega", "Calderón", "Montero", "Salas", "Acuña"]
COMPANY = ["PRODUCCIONES ANDINA", "ESTUDIO LUMEN CR", "MEDIA NORTE",
           "AGENCIA PIXELADA", "ENFOQUE DIGITAL CR", "CASA AUDIOVISUAL"]
SUFFIX = ["SOCIEDAD ANONIMA", "SOCIEDAD DE RESPONSABILIDAD LIMITADA", "LIMITADA"]
FOREIGN = [("Mariana Duarte", "AB1234567"), ("Tomás Bianchi", "YC0099812")]


def fake_individual():
    return (f"{random.choice(FIRST)} {random.choice(LAST)} {random.choice(LAST)}",
            "Cédula Física", str(random.randint(1, 7)) + f"{random.randint(0, 99999999):08d}")


def fake_company():
    return (f"{random.choice(COMPANY)} {random.choice(SUFFIX)}",
            "Cédula Jurídica", "3101" + f"{random.randint(0, 999999):06d}")


# ---- build a client pool with a few designated "regulars" -----------------
pool = [fake_individual() for _ in range(50)]
pool += [fake_company() for _ in range(6)]
pool += [(n, "Otro (Extranjero)", pid) for n, pid in FOREIGN]
regulars = random.sample(pool, 10)  # appear more often

# ---- generate invoices in date order over 2024–2025 -----------------------
rows = []
reserva = 3000
day = datetime(2024, 1, 8)
end = datetime(2025, 12, 20)
while day <= end:
    if day.weekday() < 5 and random.random() < 0.42:  # ~weekdays, some skipped
        client = random.choice(regulars if random.random() < 0.45 else pool)
        rows.append((day, client, random.choices(AMOUNTS, AMT_W)[0],
                     "Contado" if random.random() < 0.8 else "Crédito",
                     "Pagada", "Factura" if random.random() < 0.85 else "Tiquete"))
        reserva += 1
    day += timedelta(days=1)

# ---- inject the only real risk: a few unpaid / reversed invoices ----------
reg_ind = next(c for c in regulars if c[1] == "Cédula Física")
reg_co = next(c for c in regulars if c[1] == "Cédula Jurídica")
new_guy = fake_individual()
rows.append((datetime(2025, 3, 14), reg_ind, 45000, "Crédito", "Pendiente", "Factura"))
rows.append((datetime(2025, 9, 5), reg_co, 60000, "Crédito", "Pendiente", "Factura"))
rows.append((datetime(2024, 11, 22), new_guy, 50000, "Crédito", "Aceptada", "Nota de crédito"))
rows.sort(key=lambda r: r[0])

# ---- write the Hacienda-shaped workbook -----------------------------------
HEADERS = ["Fecha de emisión", "Registrante", "Sucursal", "Terminal",
           "Actividad económica", "Grupo comercial", "Tipo de identificación",
           "N° de identificación", "Nombre", "Tipo de documento",
           "N° de documento", "Condición de venta", "Estado", "Moneda",
           "Tipo de cambio", "Monto total", "Descuento", "Subtotal gravado",
           "Subtotal exento", "Subtotal no sujeto", "Subtotal",
           "Impuestos selectivo de consumo", "Impuestos único a los combustibles",
           "Impuestos específico de bebidas alcohólicas",
           "Impuestos específico de bebidas envasadas", "Impuestos tabaco",
           "Impuestos específico al cemento", "Otros impuestos",
           "Monto exoneracion", "IVA 0.5%", "IVA 1%", "IVA 2%", "IVA 4%",
           "IVA 8%", "IVA 13%", "Total IVA", "IVA devuelto", "Otros cargos",
           "Total impuesto asumido", "Total", "Comentarios"]

wb = Workbook()
ws = wb.active
ws.title = "DOC. 4.4"
ws.append(["Reporte de ventas general"])
ws.append([])
ws.append(["Empresa:", "Servicios XBM Sociedad Anónima (Fotorentas) — DEMO DATA"])
ws.append(["N° de identificación:", "3101743541", "", "Usuario:", "demo@fotorentas.cr"])
ws.append([])
ws.append(["Cantidad de registros: ", len(rows)])
ws.append([])
ws.append(["DATOS GENERALES", "", "", "", "", "", "CLIENTE", "", "", "DOCUMENTO"])
ws.append(HEADERS)

r = 4000
for emitted, (name, idtype, pid), gross, terms, state, doc in rows:
    sub = round(gross / 1.13, 2)
    iva = round(gross - sub, 2)
    deliver = emitted + timedelta(days=1)
    ret = deliver + timedelta(days=random.randint(1, 4))
    note = (f"Reserva: {r}  Fecha de entrega: {deliver.day} de {MONTHS[deliver.month-1]}"
            f"  Fecha de devolución: {ret.day} de {MONTHS[ret.month-1]}")
    row = [emitted, "demo@fotorentas.cr", "001", "00001", ACT, "", idtype, pid,
           name, doc, "", terms, state, "CRC", 1, sub, 0, sub, 0, 0, sub,
           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, iva, iva, 0, 0, 0, gross, note]
    ws.append(row)
    r += 1

out = Path("sample_data/ventas_sample.xlsx")
out.parent.mkdir(exist_ok=True)
wb.save(out)
print(f"Wrote {out} — {len(rows)} synthetic invoices "
      f"({len({c[2] for _, c, *_ in rows})} unique clients).")
