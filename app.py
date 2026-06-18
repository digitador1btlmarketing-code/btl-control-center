from flask import Flask, render_template, jsonify
import pandas as pd
from datetime import datetime
from collections import Counter
import tempfile
import os
import requests

app = Flask(__name__)

EXCEL_URL = "https://netorg12022531-my.sharepoint.com/:x:/g/personal/digitador1_btlmarketing_com_ni/IQA9XnkqjcL1T6_3eui1qzCKAeKU0dhfU-RCUuuyoKp7W7k?e=PeeJzc"
SHEET_NAME = "Sheet1"

DATE_COLUMNS = [
    "Fecha de entrega",
    "Fecha de aprobación",
    "Fecha de instalación",
    "Fecha de desinstalación",
]

TERMINADO_WORDS = ["terminado", "finalizado", "completado", "cerrado"]
PROCESO_WORDS = ["proceso", "produccion", "producción", "trabajando"]


def safe_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_done(status):
    s = safe_text(status).lower()
    return any(w in s for w in TERMINADO_WORDS)


def is_process(status):
    s = safe_text(status).lower()
    return any(w in s for w in PROCESO_WORDS)


def read_excel_safe():
    download_url = EXCEL_URL.replace("?e=", "?download=1&e=")

    response = requests.get(download_url, timeout=30)
    response.raise_for_status()

    temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    temp_path = temp_file.name
    temp_file.close()

    try:
        with open(temp_path, "wb") as f:
            f.write(response.content)

        return pd.read_excel(
            temp_path,
            sheet_name=SHEET_NAME,
            engine="openpyxl"
        )
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def get_priority(dias, status):
    if is_done(status):
        return {
            "key": "terminada",
            "label": "Terminada",
            "emoji": "✅",
            "rank": 5,
            "class": "p-done",
            "message": "Orden completada",
        }

    if dias is None:
        return {
            "key": "sin_fecha",
            "label": "Sin fecha",
            "emoji": "⚪",
            "rank": 4,
            "class": "p-none",
            "message": "Revisar fecha de entrega",
        }

    if dias < 0:
        return {
            "key": "vencida",
            "label": "Urgente",
            "emoji": "🔴",
            "rank": 0,
            "class": "p-urgent",
            "message": f"Vencida hace {abs(dias)} día(s)",
        }

    if dias == 0:
        return {
            "key": "hoy",
            "label": "Hoy",
            "emoji": "⏰",
            "rank": 1,
            "class": "p-today",
            "message": "Entrega hoy",
        }

    if dias <= 3:
        return {
            "key": "proxima",
            "label": "Próxima",
            "emoji": "🟡",
            "rank": 2,
            "class": "p-soon",
            "message": f"Faltan {dias} día(s)",
        }

    return {
        "key": "normal",
        "label": "Normal",
        "emoji": "🟢",
        "rank": 3,
        "class": "p-normal",
        "message": f"Faltan {dias} día(s)",
    }


def read_excel():
    try:
        raw_df = read_excel_safe()
    except Exception as e:
        return {
            "ok": False,
            "error": f"No pude leer el Excel desde SharePoint: {e}",
            "rows": [],
            "kpis": {},
            "responsables": [],
            "urgentes": [],
            "recientes": [],
        }

    df = raw_df.dropna(how="all").copy()

    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y")
            df[col] = df[col].fillna("")

    if "% avance" in df.columns:
        df["% avance"] = pd.to_numeric(df["% avance"], errors="coerce").fillna(0)
        df["Avance_Num"] = df["% avance"].apply(lambda x: x * 100 if x <= 1 else x)
    else:
        df["Avance_Num"] = 0

    if "Fecha de entrega" in raw_df.columns:
        fecha_entrega = pd.to_datetime(raw_df.loc[df.index, "Fecha de entrega"], errors="coerce")
        hoy = pd.Timestamp(datetime.now().date())
        df["Dias_Vencimiento_Calc"] = (fecha_entrega - hoy).dt.days
    elif "Dias de Vencimiento" in raw_df.columns:
        df["Dias_Vencimiento_Calc"] = pd.to_numeric(raw_df.loc[df.index, "Dias de Vencimiento"], errors="coerce")
    else:
        df["Dias_Vencimiento_Calc"] = None

    for col in df.columns:
        if col not in ["Avance_Num", "Dias_Vencimiento_Calc"]:
            df[col] = df[col].apply(safe_text)

    records = []

    for index, row in df.iterrows():
        avance = float(row.get("Avance_Num", 0) or 0)

        dias = row.get("Dias_Vencimiento_Calc", None)
        try:
            dias = int(dias) if pd.notna(dias) else None
        except Exception:
            dias = None

        status = safe_text(row.get("Status", "")) or "Sin status"
        prioridad = get_priority(dias, status)

        records.append({
            "index": int(index),
            "nombre": safe_text(row.get("Nombre", "")),
            "categoria": safe_text(row.get("Categoría (C)", "")),
            "op": safe_text(row.get("Número de OP", "")) or "Sin OP",
            "proyecto": safe_text(row.get("Proyecto", "")),
            "fecha_entrega": safe_text(row.get("Fecha de entrega", "")),
            "entrega": safe_text(row.get("Fecha de entrega", "")),
            "cliente": safe_text(row.get("Cliente", "")),
            "marca": safe_text(row.get("Marca", "")),
            "fecha_aprobacion": safe_text(row.get("Fecha de aprobación", "")),
            "entregar": safe_text(row.get("Entregar", "")),
            "entregar_a": safe_text(row.get("Entregar a", "")),
            "lugar_instalacion": safe_text(row.get("Lugar de instalación", "")),
            "fecha_instalacion": safe_text(row.get("Fecha de instalación", "")),
            "fecha_desinstalacion": safe_text(row.get("Fecha de desinstalación", "")),
            "responsable": safe_text(row.get("Presupuestista", "")) or "Sin responsable",
            "brief": safe_text(row.get("Brief", "")),
            "status": status,
            "avance": round(avance, 1),
            "dias": dias,
            "prioridad": prioridad,
            "sort_rank": prioridad["rank"],
        })

    records.sort(
        key=lambda r: (
            r["sort_rank"],
            9999 if r["dias"] is None else r["dias"],
            r["index"]
        )
    )

    pendientes = [r for r in records if not is_done(r["status"])]
    terminadas = [r for r in records if is_done(r["status"])]
    en_proceso = [r for r in records if is_process(r["status"])]
    vencidas = [r for r in pendientes if r["dias"] is not None and r["dias"] < 0]
    hoy = [r for r in pendientes if r["dias"] == 0]
    proximas = [r for r in pendientes if r["dias"] is not None and 1 <= r["dias"] <= 3]
    sin_fecha = [r for r in pendientes if r["dias"] is None]

    responsables = [
        {"nombre": nombre, "total": total}
        for nombre, total in Counter(r["responsable"] for r in pendientes).most_common(10)
    ]

    urgentes = [
        r for r in records
        if r["prioridad"]["key"] in ["vencida", "hoy", "proxima"]
    ][:10]

    recientes = sorted(records, key=lambda r: r["index"], reverse=True)[:8]

    return {
        "ok": True,
        "updated_at": datetime.now().strftime("%d/%m/%Y %I:%M:%S %p"),
        "rows": records,
        "urgentes": urgentes,
        "recientes": recientes,
        "responsables": responsables,
        "kpis": {
            "recibidas": len(records),
            "pendientes": len(pendientes),
            "en_proceso": len(en_proceso),
            "terminadas": len(terminadas),
            "vencidas": len(vencidas),
            "hoy": len(hoy),
            "proximas": len(proximas),
            "sin_fecha": len(sin_fecha),
        },
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/ordenes")
def api_ordenes():
    return jsonify(read_excel())


@app.after_request
def add_no_cache_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)