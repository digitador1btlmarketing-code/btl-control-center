from flask import Flask, render_template, jsonify
import pandas as pd
from datetime import datetime
import tempfile
import os
import requests

app = Flask(__name__)

EXCEL_URL = "https://netorg12022531-my.sharepoint.com/:x:/g/personal/digitador1_btlmarketing_com_ni/IQA9XnkqjcL1T6_3eui1qzCKAeKU0dhfU-RCUuuyoKp7W7k?e=PeeJzc"
SHEET_NAME = "Sheet1"

DATE_COLUMNS = [
    "Fecha de entrega",
    "Fecha de instalación",
    "Fecha de desinstalación",
]

TERMINADO_WORDS = ["terminado", "finalizado", "completado", "cerrado"]
PROCESO_WORDS = ["proceso", "produccion", "producción", "trabajando"]


# =========================
# UTIL
# =========================
def safe_text(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def is_done(status):
    s = safe_text(status).lower()
    return any(w in s for w in TERMINADO_WORDS)


def read_excel_safe():
    url = EXCEL_URL.replace("?e=", "?download=1&e=")

    r = requests.get(url, timeout=30)
    r.raise_for_status()

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(r.content)
    tmp.close()

    df = pd.read_excel(tmp.name, sheet_name=SHEET_NAME, engine="openpyxl")

    os.remove(tmp.name)
    return df


# =========================
# PRIORIDAD
# =========================
def get_priority(dias, status):
    if is_done(status):
        return {"key": "terminada", "rank": 5}

    if dias is None:
        return {"key": "sin_fecha", "rank": 4}

    if dias < 0:
        return {"key": "vencida", "rank": 0}

    if dias == 0:
        return {"key": "hoy", "rank": 1}

    if dias <= 3:
        return {"key": "proxima", "rank": 2}

    return {"key": "normal", "rank": 3}


# =========================
# CORE
# =========================
def read_excel():
    try:
        raw = read_excel_safe()
    except Exception as e:
        return {"ok": False, "error": str(e), "rows": []}

    df = raw.copy()

    # =========================
    # LIMPIEZA SEGURA
    # =========================
    df = df.dropna(how="all")

    if "Número de OP" not in df.columns:
        return {"ok": True, "rows": []}

    df["Número de OP"] = df["Número de OP"].astype(str).str.strip()

    # ❗ SOLO eliminar vacíos reales
    df = df[df["Número de OP"] != ""]
    df = df[df["Número de OP"].str.lower() != "nan"]

    # =========================
    # ELIMINAR DUPLICADOS REALES
    # =========================
    df = df.drop_duplicates(subset=["Número de OP"], keep="last")

    # =========================
    # FECHAS (SIN ROMPER INDEX)
    # =========================
    for c in DATE_COLUMNS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%d/%m/%Y")

    # =========================
    # DÍAS
    # =========================
    if "Fecha de entrega" in df.columns:
        fechas = pd.to_datetime(df["Fecha de entrega"], errors="coerce")
        hoy = pd.Timestamp(datetime.now().date())
        dias_list = (fechas - hoy).dt.days
    else:
        dias_list = None

    # =========================
    # BUILD RECORDS (SIN INDEX PROBLEM)
    # =========================
    records = []

    for i, row in df.iterrows():

        dias = None
        if dias_list is not None:
            try:
                dias = int(dias_list.loc[i])
            except:
                dias = None

        status = safe_text(row.get("Status", ""))

        records.append({
            "op": safe_text(row.get("Número de OP", "")),
            "marca": safe_text(row.get("Marca", "")),
            "presupuestista": safe_text(row.get("Presupuestista", "")),
            "lider_produccion": safe_text(row.get("Líder Producción", "")),
            "fecha_entrega": safe_text(row.get("Fecha de entrega", "")),
            "status": status,
            "hora_solicitud": safe_text(row.get("Hora de finalización", "")),
            "dias": dias,
            "prioridad": get_priority(dias, status),
        })

    # =========================
    # ORDEN
    # =========================
    records.sort(key=lambda r: (
        r["prioridad"]["rank"],
        r["dias"] if r["dias"] is not None else 9999
    ))

    return {
        "ok": True,
        "rows": records,
        "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }


# =========================
# ROUTES
# =========================
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/ordenes")
def api():
    return jsonify(read_excel())


@app.after_request
def nocache(resp):
    resp.headers["Cache-Control"] = "no-store"
    return resp


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)