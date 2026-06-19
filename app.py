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

    response = requests.get(
        download_url,
        timeout=30,
        headers={
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        },
    )
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


def format_date_value(value):
    if pd.isna(value):
        return ""

    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return safe_text(value)

    if dt.year <= 1970:
        return ""

    return dt.strftime("%d/%m/%Y")


def format_datetime_value(value):
    if pd.isna(value):
        return ""

    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return safe_text(value)

    if dt.year <= 1970:
        return ""

    return dt.strftime("%Y-%m-%d %H:%M:%S")


def read_excel():
    try:
        raw_df = read_excel_safe()
    except Exception as e:
        return {
            "ok": False,
            "error": f"No pude leer el Excel desde SharePoint: {e}",
            "rows": [],
            "kpis": {},
            "presupuestistas": [],
            "responsables": [],
            "urgentes": [],
            "recientes": [],
        }

    df = raw_df.dropna(how="all").copy()
    df = df.reset_index(drop=True)

    # Dejar solo filas con OP real
    if "Número de OP" in df.columns:
        df["Número de OP"] = df["Número de OP"].fillna("").astype(str).str.strip()
        df = df[
            (df["Número de OP"] != "") &
            (df["Número de OP"].str.lower() != "nan") &
            (df["Número de OP"].str.lower() != "sin op")
        ].copy()

    # Cuando Forms + Power Automate duplican:
    # conservar la fila buena por Id, que es la que tiene Hora de inicio llena.
    if "Id" in df.columns:
        df["Id"] = df["Id"].fillna("").astype(str).str.strip()

        if "Hora de inicio" in df.columns:
            df["_hora_inicio_llena"] = (
                df["Hora de inicio"]
                .fillna("")
                .astype(str)
                .str.strip()
                .ne("")
            )

            df = (
                df.sort_values("_hora_inicio_llena", ascending=True)
                  .drop_duplicates(subset=["Id"], keep="last")
                  .drop(columns=["_hora_inicio_llena"])
            )
        else:
            df = df.drop_duplicates(subset=["Id"], keep="last")

    elif "Número de OP" in df.columns:
        df = df.drop_duplicates(subset=["Número de OP"], keep="last")

    df = df.reset_index(drop=True)

    # Guardar fecha real para cálculo de días antes de convertir a texto
    if "Fecha de entrega" in df.columns:
        fecha_entrega_dt = pd.to_datetime(df["Fecha de entrega"], errors="coerce")
        hoy = pd.Timestamp(datetime.now().date())
        df["Dias_Vencimiento_Calc"] = (fecha_entrega_dt - hoy).dt.days
    else:
        df["Dias_Vencimiento_Calc"] = None

    # Fechas para mostrar
    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(format_date_value)

    if "Hora de finalización" in df.columns:
        df["Hora de finalización"] = df["Hora de finalización"].apply(format_datetime_value)

    if "% avance" in df.columns:
        avance_series = (
            df["% avance"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.strip()
        )
        df["% avance"] = pd.to_numeric(avance_series, errors="coerce").fillna(0)
        df["Avance_Num"] = df["% avance"].apply(lambda x: x * 100 if x <= 1 else x)
    else:
        df["Avance_Num"] = 0

    for col in df.columns:
        if col not in ["Avance_Num", "Dias_Vencimiento_Calc"]:
            df[col] = df[col].apply(safe_text)

    records = []

    for i, row in df.iterrows():
        avance = float(row.get("Avance_Num", 0) or 0)

        dias = row.get("Dias_Vencimiento_Calc", None)
        try:
            dias = int(dias) if pd.notna(dias) else None
        except Exception:
            dias = None

        status = safe_text(row.get("Status", "")) or "Sin status"
        prioridad = get_priority(dias, status)
        presupuestista = safe_text(row.get("Presupuestista", "")) or "Sin presupuestista"

        records.append({
            "index": i,
            "hora_solicitud": safe_text(row.get("Hora de finalización", "")),
            "nombre": safe_text(row.get("Nombre", "")),
            "categoria": safe_text(row.get("Categoría (C)", "")),
            "op": safe_text(row.get("Número de OP", "")) or "Sin OP",
            "proyecto": safe_text(row.get("Proyecto", "")),
            "fecha_entrega": safe_text(row.get("Fecha de entrega", "")),
            "entrega": safe_text(row.get("Fecha de entrega", "")),
            "cliente": safe_text(row.get("Cliente", "")),
            "marca": safe_text(row.get("Marca", "")),
            "entregar": safe_text(row.get("Entregar", "")),
            "lugar_instalacion": safe_text(row.get("Lugar de instalación", "")),
            "fecha_instalacion": safe_text(row.get("Fecha de instalación", "")),
            "fecha_desinstalacion": safe_text(row.get("Fecha de desinstalación", "")),
            "presupuestista": presupuestista,
            "responsable": presupuestista,
            "lider_produccion": safe_text(row.get("Líder Producción", "")),
            "brief": safe_text(row.get("Brief", "")),
            "status": status,
            "avance": round(avance, 1),
            "dias": dias,
            "dias_vencimiento": dias,
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

    carga_presupuestista = [
        {"nombre": nombre, "total": total}
        for nombre, total in Counter(r["presupuestista"] for r in pendientes).most_common(10)
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
        "presupuestistas": carga_presupuestista,
        "responsables": carga_presupuestista,
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