from flask import Flask, render_template, jsonify, request
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


def get_value(row, *names):
    for name in names:
        if name in row.index:
            value = safe_text(row.get(name, ""))
            if value:
                return value
    return ""


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

    if 1 <= dias <= 3:
        return {
            "key": "proxima",
            "label": "Próxima",
            "emoji": "🔥",
            "rank": 2,
            "class": "p-fire",
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

    if "Número de OP" not in df.columns:
        return {
            "ok": True,
            "updated_at": datetime.now().strftime("%d/%m/%Y %I:%M:%S %p"),
            "rows": [],
            "urgentes": [],
            "recientes": [],
            "presupuestistas": [],
            "responsables": [],
            "kpis": {
                "recibidas": 0,
                "pendientes": 0,
                "en_proceso": 0,
                "terminadas": 0,
                "vencidas": 0,
                "hoy": 0,
                "proximas": 0,
                "sin_fecha": 0,
            },
        }

    df["Número de OP"] = df["Número de OP"].fillna("").astype(str).str.strip()
    df = df[
        (df["Número de OP"] != "") &
        (df["Número de OP"].str.lower() != "nan") &
        (df["Número de OP"].str.lower() != "sin op")
    ].copy()

    # Elegir la fila más completa cuando Forms + Power Automate duplican la OP
    cols_preferidas = [
        "Hora de finalización",
        "Hora de inicio",
        "Nombre",
        "Presupuestista",
        "Líder Producción",
        "Lider Produccion",
        "Status",
        "% avance",
    ]

    df["_score_completo"] = 0

    for col in cols_preferidas:
        if col in df.columns:
            df["_score_completo"] += (
                df[col]
                .fillna("")
                .astype(str)
                .str.strip()
                .ne("")
            ).astype(int)

    df = (
        df.sort_values("_score_completo", ascending=True)
          .drop_duplicates(subset=["Número de OP"], keep="last")
          .drop(columns=["_score_completo"])
          .copy()
          .reset_index(drop=True)
    )

    if "Fecha de entrega" in df.columns:
        fecha_entrega_dt = pd.to_datetime(df["Fecha de entrega"], errors="coerce")
        hoy = pd.Timestamp(datetime.now().date())
        df["Dias_Vencimiento_Calc"] = (fecha_entrega_dt - hoy).dt.days
    else:
        df["Dias_Vencimiento_Calc"] = None

    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(format_date_value)

    if "Hora de finalización" in df.columns:
        df["Hora de finalización"] = df["Hora de finalización"].apply(format_datetime_value)

    if "Hora de inicio" in df.columns:
        df["Hora de inicio"] = df["Hora de inicio"].apply(format_datetime_value)

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

        presupuestista = (
            get_value(row, "Presupuestista")
            or "Sin presupuestista"
        )

        lider_produccion = get_value(
            row,
            "Líder Producción",
            "Lider Produccion",
            "Líder de Producción",
            "Lider de Produccion",
        )

        hora_solicitud = get_value(
            row,
            "Hora de finalización",
            "Hora Solicitud",
            "Hora solicitud",
            "Hora de inicio",
        )

        records.append({
            "index": i,
            "hora_solicitud": hora_solicitud,
            "nombre": get_value(row, "Nombre"),
            "categoria": get_value(row, "Categoría (C)", "Categoria (C)", "Categoría", "Categoria").upper(),
            "op": get_value(row, "Número de OP", "Numero de OP") or "Sin OP",
            "proyecto": get_value(row, "Proyecto"),
            "fecha_entrega": get_value(row, "Fecha de entrega"),
            "entrega": get_value(row, "Fecha de entrega"),
            "cliente": get_value(row, "Cliente"),
            "marca": get_value(row, "Marca"),
            "entregar": get_value(row, "Entregar"),
            "lugar_instalacion": get_value(row, "Lugar de instalación", "Lugar de instalacion"),
            "fecha_instalacion": get_value(row, "Fecha de instalación", "Fecha de instalacion"),
            "fecha_desinstalacion": get_value(row, "Fecha de desinstalación", "Fecha de desinstalacion"),
            "presupuestista": presupuestista,
            "responsable": presupuestista,
            "lider_produccion": lider_produccion,
            "brief": get_value(row, "Brief"),
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

    urgentes = [
        r for r in records
        if r["prioridad"]["key"] in ["vencida", "hoy", "proxima"]
    ][:10]

    recientes = sorted(records, key=lambda r: r["index"], reverse=True)[:8]

    carga_presupuestista = []
    for nombre in sorted(set(r["presupuestista"] for r in pendientes)):
        carga_presupuestista.append({
            "nombre": nombre,
            "total": sum(1 for r in pendientes if r["presupuestista"] == nombre)
        })

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


@app.route("/api/ordenes")
def api_ordenes():
    data = read_excel()

    if not data.get("ok"):
        return jsonify(data)

    categoria = request.args.get("categoria", "").strip().upper()

    if categoria:
        data["rows"] = [
            r for r in data["rows"]
            if (r.get("categoria") or "").strip().upper() == categoria
        ]

        data["urgentes"] = [
            r for r in data.get("urgentes", [])
            if (r.get("categoria") or "").strip().upper() == categoria
        ]

        data["recientes"] = [
            r for r in data.get("recientes", [])
            if (r.get("categoria") or "").strip().upper() == categoria
        ]

    return jsonify(data)


@app.route("/")
def home():
    return render_template("index.html")


@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)