"""
Backend · Control de Tardanzas · Innova Schools Campoy
"""
from flask import Flask, render_template, request, jsonify, send_file, session
from functools import wraps
import pandas as pd
from datetime import datetime, timedelta
import io
import os
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)

# Configuración de sesión
app.config['SESSION_COOKIE_SECURE'] = False
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)

# ============= CONFIGURACIÓN =============
HORARIO_DEFAULT = "07:20"
HORARIOS_POR_DIA = {}

# ============= USUARIOS =============
USERS = {
    "Carolina Magallanes": {"password": "40005191", "name": "Carolina Magallanes"},
    "directora": {"password": "40005191", "name": "Carolina Magallanes"},
    "subdirectora": {"password": "40005191", "name": "Carolina Magallanes"},
}

LAST_RESUMEN = None
LAST_DETALLE = None
LAST_MESES = None
LAST_DF = None


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user"):
            return jsonify({"error": "No autenticado"}), 401
        return f(*args, **kwargs)
    return decorated


def calcular_minutos_tarde_con_horario(fecha_str, registro_entrada_str):
    """
    Calcula los minutos de tardanza usando el horario configurado para ese día.
    Si no hay horario especial, usa 7:20 por defecto.
    """
    try:
        horario_str = HORARIOS_POR_DIA.get(fecha_str, "07:20")

        if ":" in horario_str:
            h1, m1 = map(int, horario_str.split(":"))
        else:
            h1, m1 = 7, 20

        if registro_entrada_str and ":" in str(registro_entrada_str):
            parts = str(registro_entrada_str).split(":")
            h2 = int(parts[0])
            m2 = int(parts[1])
        else:
            return 0

        minutos_entrada = h1 * 60 + m1
        minutos_registro = h2 * 60 + m2
        minutos_tarde = minutos_registro - minutos_entrada

        if minutos_tarde <= 0:
            return 0

        return minutos_tarde
    except Exception:
        return 0

def procesar_excel(df_raw: pd.DataFrame) -> dict:
    """
    Procesa el Excel con nombres de columnas exactos.
    """
    df = df_raw.copy()
    
    # *** NUEVO: Mapeo directo de nombres exactos (sin convertir a minúsculas) ***
    columnas_esperadas = {
        "Nombre del Colaborador": "nombre",
        "Cargo": "cargo",
        "Fecha": "fecha",
        "Día": "dia",
        "Tipo de Marcación": "tipo_marcacion",
        "Registro de Entrada": "registro_entrada",
        "Estado de Marcación": "estado_marcacion",
        "Horario de Entrada": "horario_entrada",
        "Tiempo Acumulado Tardanzas": "tiempo_acumulado"
    }
    
    # Renombrar solo las columnas que existen
    columnas_existentes = {}
    for col_original, col_nuevo in columnas_esperadas.items():
        if col_original in df.columns:
            columnas_existentes[col_original] = col_nuevo
        else:
            # Intentar buscar por coincidencia parcial (case-insensitive)
            for col in df.columns:
                if col.strip().lower() == col_original.lower():
                    columnas_existentes[col] = col_nuevo
                    break
    
    if not columnas_existentes:
        raise ValueError(f"No se encontraron columnas esperadas. Columnas disponibles: {list(df.columns)}")
    
    df = df.rename(columns=columnas_existentes)
    
    # Verificar columnas requeridas
    req = ["nombre", "cargo", "fecha", "tipo_marcacion", "registro_entrada"]
    faltantes = [c for c in req if c not in df.columns]
    if faltantes:
        raise ValueError(f"Columnas faltantes: {faltantes}. Columnas disponibles: {list(df.columns)}")
    
    # Filtrar PROFESORES
    df = df[df["cargo"].astype(str).str.strip().str.upper().str.contains("PROFESOR")].copy()
    if df.empty:
        raise ValueError("No se encontraron filas con Cargo que contenga 'PROFESOR'.")
    
    # Filtrar ENTRADAS
    df = df[df["tipo_marcacion"].astype(str).str.strip().str.lower() == "entrada"].copy()
    if df.empty:
        raise ValueError("No se encontraron filas con Tipo de Marcación = 'Entrada'.")
    
    # Procesar fechas
    df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["fecha"])
    df["mes_anio"] = df["fecha"].dt.strftime("%Y-%m")
    
    # Calcular minutos de tardanza
    df["minutos_tarde"] = df.apply(
        lambda row: calcular_minutos_tarde_con_horario(
            row["fecha"].strftime("%Y-%m-%d") if pd.notna(row["fecha"]) else "",
            str(row["registro_entrada"]) if pd.notna(row["registro_entrada"]) else ""
        ),
        axis=1
    )
    
    df["uso_cuponera"] = df["minutos_tarde"] >= 60
    df["minutos_tarde_efectivo"] = df.apply(
        lambda row: 0 if row["uso_cuponera"] else row["minutos_tarde"],
        axis=1
    )
    df["tiene_tardanza"] = df["minutos_tarde_efectivo"] > 0
    
    # Generar resumen
    resumen_rows = []
    for (nombre, mes), g in df.groupby(["nombre", "mes_anio"]):
        total = int(round(g["minutos_tarde_efectivo"].sum()))
        if total > 0:
            resumen_rows.append({
                "nombre": nombre,
                "mes": mes,
                "total_minutos": total,
                "tiempo_formateado": f"{total // 60}h {total % 60}min",
                "cuponeras": int(g["uso_cuponera"].sum()),
                "tardanzas": int(g["tiene_tardanza"].sum()),
            })
    
    # Generar detalle
    detalle = {}
    for (nombre, mes), g in df.groupby(["nombre", "mes_anio"]):
        key = f"{nombre}||{mes}"
        dias = []
        for _, row in g.sort_values("fecha").iterrows():
            fecha_str = row["fecha"].strftime("%Y-%m-%d")
            minutos = int(row["minutos_tarde_efectivo"])
            if minutos > 0:
                dias.append({
                    "fecha": row["fecha"].strftime("%d/%m/%Y"),
                    "fecha_iso": fecha_str,
                    "dia": str(row.get("dia", "") or "—"),
                    "horario_entrada": HORARIOS_POR_DIA.get(fecha_str, HORARIO_DEFAULT),
                    "registro_entrada": str(row.get("registro_entrada", "") or "—"),
                    "minutos": minutos,
                    "cuponera": False,
                })
            elif row["uso_cuponera"]:
                dias.append({
                    "fecha": row["fecha"].strftime("%d/%m/%Y"),
                    "fecha_iso": fecha_str,
                    "dia": str(row.get("dia", "") or "—"),
                    "horario_entrada": HORARIOS_POR_DIA.get(fecha_str, HORARIO_DEFAULT),
                    "registro_entrada": str(row.get("registro_entrada", "") or "—"),
                    "minutos": int(row["minutos_tarde"]),
                    "cuponera": True,
                })
        
        if dias:
            total_efectivo = int(round(g["minutos_tarde_efectivo"].sum()))
            detalle[key] = {
                "nombre": nombre,
                "mes": mes,
                "total_minutos": total_efectivo,
                "cuponeras": int(g["uso_cuponera"].sum()),
                "dias": dias,
            }
    
    meses_unicos = sorted(df["mes_anio"].unique())
    meses = [
        {
            "value": m,
            "label": pd.to_datetime(m + "-01").strftime("%B %Y").capitalize(),
        }
        for m in meses_unicos
    ]
    
    return {"resumen": resumen_rows, "detalle": detalle, "meses": meses}

def exportar_excel(resumen: list) -> io.BytesIO:
    df = pd.DataFrame(resumen)
    if df.empty:
        df = pd.DataFrame(columns=[
            "nombre", "tiempo_formateado", "cuponeras", "tardanzas", "total_minutos"
        ])
    df = df.sort_values("total_minutos", ascending=False).reset_index(drop=True)
    df.insert(0, "Ranking", range(1, len(df) + 1))
    df = df.rename(columns={
        "nombre": "Nombre del Profesor",
        "tiempo_formateado": "Tiempo (HH:MM)",
        "cuponeras": "Cuponeras Usadas",
        "tardanzas": "Días con Tardanza",
        "total_minutos": "Total Minutos Tarde",
    })
    cols = [
        "Ranking", "Nombre del Profesor", "Total Minutos Tarde",
        "Tiempo (HH:MM)", "Cuponeras Usadas", "Días con Tardanza",
    ]
    df = df[[c for c in cols if c in df.columns]]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Resumen", index=False)
    output.seek(0)
    return output


# ============= RUTAS =============

@app.route("/")
def index():
    return render_template("dashboard.html")


@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json() or {}
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""

    print(f"🔍 Login - Usuario: '{username}'")

    if username in USERS and USERS[username]["password"] == password:
        session.clear()
        session["user"] = username
        session["name"] = USERS[username]["name"]
        session["authenticated"] = True
        print(f"✅ Login exitoso: {username}")
        return jsonify({"ok": True, "name": USERS[username]["name"]})

    for key in USERS:
        if key.lower() == username.lower() and USERS[key]["password"] == password:
            session.clear()
            session["user"] = key
            session["name"] = USERS[key]["name"]
            session["authenticated"] = True
            print(f"✅ Login exitoso (case-insensitive): {key}")
            return jsonify({"ok": True, "name": USERS[key]["name"]})

    print(f"❌ Login fallido para: {username}")
    return jsonify({"error": "Usuario o contraseña incorrectos"}), 401


@app.route("/api/logout", methods=["POST"])
def api_logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/me")
def api_me():
    if session.get("user"):
        return jsonify({
            "authenticated": True,
            "name": session.get("name"),
            "user": session.get("user")
        })
    return jsonify({"authenticated": False})


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    global LAST_RESUMEN, LAST_DETALLE, LAST_MESES, LAST_DF
    print("📤 Recibiendo archivo...")
    
    try:
        if "file" not in request.files:
            return jsonify({"error": "No se envió ningún archivo"}), 400
        
        file = request.files["file"]
        
        if not file.filename:
            return jsonify({"error": "Nombre de archivo vacío"}), 400
        
        if not file.filename.endswith('.xlsx'):
            return jsonify({"error": "El archivo debe ser .xlsx"}), 400
        
        # Leer el Excel
        df = pd.read_excel(file)
        print(f"📊 Columnas encontradas: {list(df.columns)}")
        
        # Procesar
        data = procesar_excel(df)
        
        # Guardar en variables globales
        LAST_DF = df
        LAST_RESUMEN = data["resumen"]
        LAST_DETALLE = data["detalle"]
        LAST_MESES = data["meses"]
        
        print(f"✅ Archivo procesado: {len(LAST_RESUMEN)} registros")
        return jsonify(data)
        
    except Exception as e:
        import traceback
        error_msg = str(e)
        print(f"❌ Error: {error_msg}")
        print(traceback.format_exc())
        return jsonify({"error": error_msg}), 400

@app.route("/api/horarios", methods=["POST"])
@login_required
def set_horarios_por_dia():
    global HORARIOS_POR_DIA
    data = request.get_json() or {}
    horarios = data.get("horarios", {})

    print(f"📥 Recibiendo horarios: {horarios}")

    # LIMPIAR horarios existentes antes de guardar los nuevos
    # Esto asegura que cuando se limpian los horarios, realmente se eliminan
    HORARIOS_POR_DIA = {}

    for fecha, hora in horarios.items():
        if ":" in hora:
            h, m = hora.split(":")
            if 0 <= int(h) <= 23 and 0 <= int(m) <= 59:
                HORARIOS_POR_DIA[fecha] = hora

    print(f"✅ Horarios guardados: {HORARIOS_POR_DIA}")
    return jsonify({"ok": True, "horarios_aplicados": len(HORARIOS_POR_DIA)})


@app.route("/api/horarios", methods=["GET"])
@login_required
def get_horarios_por_dia():
    return jsonify(HORARIOS_POR_DIA)


@app.route("/api/recalcular", methods=["POST"])
@login_required
def recalcular():
    global LAST_RESUMEN, LAST_DETALLE, LAST_MESES, LAST_DF, HORARIOS_POR_DIA

    if LAST_DF is None:
        return jsonify({"error": "No hay datos para recalcular"}), 400

    try:
        print(f"🔄 Recalculando con horarios: {HORARIOS_POR_DIA}")
        
        # Reprocesar el archivo con los horarios actuales
        data = procesar_excel(LAST_DF)
        
        LAST_RESUMEN = data["resumen"]
        LAST_DETALLE = data["detalle"]
        LAST_MESES = data["meses"]

        print(f"✅ Datos recalculados: {len(LAST_RESUMEN)} registros")
        print(f"📅 Horarios aplicados: {HORARIOS_POR_DIA}")

        return jsonify({
            "ok": True,
            "data": {
                "resumen": LAST_RESUMEN,
                "detalle": LAST_DETALLE,
                "meses": LAST_MESES
            }
        })
    except Exception as e:
        print(f"❌ Error al recalcular: {str(e)}")
        return jsonify({"error": str(e)}), 400


@app.route("/export")
@login_required
def export():
    global LAST_RESUMEN
    if not LAST_RESUMEN:
        return "No hay datos para exportar", 400
    meses = sorted({r["mes"] for r in LAST_RESUMEN})
    mes = meses[-1] if meses else None
    filtrado = [r for r in LAST_RESUMEN if r["mes"] == mes] if mes else LAST_RESUMEN
    output = exportar_excel(filtrado)
    filename = f"resumen_tardanzas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    os.makedirs("templates", exist_ok=True)
    os.makedirs("static", exist_ok=True)
    app.run(debug=True, port=5000)