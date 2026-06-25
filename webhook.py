"""
webhook.py — Chaparro Fernández Wealth Management
Bot de WhatsApp via Twilio + Claude

Despliegue: Railway (railway.app)
Variables de entorno necesarias:
  ANTHROPIC_API_KEY   → tu clave de Anthropic
  TWILIO_AUTH_TOKEN   → token de Twilio (para verificar que el mensaje viene de Twilio)
  GDRIVE_FILE_ID      → 1CImiIbg7kSLrYNpWgzPHEBCmI3KRVlBX
  GOOGLE_CREDS_JSON   → contenido del JSON de credenciales de Google (service account)
"""

import os
import json
import base64
import hashlib
import hmac
import calendar
import re
from datetime import datetime
from io import BytesIO

import requests
from flask import Flask, request, Response
import pandas as pd

# ── Números autorizados (lista blanca) ───────────────────────────────────────
NUMEROS_AUTORIZADOS = {
    "whatsapp:+34622141605",   # Yuri
    # "whatsapp:+34XXXXXXXXX", # Socio 2  ← descomentar cuando quieras añadir
    # "whatsapp:+34XXXXXXXXX", # Socio 3  ← descomentar cuando quieras añadir
}

# ── Constantes del fondo (igual que app.py) ──────────────────────────────────
GDRIVE_FILE_ID   = os.environ.get("GDRIVE_FILE_ID", "1CImiIbg7kSLrYNpWgzPHEBCmI3KRVlBX")
ARCHIVO          = "/tmp/inversiones.xlsx"
HOJA_INVERSIONES = "INVERSIONES"
HOJA_CALENDARIO  = "CALENDARIO_NOTAS"
HOJA_CONTROL     = "CONTROL_NOTAS"

TASA_ANUAL_FUTBOL   = 0.15
TASA_ANUAL_MOTOCLICK= 0.25
TASA_ANUAL_PARAGUAY = 0.15
TASA_ANUAL_BOLIVIA  = 0.15
TASA_ANUAL_BITCOIN  = 0.20

# ── Historial en memoria por número (dura mientras el servidor esté vivo) ────
# { "whatsapp:+34622141605": [ {role, content}, ... ] }
historial_por_numero: dict[str, list] = {}

app = Flask(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# DESCARGA DEL EXCEL DESDE GOOGLE DRIVE
# ═══════════════════════════════════════════════════════════════════════════════

def descargar_excel_drive():
    """Descarga inversiones.xlsx desde Google Drive usando service account."""
    creds_json = os.environ.get("GOOGLE_CREDS_JSON", "")
    if not creds_json:
        raise RuntimeError("GOOGLE_CREDS_JSON no configurado")

    creds = json.loads(creds_json)

    # Obtener token OAuth2
    import urllib.parse, time
    import jwt  # PyJWT

    now = int(time.time())
    payload = {
        "iss": creds["client_email"],
        "scope": "https://www.googleapis.com/auth/drive.readonly",
        "aud": "https://oauth2.googleapis.com/token",
        "iat": now,
        "exp": now + 3600,
    }
    token_jwt = jwt.encode(payload, creds["private_key"], algorithm="RS256")
    resp = requests.post("https://oauth2.googleapis.com/token", data={
        "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
        "assertion": token_jwt,
    })
    access_token = resp.json()["access_token"]

    # Descargar el archivo
    url = f"https://www.googleapis.com/drive/v3/files/{GDRIVE_FILE_ID}?alt=media"
    r = requests.get(url, headers={"Authorization": f"Bearer {access_token}"})
    with open(ARCHIVO, "wb") as f:
        f.write(r.content)


def cargar_excel():
    """Descarga y carga el Excel. Devuelve (df_inv, df_cal, df_control)."""
    descargar_excel_drive()
    df_inv     = pd.read_excel(ARCHIVO, sheet_name=HOJA_INVERSIONES)
    df_cal     = pd.read_excel(ARCHIVO, sheet_name=HOJA_CALENDARIO)
    try:
        df_control = pd.read_excel(ARCHIVO, sheet_name=HOJA_CONTROL)
    except Exception:
        df_control = pd.DataFrame()

    # Normalizar columnas
    df_inv.columns     = [str(c).strip().lower() for c in df_inv.columns]
    df_cal.columns     = [str(c).strip().lower() for c in df_cal.columns]
    df_control.columns = [str(c).strip().lower() for c in df_control.columns]

    for col in ["id_inversion","inversor","tipo_inversion","subtipo_inversion",
                "nombre_activo","tipo_operacion","capital_nuevo_real","motivo"]:
        if col in df_inv.columns:
            df_inv[col] = df_inv[col].fillna("").astype(str).str.strip()

    for col in ["fecha_inversion","fecha_final_inversion"]:
        if col in df_inv.columns:
            df_inv[col] = pd.to_datetime(df_inv[col], errors="coerce", dayfirst=True)

    for col in ["capital_invertido","interes_inversor_anual","interes_nota_anual"]:
        if col in df_inv.columns:
            df_inv[col] = pd.to_numeric(df_inv[col], errors="coerce").fillna(0)
        else:
            df_inv[col] = 0

    if "tipo_evento" in df_cal.columns:
        df_cal["tipo_evento"] = df_cal["tipo_evento"].fillna("").astype(str).str.strip().str.upper()
    if "fecha" in df_cal.columns:
        df_cal["fecha"] = pd.to_datetime(df_cal["fecha"], errors="coerce", dayfirst=True)
    if "nota" in df_cal.columns:
        df_cal["nota"] = pd.to_numeric(df_cal["nota"], errors="coerce")

    return df_inv, df_cal, df_control


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def ultimo_dia_mes(anio: int, mes: int) -> int:
    return calendar.monthrange(anio, mes)[1]


def capital_activo_en_fecha(df: pd.DataFrame, fecha) -> float:
    fecha_ts = pd.Timestamp(fecha).normalize()
    df2 = df.copy()
    df2["fecha_inversion"]       = pd.to_datetime(df2.get("fecha_inversion"), errors="coerce", dayfirst=True)
    df2["fecha_final_inversion"] = pd.to_datetime(df2.get("fecha_final_inversion"), errors="coerce", dayfirst=True)
    df2["capital_invertido"]     = pd.to_numeric(df2.get("capital_invertido"), errors="coerce").fillna(0)
    df2["tipo_op"]               = df2["tipo_operacion"].astype(str).str.strip().str.upper()

    activas = df2[
        df2["tipo_op"].isin(["NUEVA","REINVERSION"]) &
        df2["fecha_inversion"].notna() &
        (df2["fecha_inversion"] <= fecha_ts) &
        (df2["fecha_final_inversion"].isna() | (df2["fecha_final_inversion"] >= fecha_ts))
    ]
    return float(activas["capital_invertido"].sum())


# ═══════════════════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL CONTEXTO (misma lógica que _contexto_excel en app.py)
# ═══════════════════════════════════════════════════════════════════════════════

def construir_contexto(df_inv: pd.DataFrame, df_cal: pd.DataFrame,
                       df_control: pd.DataFrame, pregunta: str) -> str:
    hoy        = pd.Timestamp.today().normalize()
    anio_hoy   = hoy.year
    mes_hoy    = hoy.month
    p          = pregunta.lower()

    meses_map = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
                 "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
    mes_ref  = next((v for k,v in meses_map.items() if k in p), mes_hoy)
    anio_ref = anio_hoy

    lineas = [
        f"Fecha de hoy: {hoy.strftime('%d/%m/%Y')}",
        f"Mes de referencia: {mes_ref}/{anio_ref}",
    ]

    # ── 1. CAPITAL ACTIVO ────────────────────────────────────────────────────
    try:
        cap_total = capital_activo_en_fecha(df_inv, hoy)
        lineas.append(f"\n=== CAPITAL ACTIVO HOY ===")
        lineas.append(f"  TOTAL FONDO: ${cap_total:,.2f}")
        # Por inversor
        df_c = df_inv.copy()
        df_c["fecha_inversion"]       = pd.to_datetime(df_c.get("fecha_inversion"), errors="coerce", dayfirst=True)
        df_c["fecha_final_inversion"] = pd.to_datetime(df_c.get("fecha_final_inversion"), errors="coerce", dayfirst=True)
        df_c["capital_invertido"]     = pd.to_numeric(df_c.get("capital_invertido"), errors="coerce").fillna(0)
        df_c["tipo_op"]               = df_c["tipo_operacion"].astype(str).str.strip().str.upper()
        activas = df_c[
            df_c["tipo_op"].isin(["NUEVA","REINVERSION"]) &
            df_c["fecha_inversion"].notna() &
            (df_c["fecha_inversion"] <= hoy) &
            (df_c["fecha_final_inversion"].isna() | (df_c["fecha_final_inversion"] >= hoy))
        ]
        por_inv = activas.groupby("inversor")["capital_invertido"].sum().sort_values(ascending=False)
        for inv, cap in por_inv.items():
            lineas.append(f"  {inv}: ${cap:,.2f}")
    except Exception as e:
        lineas.append(f"[Error capital: {e}]")

    # ── 2. INTERESES A INVERSORES DEL MES (lógica extractos) ────────────────
    try:
        lineas.append(f"\n=== INTERESES A PAGAR A INVERSORES {mes_ref}/{anio_ref} ===")
        df_e = df_inv.copy()
        for col in ["inversor","tipo_operacion","nombre_activo"]:
            if col in df_e.columns:
                df_e[col] = df_e[col].fillna("").astype(str).str.strip()
        df_e["tipo_op_n"] = df_e["tipo_operacion"].str.upper()
        df_e = df_e[df_e["tipo_op_n"].isin(["NUEVA","CANCELADA"])].copy()
        df_e["fecha_inversion"]       = pd.to_datetime(df_e.get("fecha_inversion"), errors="coerce", dayfirst=True)
        df_e["fecha_final_inversion"] = pd.to_datetime(df_e.get("fecha_final_inversion"), errors="coerce", dayfirst=True)
        df_e["capital_invertido"]     = pd.to_numeric(df_e.get("capital_invertido"), errors="coerce").fillna(0)
        df_e["interes_inversor_anual"]= pd.to_numeric(df_e.get("interes_inversor_anual"), errors="coerce").fillna(0)

        dias_mes   = ultimo_dia_mes(anio_ref, mes_ref)
        fecha_corte= pd.Timestamp(datetime(anio_ref, mes_ref, dias_mes))
        TRAMO_INV  = {"ROBERTO BISCAFE","CROWE BOLIVIA"}
        CORTE_T    = datetime(2026, 2, 1)
        FIN_T1     = datetime(2026, 1, 31)

        filas_int = []
        for _, row in df_e.iterrows():
            fi = row.get("fecha_inversion")
            if pd.isna(fi): continue
            fi_dt   = fi.to_pydatetime()
            tipo_op = row["tipo_op_n"]
            ff      = row.get("fecha_final_inversion")
            if tipo_op == "CANCELADA":
                if pd.isna(ff): continue
                fecha_fin_dt = min(ff.to_pydatetime(), fecha_corte.to_pydatetime())
            else:
                fecha_fin_dt = fecha_corte.to_pydatetime()
            inicio_mes_dt = datetime(anio_ref, mes_ref, 1)
            fin_mes_dt    = datetime(anio_ref, mes_ref, dias_mes)
            inicio_calc   = max(fi_dt, inicio_mes_dt)
            fin_calc      = min(fecha_fin_dt, fin_mes_dt)
            if inicio_calc > fin_calc: continue
            dias    = (fin_calc - inicio_calc).days + 1
            capital = float(row["capital_invertido"])
            tasa    = float(row["interes_inversor_anual"])
            inv_up  = str(row.get("inversor","")).strip().upper()
            if inv_up in TRAMO_INV:
                interes = 0.0
                if inicio_calc <= FIN_T1:
                    ft1 = min(fin_calc, FIN_T1)
                    interes += round((capital*0.05/12)*((ft1-inicio_calc).days+1)/dias_mes, 2)
                if fin_calc >= CORTE_T:
                    it2 = max(inicio_calc, CORTE_T)
                    interes += round((capital*0.075/12)*((fin_calc-it2).days+1)/dias_mes, 2)
            else:
                interes = round((capital*tasa/12)*dias/dias_mes, 2)
            filas_int.append({"inversor": str(row.get("inversor","")), "interes_mes": interes})

        if filas_int:
            df_int   = pd.DataFrame(filas_int)
            por_inv2 = df_int.groupby("inversor")["interes_mes"].sum().sort_values(ascending=False)
            for inv, val in por_inv2.items():
                lineas.append(f"  {inv}: ${val:,.2f}")
            lineas.append(f"  >> TOTAL A PAGAR: ${por_inv2.sum():,.2f}")
        else:
            lineas.append("  Sin datos.")
    except Exception as e:
        lineas.append(f"[Error intereses: {e}]")

    # ── 3. CALENDARIO DE NOTAS (próximos 180 días) ───────────────────────────
    try:
        lineas.append(f"\n=== CALENDARIO NOTAS PRÓXIMOS 180 DÍAS ===")
        if df_cal is not None and not df_cal.empty:
            limite = hoy + pd.Timedelta(days=180)
            df_c2  = df_cal.copy()
            df_c2["fecha"] = pd.to_datetime(df_c2["fecha"], errors="coerce")
            proximos = df_c2[(df_c2["fecha"] >= hoy) & (df_c2["fecha"] <= limite)].sort_values("fecha")

            if not proximos.empty:
                # Próximo PAGO con importe
                pagos_imp = proximos[
                    (proximos["tipo_evento"] == "PAGO") &
                    (pd.to_numeric(proximos.get("importe_cobro", pd.Series(dtype=float)), errors="coerce").fillna(0) > 0)
                ] if "importe_cobro" in proximos.columns else pd.DataFrame()

                if not pagos_imp.empty:
                    prox = pagos_imp.iloc[0]
                    lineas.append(f"  PRÓXIMO COBRO: {pd.Timestamp(prox['fecha']).strftime('%d/%m/%Y')} | Nota {prox.get('nota','')} | ${float(prox.get('importe_cobro',0)):,.2f}")
                    if len(pagos_imp) > 1:
                        sig = pagos_imp.iloc[1]
                        lineas.append(f"  SIGUIENTE: {pd.Timestamp(sig['fecha']).strftime('%d/%m/%Y')} | Nota {sig.get('nota','')} | ${float(sig.get('importe_cobro',0)):,.2f}")

                # Tabla completa próximos 60 días
                p60 = proximos[proximos["fecha"] <= hoy + pd.Timedelta(days=60)]
                cols_show = [c for c in ["fecha","tipo_evento","nota","importe_cobro","estado","detalle"] if c in p60.columns]
                if not p60.empty:
                    lineas.append(p60[cols_show].to_string(index=False))

                # Próximas observaciones por nota
                obs = proximos[proximos["tipo_evento"] == "OBSERVACION"].sort_values("fecha")
                if not obs.empty:
                    lineas.append("  Próximas observaciones:")
                    seen = set()
                    for _, r in obs.iterrows():
                        n = r.get("nota","")
                        if n not in seen:
                            lineas.append(f"    Nota {n}: {pd.Timestamp(r['fecha']).strftime('%d/%m/%Y')} | {r.get('estado','')}")
                            seen.add(n)

                # Calls
                calls = proximos[proximos["tipo_evento"] == "CALL"].sort_values("fecha")
                if not calls.empty:
                    lineas.append("  Próximos calls:")
                    for _, r in calls.iterrows():
                        lineas.append(f"    Nota {r.get('nota','')}: {pd.Timestamp(r['fecha']).strftime('%d/%m/%Y')}")
            else:
                lineas.append("  Sin eventos próximos.")
        else:
            lineas.append("  Sin calendario disponible.")
    except Exception as e:
        lineas.append(f"[Error calendario: {e}]")

    # ── 4. EXTRACTO ACUMULADO POR INVERSOR (loop mes a mes) ─────────────────
    try:
        # Detectar inversor mencionado
        inversores_conocidos = df_inv["inversor"].dropna().unique().tolist()
        inversor_preg = None
        p_up = p.upper()
        for inv in inversores_conocidos:
            if inv.upper() in p_up or any(pt in p_up for pt in inv.upper().split() if len(pt) > 3):
                inversor_preg = inv
                break

        mes_lim  = mes_ref
        anio_lim = anio_ref
        match_m  = re.search(r'(\d{1,2})[/\-](\d{4})', p)
        if match_m:
            try:
                mes_lim  = int(match_m.group(1))
                anio_lim = int(match_m.group(2))
            except: pass

        df_acum = df_inv.copy()
        for col in ["inversor","tipo_operacion"]:
            if col in df_acum.columns:
                df_acum[col] = df_acum[col].fillna("").astype(str).str.strip()
        df_acum["tipo_op_n"] = df_acum["tipo_operacion"].str.upper()
        df_acum = df_acum[df_acum["tipo_op_n"].isin(["NUEVA","CANCELADA"])].copy()
        df_acum["fecha_inversion"]       = pd.to_datetime(df_acum.get("fecha_inversion"), errors="coerce", dayfirst=True)
        df_acum["fecha_final_inversion"] = pd.to_datetime(df_acum.get("fecha_final_inversion"), errors="coerce", dayfirst=True)
        df_acum["capital_invertido"]     = pd.to_numeric(df_acum.get("capital_invertido"), errors="coerce").fillna(0)
        df_acum["interes_inversor_anual"]= pd.to_numeric(df_acum.get("interes_inversor_anual"), errors="coerce").fillna(0)

        if inversor_preg:
            df_acum = df_acum[df_acum["inversor"].str.upper() == inversor_preg.upper()].copy()

        fecha_ini_fondo = df_acum["fecha_inversion"].dropna().min()
        if pd.isna(fecha_ini_fondo):
            fecha_ini_fondo = datetime(2025, 9, 1)
        else:
            fecha_ini_fondo = fecha_ini_fondo.to_pydatetime()

        TRAMO_INV2 = {"ROBERTO BISCAFE","CROWE BOLIVIA"}
        CORTE_T2   = datetime(2026, 2, 1)
        FIN_T1_2   = datetime(2026, 1, 31)

        filas_acum = []
        ai, mi = fecha_ini_fondo.year, fecha_ini_fondo.month
        while (ai, mi) <= (anio_lim, mes_lim):
            dm = ultimo_dia_mes(ai, mi)
            fc = datetime(ai, mi, dm)
            im = datetime(ai, mi, 1)
            fm = datetime(ai, mi, dm)
            for _, row in df_acum.iterrows():
                fi = row.get("fecha_inversion")
                if pd.isna(fi): continue
                fi_dt   = fi.to_pydatetime()
                tipo_op = row["tipo_op_n"]
                ff      = row.get("fecha_final_inversion")
                if tipo_op == "CANCELADA":
                    if pd.isna(ff): continue
                    ffd = min(ff.to_pydatetime(), fc)
                else:
                    ffd = fc
                ic = max(fi_dt, im)
                fc2= min(ffd, fm)
                if ic > fc2: continue
                dias    = (fc2 - ic).days + 1
                capital = float(row["capital_invertido"])
                tasa    = float(row["interes_inversor_anual"])
                inv_up  = str(row.get("inversor","")).strip().upper()
                if inv_up in TRAMO_INV2:
                    interes = 0.0
                    if ic <= FIN_T1_2:
                        ft1 = min(fc2, FIN_T1_2)
                        interes += round((capital*0.05/12)*((ft1-ic).days+1)/dm, 2)
                    if fc2 >= CORTE_T2:
                        it2 = max(ic, CORTE_T2)
                        interes += round((capital*0.075/12)*((fc2-it2).days+1)/dm, 2)
                else:
                    interes = round((capital*tasa/12)*dias/dm, 2)
                filas_acum.append({"inversor": str(row.get("inversor","")), "mes": f"{mi:02d}/{ai}", "interes_mes": interes})
            mi = mi + 1 if mi < 12 else 1
            ai = ai if mi > 1 else ai + 1

        if filas_acum:
            df_ac = pd.DataFrame(filas_acum)
            if inversor_preg:
                lineas.append(f"\n=== EXTRACTO ACUMULADO {inversor_preg} hasta {mes_lim:02d}/{anio_lim} ===")
                for mes_k, int_k in df_ac.groupby("mes")["interes_mes"].sum().items():
                    lineas.append(f"  {mes_k}: ${float(int_k):,.2f}")
                lineas.append(f"  >> TOTAL ACUMULADO: ${float(df_ac['interes_mes'].sum()):,.2f}")
            else:
                lineas.append(f"\n=== INTERESES ACUMULADOS POR INVERSOR hasta {mes_lim:02d}/{anio_lim} ===")
                for inv, tot in df_ac.groupby("inversor")["interes_mes"].sum().sort_values(ascending=False).items():
                    lineas.append(f"  {inv}: ${float(tot):,.2f}")
                lineas.append(f"  >> GRAN TOTAL: ${float(df_ac['interes_mes'].sum()):,.2f}")
    except Exception as e:
        lineas.append(f"[Error extracto acumulado: {e}]")

    return "\n".join(lineas)


# ═══════════════════════════════════════════════════════════════════════════════
# SYSTEM PROMPT (igual que en app.py)
# ═══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """Eres el asistente financiero de Chaparro Fernández Wealth Management. Respondes con total precisión usando ÚNICAMENTE los datos del contexto que se te proporciona.

== LAS 3 FUENTES DE DATOS — USA SIEMPRE LA CORRECTA ==

FUENTE 1 — COBROS DE NOTAS Y FECHAS (sección CALENDARIO NOTAS del contexto):
  Usa esta fuente para: ¿cuánto cobraremos de notas este mes? ¿cuándo es el próximo cobro? ¿qué cobros hay entre fecha X e Y? ¿cuándo es el próximo call u observación?
  Los importes ya están calculados en el calendario. No los recalcules tú.

FUENTE 2 — INTERESES A INVERSORES (secciones INTERESES A PAGAR y EXTRACTO ACUMULADO del contexto):
  Usa esta fuente para: ¿cuánto cobra PAM este mes? ¿cuánto hemos pagado a JEP desde el inicio? ¿intereses acumulados de un inversor?

FUENTE 3 — CAPITAL ACTIVO (sección CAPITAL ACTIVO HOY del contexto):
  Usa esta fuente para: ¿cuánto capital tenemos activo? ¿cuánto tiene invertido cada inversor?

== REGLA DE ORO ==
NUNCA inventes ni calcules por tu cuenta si el dato ya está en el contexto.
Si el dato no está, dilo claramente.

== INVERSORES Y TASAS ==
LEO: 10% | JORDI CHAPARRO: 15% | YURI FERNANDEZ: 15%
ROBERTO BISCAFE: 5% hasta 31/01/2026, 7.5% desde 01/02/2026
CROWE BOLIVIA: 5% hasta 31/01/2026, 7.5% desde 01/02/2026
2012 JACC GROUP: 10% | PEDRO MAGAÑA: 10% | PAM: 10%
CHAPARRO FERNANDEZ: 0% — sociedad gestora
GOLDEN BRICKS: 10% | TERESA: 10% | JEP: 15%
JORDI ESPECIAL: 10% | EVA CHAPARRO: 15% | PAOLA CHAPARRO: 15% | JAPAN JORDI: 15%

== FORMATO WHATSAPP ==
Responde en texto plano sin markdown. Sin asteriscos, sin #, sin tablas.
Usa guiones simples para listas. Respuestas concisas y directas.
Fechas DD/MM/YYYY, importes con $ y 2 decimales."""


# ═══════════════════════════════════════════════════════════════════════════════
# LLAMADA A CLAUDE
# ═══════════════════════════════════════════════════════════════════════════════

def llamar_claude(numero: str, pregunta: str, contexto: str) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return "Error: API key no configurada."

    # Mantener historial por número (últimos 6 turnos)
    if numero not in historial_por_numero:
        historial_por_numero[numero] = []

    historial = historial_por_numero[numero]

    # Añadir mensaje del usuario con contexto
    historial.append({
        "role": "user",
        "content": f"DATOS DEL FONDO:\n\n{contexto[:18000]}\n\n---\nPREGUNTA: {pregunta}"
    })

    # Llamar a la API
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        },
        json={
            "model": "claude-sonnet-4-6",
            "max_tokens": 1000,
            "system": SYSTEM_PROMPT,
            "messages": historial[-8:],  # últimos 4 turnos
        },
        timeout=60,
    )

    data = resp.json()
    respuesta = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    if not respuesta:
        respuesta = f"Error: {data.get('error', {}).get('message', str(data))}"

    # Guardar respuesta en historial (sin el contexto pesado, solo el texto)
    historial.append({"role": "assistant", "content": respuesta})

    # Mantener solo los últimos 8 mensajes (4 turnos)
    historial_por_numero[numero] = historial[-8:]

    return respuesta


# ═══════════════════════════════════════════════════════════════════════════════
# WEBHOOK TWILIO
# ═══════════════════════════════════════════════════════════════════════════════

def verificar_twilio(request_obj) -> bool:
    """Verifica que el request viene realmente de Twilio."""
    auth_token = os.environ.get("TWILIO_AUTH_TOKEN", "")
    if not auth_token:
        return True  # Sin token configurado, permitir (solo para desarrollo)

    signature  = request_obj.headers.get("X-Twilio-Signature", "")
    url        = request_obj.url
    params     = request_obj.form

    from twilio.request_validator import RequestValidator
    validator = RequestValidator(auth_token)
    return validator.validate(url, params, signature)


@app.route("/webhook", methods=["POST"])
def webhook():
    # Verificar que viene de Twilio
    if not verificar_twilio(request):
        return Response("Unauthorized", status=403)

    numero   = request.form.get("From", "")
    mensaje  = request.form.get("Body", "").strip()

    # Verificar lista blanca
    if numero not in NUMEROS_AUTORIZADOS:
        twiml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>No tienes acceso a este servicio.</Message>
</Response>"""
        return Response(twiml, mimetype="text/xml")

    if not mensaje:
        return Response("<Response></Response>", mimetype="text/xml")

    # Cargar Excel y construir contexto
    try:
        df_inv, df_cal, df_control = cargar_excel()
        contexto = construir_contexto(df_inv, df_cal, df_control, mensaje)
    except Exception as e:
        contexto = f"Error cargando datos: {e}"

    # Llamar a Claude
    respuesta = llamar_claude(numero, mensaje, contexto)

    # Responder a WhatsApp via TwiML
    # Escapar caracteres especiales XML
    respuesta_xml = respuesta.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    twiml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Message>{respuesta_xml}</Message>
</Response>"""
    return Response(twiml, mimetype="text/xml")


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok", "servicio": "CF Wealth WhatsApp Bot"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
