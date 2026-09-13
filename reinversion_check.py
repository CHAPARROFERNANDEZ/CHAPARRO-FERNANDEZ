"""
Chequeo diario de reinversiones pendientes — pensado para ejecutarse como un servicio Cron
de Railway (no como parte de la app Streamlit), una vez al día.

Reutiliza EXACTAMENTE la misma lógica que el botón manual del panel "Reinversiones" dentro de
la app (ejecutar_chequeo_reinversiones, definida en app.py): calcula qué inversores tienen una
reinversión pendiente hoy y, si hay alguna nueva que no se avisó ya hoy, envía un único email a
Yuri. No manda nada a los inversores.

Requisitos para que funcione igual que en la app:
  - Debe vivir en el mismo directorio que app.py (mismo repo/deploy), para poder importarlo.
  - Necesita las mismas credenciales que usa la app: o bien un .streamlit/secrets.toml con
    [email] sender/password, o las variables de entorno SMTP_SENDER y SMTP_PASSWORD.
  - Necesita las mismas credenciales de Google Drive que usa la app (st.secrets['gcp_service_account']
    si vienen por ahí, o el mecanismo que ya tengáis) para poder leer/escribir el Excel.

Uso en Railway: crear un NUEVO servicio en el mismo proyecto (o reutilizar uno de tipo cron
existente), con:
  - Start command:  python reinversion_check.py
  - Cron schedule:  0 12 * * *   (12:00 UTC todos los días — ajustar a la hora que prefieras)
No hace falta que sea el mismo servicio que sirve la app web; puede ser un servicio ligero
aparte que comparta el mismo repo de GitHub.
"""
import sys

try:
    from app import ejecutar_chequeo_reinversiones
except Exception as e:
    print(f"[reinversion_check] No se pudo importar app.py: {e}", file=sys.stderr)
    sys.exit(1)

if __name__ == "__main__":
    try:
        pendientes = ejecutar_chequeo_reinversiones(enviar_email=True)
        disparadas = pendientes[pendientes["dispara"]] if not pendientes.empty else pendientes
        print(f"[reinversion_check] Chequeo completado. Pendientes disparadas hoy: {len(disparadas)}.")
    except Exception as e:
        print(f"[reinversion_check] Error durante el chequeo: {e}", file=sys.stderr)
        sys.exit(1)
