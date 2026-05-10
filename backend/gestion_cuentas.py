import json
import os
from datetime import datetime, timedelta
import logging

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CUENTAS_FILE = os.path.join(BASE_DIR, "cuentas.json")
LIMITE_DIARIO = 450
logger = logging.getLogger("envio_masivos")

def cargar_cuentas():
    if not os.path.exists(CUENTAS_FILE):
        try:
            logger.warning(f"No se encontró el archivo de cuentas en {CUENTAS_FILE}")
        except Exception:
            pass
        return []
    try:
        with open(CUENTAS_FILE, "r", encoding="utf-8") as f:
            cuentas = json.load(f)
        if not isinstance(cuentas, list):
            cuentas = []
        return cuentas
    except Exception as e:
        try:
            logger.error(f"Error leyendo cuentas desde {CUENTAS_FILE}: {e}")
        except Exception:
            pass
        return []

def guardar_todas_las_cuentas(cuentas):
    try:
        with open(CUENTAS_FILE, "w", encoding="utf-8") as f:
            json.dump(cuentas, f, indent=4, ensure_ascii=False)
    except Exception as e:
        try:
            logger.error(f"Error guardando cuentas en {CUENTAS_FILE}: {e}")
        except Exception:
            pass

def agregar_o_actualizar_cuenta(datos):
    cuentas = cargar_cuentas()
    existente = next((c for c in cuentas if c["user"] == datos["user"]), None)
    
    now_iso = datetime.now().isoformat()
    
    nueva_cuenta = {
        "server": datos.get("server", "smtp.gmail.com"),
        "port": int(datos.get("port", 587)),
        "user": datos["user"],
        "pass": datos["pass"],
        "daily_count": existente["daily_count"] if existente else 0,
        "last_reset_date": existente["last_reset_date"] if existente else now_iso
    }
    
    if existente:
        cuentas = [c if c["user"] != datos["user"] else nueva_cuenta for c in cuentas]
    else:
        cuentas.append(nueva_cuenta)
    
    guardar_todas_las_cuentas(cuentas)

def eliminar_cuenta(email):
    cuentas = cargar_cuentas()
    cuentas = [c for c in cuentas if c["user"] != email]
    guardar_todas_las_cuentas(cuentas)

def obtener_cuenta_disponible():
    cuentas = cargar_cuentas()
    now = datetime.now()
    cambios = False
    cuenta_seleccionada = None
    
    for c in cuentas:
        last_reset_str = c.get("last_reset_date", "")
        try:
            last_reset = datetime.fromisoformat(last_reset_str)
        except ValueError:
            try:
                last_reset = datetime.strptime(last_reset_str, "%Y-%m-%d")
            except Exception:
                last_reset = datetime.min

        if (now - last_reset).total_seconds() >= 86460:
            c["daily_count"] = 0
            c["last_reset_date"] = now.isoformat()
            cambios = True
        
        if cuenta_seleccionada is None and c.get("daily_count", 0) < LIMITE_DIARIO:
            cuenta_seleccionada = c

    if cambios:
        guardar_todas_las_cuentas(cuentas)
    
    if not cuenta_seleccionada:
        try:
            logger.error(f"No hay cuenta SMTP disponible. Total cuentas: {len(cuentas)}")
        except Exception:
            pass
    
    return cuenta_seleccionada

def incrementar_contador(email):
    cuentas = cargar_cuentas()
    now = datetime.now()
    cambios = False
    
    for c in cuentas:
        if c["user"] == email:
            last_reset_str = c.get("last_reset_date", "")
            try:
                last_reset = datetime.fromisoformat(last_reset_str)
            except ValueError:
                try:
                    last_reset = datetime.strptime(last_reset_str, "%Y-%m-%d")
                except Exception:
                    last_reset = datetime.min
            
            if (now - last_reset).total_seconds() >= 86460:
                c["daily_count"] = 0
                c["last_reset_date"] = now.isoformat()
            
            c["daily_count"] += 1
            c["last_sent_date"] = now.isoformat()
            cambios = True
            break
            
    if cambios:
        guardar_todas_las_cuentas(cuentas)
