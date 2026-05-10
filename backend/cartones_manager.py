import json
import os
import re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUTA_REGISTRO = os.path.join(BASE_DIR, "cartones_registro.json")


def _cargar():
    if not os.path.exists(RUTA_REGISTRO):
        return {}
    try:
        with open(RUTA_REGISTRO, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar(data):
    with open(RUTA_REGISTRO, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _clave_evento(drive_url):
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", drive_url)
    if m:
        return m.group(1)
    return drive_url.strip()


def obtener_cartones_usados(drive_url):
    clave = _clave_evento(drive_url)
    data = _cargar()
    evento = data.get(clave, {})
    return set(int(k) for k in evento.get("enviados", {}).keys())


def registrar_envios(drive_url, envios):
    clave = _clave_evento(drive_url)
    data = _cargar()
    if clave not in data:
        data[clave] = {"drive_url": drive_url, "enviados": {}}
    for e in envios:
        num = str(e["carton"])
        data[clave]["enviados"][num] = {
            "email": e["email"],
            "nombre": e.get("nombre", ""),
            "fecha": e.get("fecha", datetime.now().isoformat())
        }
    _guardar(data)


def obtener_reporte_evento(drive_url):
    clave = _clave_evento(drive_url)
    data = _cargar()
    evento = data.get(clave, {})
    result = []
    for num, info in evento.get("enviados", {}).items():
        result.append({
            "carton": int(num),
            "email": info.get("email", ""),
            "nombre": info.get("nombre", ""),
            "fecha": info.get("fecha", "")
        })
    result.sort(key=lambda x: x["carton"])
    return result


def extraer_numero_carton(nombre_archivo):
    base = os.path.splitext(nombre_archivo)[0]
    nums = re.findall(r'\d+', base)
    if nums:
        return int(nums[-1])
    return None
