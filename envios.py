import os
from backend.servicio_envios import enviar_correos

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    ruta_excel = os.getenv("RUTA_EXCEL") or os.path.join(BASE_DIR, "base.xlsx")
    pausa = float(os.getenv("PAUSA") or "0.5")
    prueba = os.getenv("MODO_PRUEBA") == "1"
    enviar_correos(ruta_excel, pausa, prueba)
