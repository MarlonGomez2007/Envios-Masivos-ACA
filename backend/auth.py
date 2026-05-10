import json
import os
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin

RUTA_USUARIOS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "usuarios.json")

def cargar_datos():
    if not os.path.exists(RUTA_USUARIOS):
        datos_iniciales = {
            "usuarios": {
                "admin": {
                    "password": generate_password_hash("2o26Nova*."),
                    "role": "admin",
                    "created_at": datetime.now().isoformat(),
                    "last_login": None,
                    "login_count": 0,
                    "active": True
                }
            },
            "historial": []
        }
        guardar_datos(datos_iniciales)
        return datos_iniciales
    
    try:
        with open(RUTA_USUARIOS, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {"usuarios": {}, "historial": []}

def guardar_datos(datos):
    with open(RUTA_USUARIOS, 'w', encoding='utf-8') as f:
        json.dump(datos, f, indent=4)

class Usuario(UserMixin):
    def __init__(self, id, role="user", active=True):
        self.id = id
        self.role = role
        self.active = active

    @property
    def is_active(self):
        return self.active

    @property
    def is_admin(self):
        return self.role == "admin"

def obtener_usuario(username):
    datos = cargar_datos()
    user_data = datos["usuarios"].get(username)
    if user_data:
        return Usuario(username, user_data.get("role", "user"), user_data.get("active", True))
    return None

def verificar_credenciales(username, password):
    datos = cargar_datos()
    user_data = datos["usuarios"].get(username)
    if user_data and check_password_hash(user_data["password"], password):
        if not user_data.get("active", True):
            return None, "Usuario desactivado"
            
        user_data["last_login"] = datetime.now().isoformat()
        user_data["login_count"] = user_data.get("login_count", 0) + 1
        datos["usuarios"][username] = user_data
        guardar_datos(datos)
        return Usuario(username, user_data.get("role", "user"), user_data.get("active", True)), "Exito"
    return None, "Credenciales inválidas"

def cambiar_estado_usuario(username, nuevo_estado):
    datos = cargar_datos()
    if username in datos["usuarios"]:
        datos["usuarios"][username]["active"] = nuevo_estado
        guardar_datos(datos)
        return True
    return False

def crear_usuario(username, password, role="user"):
    datos = cargar_datos()
    if username in datos["usuarios"]:
        return False, "El usuario ya existe"
    
    datos["usuarios"][username] = {
        "password": generate_password_hash(password),
        "role": role,
        "created_at": datetime.now().isoformat(),
        "last_login": None,
        "login_count": 0,
        "active": True
    }
    guardar_datos(datos)
    return True, "Usuario creado exitosamente"

def eliminar_usuario(username):
    datos = cargar_datos()
    if username in datos["usuarios"]:
        del datos["usuarios"][username]
        guardar_datos(datos)
        return True
    return False

def cambiar_password(username, new_password):
    datos = cargar_datos()
    if username in datos["usuarios"]:
        datos["usuarios"][username]["password"] = generate_password_hash(new_password)
        guardar_datos(datos)
        return True
    return False

def registrar_accion(username, accion, detalles=None):
    datos = cargar_datos()
    evento = {
        "usuario": username,
        "accion": accion,
        "detalles": detalles or {},
        "timestamp": datetime.now().isoformat()
    }
    datos["historial"].insert(0, evento)
    if len(datos["historial"]) > 1000:
        datos["historial"] = datos["historial"][:1000]
    guardar_datos(datos)

def obtener_todos_usuarios():
    datos = cargar_datos()
    return datos["usuarios"]

def obtener_historial():
    datos = cargar_datos()
    return datos["historial"]
