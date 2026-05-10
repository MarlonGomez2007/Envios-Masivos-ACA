import unittest

from backend.gestion_cuentas import cargar_cuentas, agregar_o_actualizar_cuenta, eliminar_cuenta
#se llaman las funciones que vamos a probar
class TestGestionCuentas(unittest.TestCase):
    #clasesita con unnittest.testcase para que podamos usar todos los metodos de self assert

    def test_agregar_cuenta(self):
        #CUANDO VAYA A DEFINIR OBLIGATORIO PONER TEST AL INICIO DEL NOMBRE DE LA FUNCION PARA QUE SE EJECUTE AUTOMATICAMENTE Y CON GUION BAJO, con punto se traba, todo rarito
        agregar_o_actualizar_cuenta({"user": "test@gmail.com", "pass": "1234"})
        cuentas = cargar_cuentas()
        usuarios = [c["user"] for c in cuentas]
        self.assertIn("test@gmail.com", usuarios)
        #existen varios metodos de assert, como assertin, assertnotin, assertequal y pues esos son los mas usados

    def test_eliminar_cuenta(self):
        agregar_o_actualizar_cuenta({"user": "borrar@gmail.com", "pass": "abc"})
        eliminar_cuenta("borrar@gmail.com")
        cuentas = cargar_cuentas()
        usuarios = [c["user"] for c in cuentas]
        self.assertNotIn("borrar@gmail.com", usuarios)

    def test_cargar_cuentas_retorna_lista(self):
        #este test verifica que cargar_cuentas siempre retorne una lista, incluso si el archivo esta vacio o no existe
        #es importante porque si retornara None o algo raro, todo el modulo de envios se romperia feo
        resultado = cargar_cuentas()
        self.assertIsInstance(resultado, list)
        #assertIsInstance verifica que el resultado sea del tipo que le digamos, en este caso list
        #si retorna un dict, un None, o cualquier otra cosa, el test falla

if __name__ == "__main__":
    unittest.main()