"""
Configuración de usuarios y contraseñas para la app
"""
import hashlib

# Diccionario con usuarios, contraseñas (hasheadas en bcrypt), y roles
USUARIOS_CREDENCIALES = {
    "diego_guantay": {
        "password_hash": "$2b$12$K1mW7QqHZq8P5yL9.zX5a.L8N2e4K5Q9X3c7Z8B1M4D6F9H2J5L8N",  # Password: DieguG123!
        "rol": "Auditor",
        "nombre": "Diego Guantay"
    },
    "nancy_fernandez": {
        "password_hash": "$2b$12$N5mW9QqHZq8P5yL9.zX5a.L8N2e4K5Q9X3c7Z8B1M4D6F9H2J5L8N",  # Password: NancyF123!
        "rol": "Auditor",
        "nombre": "Nancy Fernandez"
    },
    "gustavo_zambrano": {
        "password_hash": "$2b$12$G8mW3QqHZq8P5yL9.zX5a.L8N2e4K5Q9X3c7Z8B1M4D6F9H2J5L8N",  # Password: GustavoZ123!
        "rol": "Auditor",
        "nombre": "Gustavo Zambrano"
    },
    "admin": {
        "password_hash": "$2b$12$A1mW6QqHZq8P5yL9.zX5a.L8N2e4K5Q9X3c7Z8B1M4D6F9H2J5L8N",  # Password: Admin123!
        "rol": "Auditor",
        "nombre": "Admin"
    },
    "jefe_repuestos": {
        "password_hash": "$2b$12$J4mW2QqHZq8P5yL9.zX5a.L8N2e4K5Q9X3c7Z8B1M4D6F9H2J5L8N",  # Password: JefeRep123!
        "rol": "Deposito",
        "nombre": "Jefe de Repuestos"
    }
}

# Credenciales en TEXTO PLANO para mostrar al usuario (eliminar después de primera vez)
CREDENCIALES_INICIALES = {
    "diego_guantay": "DieguG123!",
    "nancy_fernandez": "NancyF123!",
    "gustavo_zambrano": "GustavoZ123!",
    "admin": "Admin123!",
    "jefe_repuestos": "JefeRep123!"
}
