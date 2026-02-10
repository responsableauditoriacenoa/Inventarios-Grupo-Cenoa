"""
Configuración de usuarios y contraseñas para la app
"""
import bcrypt

# Diccionario con usuarios, contraseñas (hasheadas en bcrypt), y roles
USUARIOS_CREDENCIALES = {
    "diego_guantay": {
        "password_hash": "$2b$12$AJBD2HZ7croVlJikPmhxhewMTdKAU6ZjrfBgamAHX0rbHnUcVz4Aq",
        "rol": "Auditor",
        "nombre": "Diego Guantay"
    },
    "nancy_fernandez": {
        "password_hash": "$2b$12$PcQDPif08S3vGh2ndyS9reCPyHVJKXldBAcbKf4YoWMCAcN4dugG2",
        "rol": "Auditor",
        "nombre": "Nancy Fernandez"
    },
    "gustavo_zambrano": {
        "password_hash": "$2b$12$kMt84HgobeYfcU2FlENFue8BEnd5hRMX6m3sGuj.3ihiUqigfZF72",
        "rol": "Auditor",
        "nombre": "Gustavo Zambrano"
    },
    "admin": {
        "password_hash": "$2b$12$aMJ7dG7Vqd5wq3p9oUsTCu/ahd/QVO9WcW3TZ46Z/r1i3u6k2M2AS",
        "rol": "Auditor",
        "nombre": "Admin"
    },
    "jefe_repuestos": {
        "password_hash": "$2b$12$sqFCkPlt6rTRLuggPQ9GA.42FJ0kgb3wniyhTCYv5kCg/TXv3Z5x2",
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
