"""
Configuración de usuarios y contraseñas para la app
"""
import bcrypt

# Diccionario con usuarios, contraseñas (hasheadas en bcrypt), y roles
USUARIOS_CREDENCIALES = {
    "Lpalacios": {
        "password_hash": "$2b$12$9DhF1sYv.H3EpQ1/IuB5LePBro3e8KZtJ9KzBZyng8SxjZln1m7f.",
        "rol": "Administrador",
        "nombre": "Luis Palacios"
    },
    "Nfernandez": {
        "password_hash": "$2b$12$p7znbgoP63/ffXnwoqOdzOqvIRZGn7G3E6NBjTAxfOjbQwQfrSURG",
        "rol": "Auditor",
        "nombre": "Nancy Fernandez"
    },
    "Gzambrano": {
        "password_hash": "$2b$12$FY.q0RpJwBebyYG5kv3cIegVpr4BBAu2kI9CWIPsiV7fid9eL1xWy",
        "rol": "Auditor",
        "nombre": "Gustavo Zambrano"
    },
    "Dguantay": {
        "password_hash": "$2b$12$TeCgIhdeTu0HLvZp0oemmuqU2zDQj30JbysihId4mahaOa1btJRdK",
        "rol": "Auditor",
        "nombre": "Diego Guantay"
    },
    "Jefeautosol": {
        "password_hash": "$2b$12$M4aYpKfgbpZ/jtxEyTghIOf9pgC0wkMZzhZuXrgsdj/8NX5k5Pjdm",
        "rol": "Jefe de Repuestos",
        "nombre": "Jefe Autosol"
    },
    "Jefeautolux": {
        "password_hash": "$2b$12$htFf8trH0PjFyu/je.wCK.0VRERZiNSS0XJDrVSOtZqxypIQ1ZWBe",
        "rol": "Jefe de Repuestos",
        "nombre": "Jefe Autolux"
    },
    "Jefeciel": {
        "password_hash": "$2b$12$PBRfsJIW3Oj3bMKVqbUpW.44NayWXG0CK6lWhvYX85mOGhRujI4h6",
        "rol": "Jefe de Repuestos",
        "nombre": "Jefe Ciel"
    }
}

# Credenciales en TEXTO PLANO para mostrar al usuario (eliminar después de primera vez)
CREDENCIALES_INICIALES = {
    "Lpalacios": "Lamp4201",
    "Nfernandez": "Fernandez123",
    "Gzambrano": "Zambrano123",
    "Dguantay": "Guantay123",
    "Jefeautosol": "Autosol123",
    "Jefeautolux": "Autolux123",
    "Jefeciel": "Ciel123"
}
