import bcrypt

# Contraseñas en texto plano
passwords = {
    "diego_guantay": "DieguG123!",
    "nancy_fernandez": "NancyF123!",
    "gustavo_zambrano": "GustavoZ123!",
    "admin": "Admin123!",
    "jefe_repuestos": "JefeRep123!"
}

print("Generando hashes bcrypt...\n")
for user, password in passwords.items():
    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    print(f'"{user}": "{hashed}",')
