import json
import hashlib

ruta_json = r'C:\Users\user\Pictures\BuscadorPDFs\AppBuscador\Users\usuarios.json'

with open(ruta_json, 'r', encoding='utf-8') as f:
    usuarios = json.load(f)

for u in usuarios:
    usuario_plano = u['usuario']
    password_plano = u['contrasena']
    u['usuario'] = hashlib.sha256(usuario_plano.encode('utf-8')).hexdigest()
    u['contrasena'] = hashlib.sha256(password_plano.encode('utf-8')).hexdigest()

with open(ruta_json, 'w', encoding='utf-8') as f:
    json.dump(usuarios, f, indent=2, ensure_ascii=False)

print("Archivo JSON actualizado con usuarios y contraseñas hasheadas.")
