import re

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/familiar/familiar.component.html', 'r') as f:
    text = f.read()

# Replace images
text = re.sub(r'img/frontalNLFamiliar.jpg', 'img/frente_familiar.jpg', text)
text = re.sub(r'img/reversoNLFINAL.jpg', 'img/reverso_familiar.jpg', text)

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/familiar/familiar.component.html', 'w') as f:
    f.write(text)
