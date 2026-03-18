import re

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/familiar/familiar.component.html', 'r') as f:
    html = f.read()

# Remove the Firma Digital item
html = re.sub(
    r'<div class="list-group-item d-flex justify-content-between align-items-center">\s*<span>Firma Digital</span>.*?</div>',
    '',
    html,
    flags=re.DOTALL
)

# Remove the disabled condition for signature
html = html.replace(
    '[disabled]="!empleado.foto || !empleado.firma || guardando"',
    '[disabled]="!empleado.foto || guardando"'
)

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/familiar/familiar.component.html', 'w') as f:
    f.write(html)
