with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/provisional/provisional.component.ts', 'r') as f:
    text = f.read()

text = text.replace(
    'if(!this.empleado.foto || !this.empleado.firma) {',
    'if(!this.empleado.foto || (!this.empleado.firma && !this.esFamiliar)) {'
)

text = text.replace(
    "this.utils.MuestrasToast(TipoToast.Warning, 'Falta capturar foto o firma');",
    "this.utils.MuestrasToast(TipoToast.Warning, this.esFamiliar ? 'Falta capturar foto' : 'Falta capturar foto o firma');"
)

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/provisional/provisional.component.ts', 'w') as f:
    f.write(text)
