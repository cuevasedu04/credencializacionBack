with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/provisional/provisional.component.ts', 'r') as f:
    text = f.read()

new_payload = """
      if (this.esFamiliar) {
        payload.folio_familiares = this.empleado.folio;
        payload.vivienda = this.empleado.vivienda || null;
        payload.marca = this.empleado.marca || null;
        payload.color = this.empleado.color || null;
        payload.modelo = this.empleado.modelo || null;
        payload.placas = this.empleado.placas || null;
        payload.firma = null;
      }
"""

text = text.replace(
"""      if (this.esFamiliar) {
        payload.folio_familiares = this.empleado.folio;
      }""",
    new_payload
)

with open('/home/dev/credencializacion/CredencializacionFront/src/app/content/provisional/provisional.component.ts', 'w') as f:
    f.write(text)
