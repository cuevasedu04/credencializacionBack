"""
Importa el catalogo de unidades administrativas y su nombre compactado.

Origen: el reporte "Plantilla_Empleados" del sistema de control de plazas,
que trae una fila POR EMPLEADO con las columnas 'Unidad Administrativa' y
'DG o Aduana compactada'.
"""
from collections import Counter, defaultdict

from django.core.management.base import BaseCommand, CommandError

from credencializacion.models import UnidadAdministrativa

FILAS_ENCABEZADO_REPORTE = 3   # membrete, fecha de generacion y renglon vacio

# Direcciones Generales: las unicas que imprimen su nombre corto propio. El
# resto (las aduanas) imprimen 'ANAM'.
#
# La lista va aqui y NO se deduce del Excel de origen a proposito: ese archivo
# trae una fila por empleado, y el personal de una Direccion General
# comisionado en una aduana aparece con el nombre de la aduana (AICM, AIFA,
# Veracruz...), de modo que deducirla metia aduanas en la lista de DG.
DIRECCIONES_GENERALES = [
    'ANAM', 'OIC', 'DGE', 'DGPEDA', 'DGOA', 'DGIA', 'DGAAAI',
    'DGMEIA', 'DGJA', 'DGR', 'DGTI', 'UAF', 'DGPA',
]


class Command(BaseCommand):
    help = 'Importa el catalogo de unidades administrativas desde el Excel de plantilla de empleados.'

    def add_arguments(self, parser):
        parser.add_argument('archivo', help='Ruta del .xlsx')
        parser.add_argument('--hoja', default='Plantilla_Empleados')
        parser.add_argument(
            '--simular', action='store_true',
            help='Muestra lo que haria sin escribir en la base de datos.',
        )
        parser.add_argument(
            '--forzar', action='store_true',
            help='Sobrescribe las unidades que ya existen. Por omision se '
                 'respetan, para no pisar lo editado desde la pantalla de Catalogos.',
        )

    def handle(self, *args, **opciones):
        try:
            import pandas as pd
        except ImportError as exc:
            raise CommandError('Se requiere pandas para leer el Excel.') from exc

        try:
            df = pd.read_excel(
                opciones['archivo'], sheet_name=opciones['hoja'],
                skiprows=FILAS_ENCABEZADO_REPORTE,
            )
        except Exception as exc:
            raise CommandError(f'No se pudo leer el Excel: {exc}') from exc

        if df.shape[1] < 2:
            raise CommandError('Se esperaban al menos 2 columnas: unidad y compactada.')

        df = df.iloc[:, :2]
        df.columns = ['unidad', 'compactada']

        # El archivo trae una fila por empleado, no por unidad: hay que
        # agrupar. Y una misma unidad puede aparecer con varios compactados
        # -- pasa cuando alguien de una Direccion General esta comisionado en
        # una aduana y el reporte anota la aduana. Se toma el valor MAS
        # FRECUENTE, que es el acronimo real de la unidad; en los datos reales
        # ese dominante va del 84% al 99%.
        conteos = defaultdict(Counter)
        nombres = {}

        for unidad, compactada in zip(df['unidad'], df['compactada']):
            unidad = str(unidad or '').strip()
            compactada = str(compactada or '').strip()
            if not unidad or unidad.lower() == 'nan':
                continue
            if not compactada or compactada.lower() == 'nan':
                continue

            clave = UnidadAdministrativa.normalizar(unidad)
            conteos[clave][compactada] += 1
            nombres.setdefault(clave, unidad)

        if not conteos:
            raise CommandError('El archivo no contiene filas utilizables.')

        creados = actualizados = sin_cambio = 0
        ambiguas = []

        for clave, contador in sorted(conteos.items()):
            compactado, veces = contador.most_common(1)[0]
            total = sum(contador.values())
            if len(contador) > 1:
                ambiguas.append((nombres[clave], compactado, veces, total, len(contador)))

            # Regla de negocio: solo las Direcciones Generales conservan su
            # acronimo; toda aduana imprime 'ANAM'. Se resuelve aqui y se
            # guarda ya listo, para que la tabla diga literalmente lo que se
            # va a imprimir y sea editable a mano despues.
            if compactado not in DIRECCIONES_GENERALES:
                compactado = UnidadAdministrativa.NOMBRE_GENERICO

            if opciones['simular']:
                continue

            registro = UnidadAdministrativa.objects.filter(nombre_normalizado=clave).first()
            if registro is None:
                UnidadAdministrativa.objects.create(
                    nombre=nombres[clave], nombre_compactado=compactado, activo=True,
                )
                creados += 1
            elif not opciones['forzar']:
                # Ya existe: se respeta lo que haya, que pudo editarse a mano.
                sin_cambio += 1
            elif registro.nombre_compactado != compactado or registro.nombre != nombres[clave]:
                registro.nombre = nombres[clave]
                registro.nombre_compactado = compactado
                registro.save()
                actualizados += 1
            else:
                sin_cambio += 1

        if ambiguas:
            self.stdout.write(self.style.WARNING(
                f'\n{len(ambiguas)} unidades traian mas de un compactado; se tomo el dominante:'
            ))
            for nombre, elegido, veces, total, distintos in ambiguas:
                pct = veces / total * 100
                self.stdout.write(
                    f'   {nombre[:58]:<58} -> {elegido:<12} '
                    f'{veces}/{total} ({pct:.0f}%), {distintos} valores'
                )

        if opciones['simular']:
            self.stdout.write(self.style.WARNING(
                f'\n[simulacion] {len(conteos)} unidades listas para importar. No se escribio nada.'
            ))
            return

        self.stdout.write(self.style.SUCCESS(
            f'\nCatalogo importado: {creados} creadas, {actualizados} actualizadas, '
            f'{sin_cambio} sin cambio. Total en BD: {UnidadAdministrativa.objects.count()}'
        ))
