"""
Puebla `sicre_tbl_enrolamiento_credencial` a partir de la tabla historica
`sicre_tbl_enrolamiento`, copiando solo los campos de sistema (folio,
vigencias, plantilla, estatus de impresion) -- los datos del empleado
(nombre, curp, puesto, etc.) ya NO se copian aqui: viven en sicre_tbl_sig y
se consultan cruzando por num_empleado. Foto/firma tampoco se copian: se
resuelven en el momento por convencion de archivo en MEDIA_ROOT, nunca se
persisten (ver docstring del modelo EnrolamientoCredencial).

Uso:
    python manage.py migrar_enrolamiento_credencial --dry-run
    python manage.py migrar_enrolamiento_credencial
"""
from django.core.management.base import BaseCommand
from django.db import transaction

from credencializacion import media_utils
from credencializacion.models import Enrolamiento, EnrolamientoCredencial

# Campos que se copian tal cual del modelo historico al nuevo (mismo nombre
# en ambos lados). `layout_credencial` -> `plantilla_credencial` se mapea
# aparte porque el nombre cambio.
CAMPOS_COPIABLES = [
    'num_empleado', 'rfc', 'inicio_vig', 'fin_vig', 'folio', 'fecha_expedicion',
    'provisional', 'impreso', 'id_usuario_registra', 'id_usuario_modifica',
    'fecha_modificacion', 'activo',
]


class Command(BaseCommand):
    help = 'Migra enrolamientos historicos (solo campos de sistema) a la tabla nueva.'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Solo reporta lo que haria, sin escribir en la base de datos.',
        )
        parser.add_argument(
            '--limite',
            type=int,
            default=0,
            help='Procesa a lo mucho N registros (0 = todos).',
        )

    def handle(self, *args, **opciones):
        dry_run = opciones['dry_run']
        limite = opciones['limite']

        if dry_run:
            self.stdout.write(self.style.WARNING('MODO DRY-RUN: no se escribira nada.'))

        self._migrar(dry_run, limite)

    # ------------------------------------------------------------------

    def _migrar(self, dry_run, limite):
        existentes = set(
            EnrolamientoCredencial.objects
            .exclude(num_empleado__isnull=True)
            .exclude(num_empleado='')
            .values_list('num_empleado', flat=True)
        )

        origen = Enrolamiento.objects.all().order_by('id_enrolamiento')
        if limite:
            origen = origen[:limite]

        total = origen.count()
        self.stdout.write(f'Registros en tabla historica: {total}')
        self.stdout.write(f'Ya presentes en tabla nueva: {len(existentes)}')

        nuevos = []
        omitidos = 0
        con_foto = 0
        con_firma = 0

        for registro in origen.iterator(chunk_size=500):
            num_empleado = (registro.num_empleado or '').strip()

            if not num_empleado or num_empleado in existentes:
                omitidos += 1
                continue

            datos = {campo: getattr(registro, campo) for campo in CAMPOS_COPIABLES}
            datos['num_empleado'] = num_empleado
            datos['rfc'] = registro.rfc or num_empleado
            datos['plantilla_credencial'] = registro.layout_credencial

            # Solo para reportar cobertura -- no se guarda en el modelo nuevo,
            # foto_url/firma_url se resuelven solas por num_empleado.
            if media_utils.resolver_foto(num_empleado):
                con_foto += 1
            if media_utils.resolver_firma(num_empleado):
                con_firma += 1

            nuevos.append(EnrolamientoCredencial(**datos))
            existentes.add(num_empleado)

        self.stdout.write('')
        self.stdout.write(f'A crear:            {len(nuevos)}')
        self.stdout.write(f'  con foto en disco:  {con_foto}')
        self.stdout.write(f'  con firma en disco: {con_firma}')
        self.stdout.write(f'Omitidos (sin num_empleado o ya existentes): {omitidos}')

        if dry_run or not nuevos:
            return

        with transaction.atomic():
            EnrolamientoCredencial.objects.bulk_create(nuevos, batch_size=500)

        self.stdout.write(self.style.SUCCESS(f'Creados {len(nuevos)} registros.'))
