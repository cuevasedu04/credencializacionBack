"""
Puebla `sicre_tbl_enrolamiento_credencial` a partir de la tabla historica
`sicre_tbl_enrolamiento`, vinculando foto y firma con los archivos que ya viven
en MEDIA_ROOT (nombrados por num_empleado).

Uso:
    python manage.py migrar_enrolamiento_credencial --dry-run
    python manage.py migrar_enrolamiento_credencial
    python manage.py migrar_enrolamiento_credencial --solo-medios
"""
from django.core.management.base import BaseCommand
from django.db import transaction

from credencializacion import media_utils
from credencializacion.models import Enrolamiento, EnrolamientoCredencial

# Campos que se copian tal cual del modelo historico al nuevo.
CAMPOS_COPIABLES = [
    'num_empleado', 'rfc', 'curp', 'nombre', 'paterno', 'materno', 'apellidos',
    'puesto', 'adscripcion', 'inicio_vig', 'fin_vig', 'folio', 'fecha_expedicion',
    'provisional', 'impreso', 'layout_credencial', 'id_usuario_registra',
    'id_usuario_modifica', 'fecha_modificacion', 'activo',
]


class Command(BaseCommand):
    help = 'Migra enrolamientos historicos a la tabla nueva vinculando fotos y firmas de MEDIA_ROOT.'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Solo reporta lo que haria, sin escribir en la base de datos.',
        )
        parser.add_argument(
            '--solo-medios',
            action='store_true',
            help='No crea registros nuevos; solo revincula foto/firma de los ya existentes.',
        )
        parser.add_argument(
            '--limite',
            type=int,
            default=0,
            help='Procesa a lo mucho N registros (0 = todos).',
        )

    def handle(self, *args, **opciones):
        dry_run = opciones['dry_run']
        solo_medios = opciones['solo_medios']
        limite = opciones['limite']

        if dry_run:
            self.stdout.write(self.style.WARNING('MODO DRY-RUN: no se escribira nada.'))

        if solo_medios:
            self._revincular_medios(dry_run, limite)
            return

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

            # La clave natural es num_empleado; sin ella no podemos resolver medios.
            if not num_empleado or num_empleado in existentes:
                omitidos += 1
                continue

            datos = {campo: getattr(registro, campo) for campo in CAMPOS_COPIABLES}
            datos['num_empleado'] = num_empleado
            datos['rfc'] = registro.rfc or num_empleado

            ruta_foto = media_utils.resolver_foto(num_empleado)
            ruta_firma = media_utils.resolver_firma(num_empleado)
            datos['foto'] = ruta_foto
            datos['firma'] = ruta_firma

            if ruta_foto:
                con_foto += 1
            if ruta_firma:
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

    # ------------------------------------------------------------------

    def _revincular_medios(self, dry_run, limite):
        qs = EnrolamientoCredencial.objects.filter(
            models_q_sin_medios()
        ).order_by('id_enrolamiento')

        if limite:
            qs = qs[:limite]

        actualizados = 0
        revisados = 0

        for registro in qs.iterator(chunk_size=500):
            revisados += 1
            if registro.resolver_archivo_existente():
                actualizados += 1
                if not dry_run:
                    registro.save(update_fields=['foto', 'firma'])

        self.stdout.write(f'Revisados:   {revisados}')
        self.stdout.write(
            self.style.SUCCESS(f'Revinculados: {actualizados}')
            if not dry_run
            else self.style.WARNING(f'Se revincularian: {actualizados}')
        )


def models_q_sin_medios():
    """Registros a los que les falta foto o firma."""
    from django.db.models import Q
    return Q(foto__isnull=True) | Q(foto='') | Q(firma__isnull=True) | Q(firma='')
