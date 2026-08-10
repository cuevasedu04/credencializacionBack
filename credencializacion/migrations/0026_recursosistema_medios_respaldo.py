from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('credencializacion', '0025_grupo_heredado'),
    ]

    operations = [
        migrations.AlterModelOptions(
            name='recursosistema',
            options={
                'verbose_name': 'Recurso del sistema',
                'verbose_name_plural': 'Recursos del sistema (catálogo de permisos)',
                'db_table': 'sicre_cat_recurso_sistema',
                'managed': True,
                'permissions': [
                    ('ver_enrolamiento', 'Ver: Enrolamiento'),
                    ('ver_registro_empleado', 'Ver: Registro rápido de empleado'),
                    ('ver_plantilla_anam', 'Ver: Carga manual ANAM'),
                    ('ver_provisional', 'Ver: Carga manual Nuevo Laredo'),
                    ('ver_familiar', 'Ver: Familiares Nuevo Laredo'),
                    ('ver_busqueda_avanzada', 'Ver: Búsqueda avanzada'),
                    ('ver_busqueda_enrolamiento_masivos', 'Ver: Búsqueda de enrolamiento masivo'),
                    ('ver_enrolamiento_masivo', 'Ver: Enrolamiento masivo'),
                    ('ver_carga_masiva', 'Ver: Carga masiva de Excel'),
                    ('ver_credencializacion', 'Ver: Credencialización (legado)'),
                    ('ver_reportes', 'Ver: Reportes'),
                    ('ver_plantillas', 'Ver: Editor de plantillas'),
                    ('ver_imprimir_credenciales', 'Ver: Imprimir credenciales'),
                    ('ver_enrolamiento_previo', 'Ver: Enrolamiento previo'),
                    ('ver_inventario_medios', 'Ver: Inventario de medios'),
                    ('ver_catalogo_areas', 'Ver: Catálogo de áreas'),
                    ('ver_auditoria_credenciales', 'Ver: Auditoría de credenciales'),
                    ('plantillas_administrar', 'Plantillas: crear, editar, eliminar y marcar por defecto'),
                    ('medios_administrar', 'Inventario de medios: reemplazar, renombrar y eliminar'),
                    ('medios_respaldo', 'Administración: descargar respaldo (zip) de fotos y firmas'),
                    ('areas_administrar', 'Catálogo de áreas: crear, editar y eliminar'),
                    ('credenciales_editar_ajustes', 'Imprimir credenciales: usar el modo edición rápida'),
                ],
            },
        ),
    ]
