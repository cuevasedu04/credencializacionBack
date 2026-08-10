from django.contrib.auth.management import create_permissions
from django.db import migrations


CODENAMES_CATALOGO = [
    'ver_enrolamiento', 'ver_registro_empleado', 'ver_plantilla_anam',
    'ver_provisional', 'ver_familiar', 'ver_busqueda_avanzada',
    'ver_busqueda_enrolamiento_masivos', 'ver_enrolamiento_masivo',
    'ver_carga_masiva', 'ver_credencializacion', 'ver_reportes',
    'ver_plantillas', 'ver_imprimir_credenciales', 'ver_enrolamiento_previo',
    'ver_inventario_medios', 'ver_catalogo_areas', 'ver_auditoria_credenciales',
    'plantillas_administrar', 'medios_administrar', 'areas_administrar',
    'credenciales_editar_ajustes',
]

NOMBRE_GRUPO = 'Acceso heredado (migración)'


def crear_grupo_heredado(apps, schema_editor):
    """
    Hasta ahora CUALQUIER usuario autenticado veia TODO el menu:
    `rolesPermitidos` era `[1, 2, 3, 4, 9999]` en absolutamente todas las
    rutas salvo `/test` (ver CLAUDE.md, gotcha #4) -- osea que en la
    practica nunca hubo distincion real por rol.

    Si el nuevo sistema de permisos entrara en vigor sin mas, todo usuario
    NO superusuario perderia de golpe el acceso a todo el sistema, porque
    hoy nadie tiene ningun permiso asignado (la tabla de asignaciones nace
    vacia). Para que esto sea un cambio de infraestructura y no una
    regresion en produccion, se crea un rol con TODOS los permisos del
    catalogo y se asigna a cada usuario activo que no sea superusuario --
    reproduce exactamente el comportamiento de hoy. De aqui en adelante, el
    superusuario puede crear roles mas finos y mover gente a ellos desde
    /administracion; este rol de respaldo no se vuelve a tocar solo.
    """
    # Los permisos de RecursoSistema.Meta.permissions los crea el signal
    # post_migrate, que corre hasta el FINAL del comando `migrate` -- en este
    # punto de la corrida todavia no existen. Se fuerza su creacion aqui.
    # Patron documentado por Django para exactamente este caso:
    # https://docs.djangoproject.com/en/stable/topics/auth/default/#programmatically-creating-permissions
    app_config = apps.get_app_config('credencializacion')
    app_config.models_module = True
    create_permissions(app_config, apps=apps, verbosity=0)
    app_config.models_module = None

    Group = apps.get_model('auth', 'Group')
    Permission = apps.get_model('auth', 'Permission')
    User = apps.get_model('auth', 'User')

    permisos = Permission.objects.filter(
        content_type__app_label='credencializacion',
        content_type__model='recursosistema',
        codename__in=CODENAMES_CATALOGO,
    )
    if not permisos.exists():
        return

    grupo, _ = Group.objects.get_or_create(name=NOMBRE_GRUPO)
    grupo.permissions.set(permisos)

    for usuario in User.objects.filter(is_active=True, is_superuser=False):
        usuario.groups.add(grupo)


def revertir(apps, schema_editor):
    Group = apps.get_model('auth', 'Group')
    Group.objects.filter(name=NOMBRE_GRUPO).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('credencializacion', '0024_recursosistema'),
    ]

    operations = [
        migrations.RunPython(crear_grupo_heredado, revertir),
    ]
