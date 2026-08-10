from django.db import migrations


def marcar_ajustes_previos(apps, schema_editor):
    """
    Bajo el esquema anterior el lienzo se guardaba UNICAMENTE cuando el
    operador habia ajustado algo en el modo edicion, asi que "tiene canvas"
    equivalia exactamente a `con_ajustes`.

    Al agregarse la columna, todas las filas existentes quedaron en False. Sin
    este respaldo, las credenciales que si se ajustaron a mano dejarian de
    restaurar sus ajustes al reimprimirlas -- volverian a salir con la
    plantilla limpia, perdiendo un trabajo que alguien ya hizo.

    De aqui en adelante la equivalencia deja de valer (el lienzo se guarda
    siempre), por eso esto corre una sola vez y no como regla en el codigo.
    """
    EnrolamientoCredencial = apps.get_model('credencializacion', 'EnrolamientoCredencial')
    EnrolamientoCredencial.objects.filter(
        canvas_frente__isnull=False
    ).update(con_ajustes=True)


def revertir(apps, schema_editor):
    # La columna entera se va con la migracion anterior; no hay nada que
    # deshacer aqui sin inventarse informacion.
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('credencializacion', '0021_enrolamientocredencial_con_ajustes'),
    ]

    operations = [
        migrations.RunPython(marcar_ajustes_previos, revertir),
    ]
