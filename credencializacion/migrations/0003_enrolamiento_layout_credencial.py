from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('credencializacion', '0002_enrolamiento_nivel_credencial'),
    ]

    operations = [
        migrations.AddField(
            model_name='enrolamiento',
            name='layout_credencial',
            field=models.CharField(blank=True, max_length=30, null=True),
        ),
    ]
