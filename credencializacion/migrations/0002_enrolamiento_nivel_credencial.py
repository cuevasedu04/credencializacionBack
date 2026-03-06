from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('credencializacion', '0001_initial'),
    ]

    operations = [
        migrations.AddField(
            model_name='enrolamiento',
            name='nivel_credencial',
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
    ]
