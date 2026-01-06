from django.contrib import admin
from .models import Enrolamiento

@admin.register(Enrolamiento)
class EnrolamientoAdmin(admin.ModelAdmin):
    list_display = ('id_enrolamiento', 'rfc', 'nombre', 'paterno', 'puesto', 'activo')
    search_fields = ('rfc', 'nombre', 'paterno')