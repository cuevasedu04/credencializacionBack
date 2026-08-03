"""
URL configuration for sicre_ws project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/6.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.urls import path, include
from rest_framework.routers import DefaultRouter
from credencializacion.views import (
    EnrolamientoViewSet, SigViewSet, CustomLoginView, EnrolamientoFamiliarViewSet,
    foto_firma_empleado, CargaMasivaViewSet, EnrolamientoCredencialViewSet,
    PlantillaCredencialViewSet,
)


router = DefaultRouter()
# ruta: cada vez que alguien entre a /api/expedientes, lo atiende el ViewSet
router.register(r'expedientes', EnrolamientoViewSet)
router.register(r'expedientes-familiares', EnrolamientoFamiliarViewSet)
router.register(r'empleados-sig', SigViewSet)
router.register(r'carga-masiva', CargaMasivaViewSet, basename='carga-masiva')
# Nuevo esquema: enrolamientos con foto/firma en disco + plantillas tipo canvas.
router.register(r'enrolamiento-credencial', EnrolamientoCredencialViewSet, basename='enrolamiento-credencial')
router.register(r'plantillas-credencial', PlantillaCredencialViewSet, basename='plantillas-credencial')

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/', include(router.urls)),
    path('api/login/', CustomLoginView.as_view(), name='api_login'),
    path('api/foto-firma/<str:emplid>/', foto_firma_empleado, name='foto_firma_empleado'),
]

# Servir fotos, firmas y fondos de plantillas desde MEDIA_ROOT.
# En produccion esto lo debe atender nginx directamente sobre /media/.
urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
