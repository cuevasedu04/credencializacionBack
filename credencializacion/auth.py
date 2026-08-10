"""
Autenticacion y permisos del sistema de control de accesos.

Todo se apoya en lo que Django ya trae -- `auth_user`, `auth_group`,
`auth_permission` -- sin inventar un esquema propio. Lo unico "nuevo" es:

- `BearerTokenAuthentication`: el frontend ya manda
  `Authorization: Bearer <token>` (ver TokenInterceptor), pero el
  `TokenAuthentication` de DRF por omision solo reconoce el prefijo
  `Token`. En vez de tocar el interceptor (usado en TODAS las peticiones),
  se adapta el backend para aceptar el prefijo que ya esta en produccion.
- `TienePermiso`: una clase de permiso parametrizable por codename, para
  exigir un permiso concreto del catalogo (`RecursoSistema.Meta.permissions`)
  en una vista o accion especifica.
- `EsSuperusuario`: gate para las pantallas de administracion.
"""
from rest_framework.authentication import TokenAuthentication
from rest_framework.permissions import BasePermission


class BearerTokenAuthentication(TokenAuthentication):
    keyword = 'Bearer'


class EsSuperusuario(BasePermission):
    """
    Solo `is_superuser`, el flag nativo de Django -- no una tabla de roles
    propia. Gatea la administracion de usuarios y roles: quien puede crear
    cuentas o repartir permisos debe ser alguien de maxima confianza, y
    Django ya resuelve ese caso sin necesitar codigo adicional.
    """
    def has_permission(self, request, view):
        return bool(request.user and request.user.is_authenticated and request.user.is_superuser)


def tiene_permiso(codename: str):
    """
    Fabrica una clase de permiso DRF que exige `codename` del catalogo
    `credencializacion.<codename>`.

    Los superusuarios siempre pasan (`user.has_perm` ya lo resuelve asi en
    Django: un superusuario tiene todos los permisos sin necesidad de que se
    los asignen uno por uno).
    """
    class _TienePermiso(BasePermission):
        def has_permission(self, request, view):
            usuario = request.user
            return bool(
                usuario and usuario.is_authenticated
                and usuario.has_perm(f'credencializacion.{codename}')
            )

    _TienePermiso.__name__ = f'TienePermiso_{codename}'
    return _TienePermiso
