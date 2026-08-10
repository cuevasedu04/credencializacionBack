from django.conf import settings
from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group, Permission
from rest_framework import serializers
from .models import (
    Enrolamiento, SicreTblSig, EnrolamientoFamiliar, CargaMasiva,
    EnrolamientoCredencial, PlantillaCredencial, UnidadAdministrativa,
    PerfilUsuario,
)
from . import media_utils
import base64
import binascii

Usuario = get_user_model()


def detectar_tipo_imagen(data):
    """
    Detecta el tipo de imagen leyendo los magic bytes (primeros bytes del archivo).
    """
    if not data or len(data) < 12:
        return 'jpeg'
    
    # PNG: 89 50 4E 47
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return 'png'
    # JPEG: FF D8 FF
    elif data[:3] == b'\xff\xd8\xff':
        return 'jpeg'
    # GIF: 47 49 46
    elif data[:3] == b'GIF':
        return 'gif'
    # WebP: 52 49 46 46 ... 57 45 42 50
    elif data[:4] == b'RIFF' and data[8:12] == b'WEBP':
        return 'webp'
    # BMP: 42 4D
    elif data[:2] == b'BM':
        return 'bmp'
    else:
        return 'jpeg'  # Default


class Base64BinaryField(serializers.Field):
    """
    Este campo personalizado recibe un string Base64 del Frontend,
    lo limpia (quita el encabezado 'data:image/...') y lo convierte a bytes
    para guardarlo en el BLOB de MySQL.
    """
    def to_representation(self, value):
        if not value:
            return None
        # Flag byte: foto/firma externalizada a safirho_db.NW_EMPL_FOTO_ANAM
        if len(value) <= 2:
            return '1'
        try:
            # Detectar el tipo de imagen
            image_type = detectar_tipo_imagen(value)
            
            # Convertir a base64 con el prefijo correcto
            base64_str = base64.b64encode(value).decode('utf-8')
            return f"data:image/{image_type};base64,{base64_str}"
        except Exception:
            return None

    def to_internal_value(self, data):
        if not data:
            return None
        
        try:
            if "base64," in data:
                data = data.split("base64,")[1]
            
            decoded = base64.b64decode(data)
            return decoded
        except (TypeError, binascii.Error):
            raise serializers.ValidationError("La imagen no tiene un formato Base64 válido.")

class EnrolamientoSerializer(serializers.ModelSerializer):
    foto = Base64BinaryField(
        required=False, 
        allow_null=True, 
        style={'base_template': 'textarea.html'}
    )
    firma = Base64BinaryField(
        required=False, 
        allow_null=True, 
        style={'base_template': 'textarea.html'}
    )

    class Meta:
        model = Enrolamiento
        fields = '__all__'

class EnrolamientoDataTableSerializer(serializers.ModelSerializer):
    foto = Base64BinaryField(
        required=False,
        allow_null=True,
        style={'base_template': 'textarea.html'}
    )
    firma = Base64BinaryField(
        required=False,
        allow_null=True,
        style={'base_template': 'textarea.html'}
    )

    class Meta:
        model = Enrolamiento
        fields = '__all__'


class EnrolamientoFamiliarSerializer(serializers.ModelSerializer):
    foto = Base64BinaryField(
        required=False,
        allow_null=True,
        style={'base_template': 'textarea.html'}
    )
    firma = Base64BinaryField(
        required=False,
        allow_null=True,
        style={'base_template': 'textarea.html'}
    )

    class Meta:
        model = EnrolamientoFamiliar
        fields = '__all__'


class SigSerializer(serializers.ModelSerializer):
    class Meta:
        model = SicreTblSig
        fields = '__all__'

class ArchivoExcelSerializer(serializers.Serializer):
    archivo = serializers.FileField()

class LoginSerializer(serializers.Serializer):
        email = serializers.CharField()  # acepta username o email
        password = serializers.CharField()
        idSistema = serializers.IntegerField(required = False)
class CargaMasivaSerializer(serializers.ModelSerializer):
    foto = Base64BinaryField(required=False, allow_null=True)
    firma = Base64BinaryField(required=False, allow_null=True)

    class Meta:
        model = CargaMasiva
        fields = '__all__'



# ==========================================================================
# NUEVO ESQUEMA: foto/firma en disco (MEDIA_ROOT) + plantillas tipo canvas
# ==========================================================================

class RutaMediaField(serializers.Field):
    """
    Campo generico para un CharField de ruta relativa dentro de MEDIA_ROOT.

    No esta enganchado a ningun modelo actualmente -- `EnrolamientoCredencial`
    ya no guarda foto/firma como columna (se resuelven por convencion via
    media_utils.resolver_foto/resolver_firma, ver EnrolamientoCredencialSerializer
    mas abajo). Se deja disponible por si se necesita un campo de archivo con
    esta misma logica de guardado (ej. una futura columna de tipo archivo en
    PlantillaCredencial u otro modelo).

    Lectura  -> URL publica servible por el navegador ('/media/fotos/123.jpg').
    Escritura-> acepta un data-URI base64 (lo guarda en disco y almacena la ruta),
                una ruta relativa ya existente, o null para limpiar el campo.

    El nombre del archivo se deriva del num_empleado, por lo que el guardado es
    deterministico y sobrescribe la version previa del mismo empleado.
    """

    def __init__(self, carpeta, extensiones, **kwargs):
        self.carpeta = carpeta
        self.extensiones = extensiones
        super().__init__(**kwargs)

    def to_representation(self, value):
        if not value:
            return None
        return media_utils.url_publica(value)

    def to_internal_value(self, data):
        if data in (None, '', 'null'):
            return None

        if not isinstance(data, str):
            raise serializers.ValidationError('Se esperaba una cadena base64 o una ruta.')

        # Ya es una ruta relativa o una URL /media/... que solo hay que normalizar.
        if not media_utils.es_data_uri(data):
            ruta = data.replace(settings.MEDIA_URL, '', 1) if data.startswith(settings.MEDIA_URL) else data
            return ruta.lstrip('/')

        identificador = self._identificador()
        if not identificador:
            raise serializers.ValidationError(
                'Se requiere num_empleado para poder guardar la imagen en disco.'
            )

        try:
            return media_utils.guardar_imagen(data, identificador, self.carpeta, self.extensiones)
        except media_utils.MediaError as exc:
            raise serializers.ValidationError(str(exc)) from exc

    def _identificador(self):
        """num_empleado del payload entrante, o el del registro que se esta editando."""
        datos = getattr(self.parent, 'initial_data', None) or {}
        identificador = datos.get('num_empleado')

        if not identificador:
            instancia = getattr(self.parent, 'instance', None)
            identificador = getattr(instancia, 'num_empleado', None)

        return identificador


class EnrolamientoCredencialSerializer(serializers.ModelSerializer):
    """
    `EnrolamientoCredencial` ya no guarda foto/firma como columna (ver
    docstring del modelo) -- se resuelven en el momento por convencion de
    nombre de archivo (propiedades foto_url/firma_url del modelo). Para
    capturar/guardar una foto o firma nueva se usa el action
    `guardar-medios` del viewset, que escribe directo a MEDIA_ROOT via
    num_empleado; este serializer no participa en esa escritura.

    Se exponen bajo las claves `foto`/`firma` (no `foto_url`/`firma_url`) a
    propósito: el editor de plantillas ya consume empleados con esas claves
    (plantilla-editor.const.ts, campo: 'foto'/'firma').
    """
    foto = serializers.CharField(source='foto_url', read_only=True)
    firma = serializers.CharField(source='firma_url', read_only=True)

    class Meta:
        model = EnrolamientoCredencial
        # Los lienzos quedan FUERA a proposito. El viewset los difiere para no
        # reventar el buffer de ordenamiento de MySQL (pesan ~47 KB por fila);
        # si el serializer los pidiera, Django los cargaria uno por uno al
        # serializar y cada fila del listado costaria una consulta extra a la
        # BD remota. Quien necesite el lienzo usa `auditoria-detalle`.
        exclude = ('canvas_frente', 'canvas_reverso')


class PlantillaCredencialSerializer(serializers.ModelSerializer):
    fondo_frente_url = serializers.CharField(read_only=True)
    fondo_reverso_url = serializers.CharField(read_only=True)

    class Meta:
        model = PlantillaCredencial
        fields = '__all__'

    def validate_clave(self, value):
        valor = (value or '').strip().upper().replace(' ', '_')
        if not valor:
            raise serializers.ValidationError('La clave de la plantilla es obligatoria.')
        return valor


class UnidadAdministrativaSerializer(serializers.ModelSerializer):
    """
    Catalogo de unidades administrativas.

    `nombre_normalizado` se expone solo de lectura: lo calcula el modelo al
    guardar a partir de `nombre`, y es la clave con la que se cruza contra el
    campo `area` del roster. Dejarlo editable permitiria romper el cruce sin
    que se note hasta que una credencial saliera con el area en blanco.
    """
    total_empleados = serializers.SerializerMethodField()

    class Meta:
        model = UnidadAdministrativa
        fields = [
            'id_unidad', 'nombre', 'nombre_normalizado', 'nombre_compactado',
            'activo', 'total_empleados', 'fecha_registro', 'fecha_modificacion',
        ]
        read_only_fields = ['nombre_normalizado', 'fecha_registro', 'fecha_modificacion']

    def get_total_empleados(self, obj):
        # Lo inyecta la vista para no disparar una consulta por fila.
        return getattr(obj, 'total_empleados', None)


# ==========================================================================
# Administracion: usuarios y roles (auth_user / auth_group / auth_permission,
# ninguna tabla propia -- ver credencializacion/auth.py)
# ==========================================================================

class PermisoSerializer(serializers.ModelSerializer):
    """
    Un permiso del catalogo (`RecursoSistema.Meta.permissions`). Solo lectura:
    el catalogo se define en codigo, no se edita desde la pantalla -- lo que
    SI se edita ahi es que roles lo tienen asignado.
    """
    class Meta:
        model = Permission
        fields = ['id', 'codename', 'name']


class RolSerializer(serializers.ModelSerializer):
    """
    Un rol es un `auth_group` con un subconjunto de permisos del catalogo.

    `permisos` acepta y regresa CODENAMES (`'ver_plantillas'`), no ids: es lo
    que ya usa el frontend para decidir que mostrar, asi que evita una vuelta
    extra resolviendo ids <-> nombres en cada pantalla.
    """
    permisos = serializers.ListField(child=serializers.CharField(), write_only=True, required=False)
    permisos_detalle = PermisoSerializer(source='permissions', many=True, read_only=True)
    total_usuarios = serializers.SerializerMethodField()

    class Meta:
        model = Group
        fields = ['id', 'name', 'permisos', 'permisos_detalle', 'total_usuarios']

    def get_total_usuarios(self, obj):
        return getattr(obj, 'total_usuarios', None)

    def _permisos_del_catalogo(self, codenames):
        return Permission.objects.filter(
            content_type__app_label='credencializacion',
            content_type__model='recursosistema',
            codename__in=codenames or [],
        )

    def create(self, validated_data):
        codenames = validated_data.pop('permisos', [])
        grupo = Group.objects.create(**validated_data)
        grupo.permissions.set(self._permisos_del_catalogo(codenames))
        return grupo

    def update(self, instance, validated_data):
        codenames = validated_data.pop('permisos', None)
        instance.name = validated_data.get('name', instance.name)
        instance.save()
        # None = el cliente no mando el campo (p.ej. solo renombrando el rol);
        # [] SI debe vaciar los permisos -- son casos distintos.
        if codenames is not None:
            instance.permissions.set(self._permisos_del_catalogo(codenames))
        return instance


class UsuarioSerializer(serializers.ModelSerializer):
    """
    CRUD de `auth_user` para la pantalla de administracion.

    La contrasena NUNCA se expone ni se acepta aqui -- tiene su propio action
    (`set-password` en `UsuarioViewSet`) para que quede claro en la API que
    cambiarla es una operacion distinta, deliberada, no un campo mas de un
    formulario de edicion general.
    """
    roles = serializers.PrimaryKeyRelatedField(source='groups', many=True, queryset=Group.objects.all(), required=False)
    roles_detalle = serializers.SerializerMethodField()
    # num_empleado/foto no viven en auth_user (ver PerfilUsuario) -- se leen
    # con SerializerMethodField y se escriben "a mano" desde initial_data en
    # create()/update(), mismo patron que ya usa `password` aqui abajo: no
    # son columnas del modelo que este serializer expone, asi que declararlos
    # como campos normales (con `source=`) no los haria escribibles solos.
    num_empleado = serializers.SerializerMethodField()
    foto = serializers.SerializerMethodField()

    class Meta:
        model = Usuario
        fields = [
            'id', 'username', 'email', 'first_name', 'last_name',
            'is_active', 'is_superuser', 'is_staff', 'date_joined',
            'last_login', 'roles', 'roles_detalle', 'num_empleado', 'foto',
        ]
        read_only_fields = ['date_joined', 'last_login']

    def get_roles_detalle(self, obj):
        return [{'id': g.id, 'name': g.name} for g in obj.groups.all()]

    def get_num_empleado(self, obj):
        # getattr con default funciona aunque no exista el PerfilUsuario:
        # Django hace que RelatedObjectDoesNotExist herede de AttributeError
        # exactamente para que este patron sea seguro.
        return getattr(getattr(obj, 'perfil', None), 'num_empleado', '') or ''

    def get_foto(self, obj):
        num_empleado = self.get_num_empleado(obj)
        if not num_empleado:
            return None
        ruta = media_utils.resolver_foto(num_empleado)
        return media_utils.url_publica(ruta) if ruta else None

    def _guardar_num_empleado(self, usuario):
        if 'num_empleado' not in self.initial_data:
            return
        num_empleado = (self.initial_data.get('num_empleado') or '').strip()
        PerfilUsuario.objects.update_or_create(
            usuario=usuario, defaults={'num_empleado': num_empleado},
        )

    def create(self, validated_data):
        roles = validated_data.pop('groups', [])
        password = self.initial_data.get('password')
        usuario = Usuario(**validated_data)
        # Sin contrasena usable: nadie podria iniciar sesion con esta cuenta
        # hasta que el superusuario le fije una desde `set-password`. Mejor
        # eso que aceptar un valor implicito o vacio como contrasena real.
        usuario.set_password(password or Usuario.objects.make_random_password())
        usuario.save()
        usuario.groups.set(roles)
        self._guardar_num_empleado(usuario)
        return usuario

    def update(self, instance, validated_data):
        usuario = super().update(instance, validated_data)
        self._guardar_num_empleado(usuario)
        return usuario
