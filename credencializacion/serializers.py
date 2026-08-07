from django.conf import settings
from rest_framework import serializers
from .models import (
    Enrolamiento, SicreTblSig, EnrolamientoFamiliar, CargaMasiva,
    EnrolamientoCredencial, PlantillaCredencial, UnidadAdministrativa,
)
from . import media_utils
import base64
import binascii


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
        fields = '__all__'


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
