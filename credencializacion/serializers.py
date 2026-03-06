from rest_framework import serializers
from .models import Enrolamiento, SicreTblSig, EnrolamientoFamiliar
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
        fields = [
            'id_enrolamiento',
            'num_empleado',
            'rfc',
            'curp',
            'nombre',
            'paterno',
            'materno',
            'apellidos',
            'puesto',
            'adscripcion',
            'inicio_vig',
            'fin_vig',
            'eladia',
            'foto',
            'firma',
            'folio_familiares',
            'impreso',
            'fecha_expedicion',
            'id_usuario_registra',
            'fecha_registro',
            'fecha_enrolamiento',
            'id_usuario_modifica',
            'fecha_modificacion',
            'ultima_carga',
            'activo',
        ]


class SigSerializer(serializers.ModelSerializer):
    class Meta:
        model = SicreTblSig
        fields = '__all__'

class ArchivoExcelSerializer(serializers.Serializer):
    archivo = serializers.FileField()

class LoginSerializer(serializers.Serializer):
        email = serializers.EmailField()
        password = serializers.CharField()
        idSistema = serializers.IntegerField(required = False)