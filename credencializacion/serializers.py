from rest_framework import serializers
from .models import Enrolamiento, SicreTblSig
import base64
import binascii

class Base64BinaryField(serializers.Field):
    """
    Este campo personalizado recibe un string Base64 del Frontend,
    lo limpia (quita el encabezado 'data:image/...') y lo convierte a bytes
    para guardarlo en el BLOB de MySQL.
    """
    def to_representation(self, value):
        if not value:
            return None
        try:
            return base64.b64encode(value).decode('utf-8')
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
    class Meta:
        model = Enrolamiento
        fields = '__all__'


class SigSerializer(serializers.ModelSerializer):
    class Meta:
        model = SicreTblSig
        fields = '__all__'

class LoginSerializer(serializers.Serializer):
        email = serializers.EmailField()
        password = serializers.CharField()
        idSistema = serializers.IntegerField(required = False)