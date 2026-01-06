# enrolamiento/views.py
from rest_framework import viewsets, filters, status
from rest_framework.decorators import action
from rest_framework.response import Response
from .models import Enrolamiento, SicreTblSig
from .serializers import EnrolamientoSerializer, SigSerializer, EnrolamientoDataTableSerializer
from django.db.models import Q
from django.contrib.auth import authenticate
from rest_framework.authtoken.models import Token
from .serializers import LoginSerializer
from rest_framework.views import APIView


class EnrolamientoViewSet(viewsets.ModelViewSet):
    queryset = Enrolamiento.objects.all().order_by('-id_enrolamiento')
    serializer_class = EnrolamientoSerializer
    
    filter_backends = [filters.SearchFilter]
    search_fields = ['rfc', 'num_empleado', 'nombre', 'paterno']   

    @action(detail=False, methods=['get'], url_path='listos-imprimir')
    def listos_para_imprimir(self, request):
        """
        Muestra solo los registros COMPLETOS:
        - Tienen Foto
        - Tienen Firma
        - Tienen Número de Empleado
        """
        queryset = self.get_queryset().exclude(
            Q(foto__isnull=True) | Q(foto='') |
            Q(firma__isnull=True) | Q(firma='') |
            Q(num_empleado__isnull=True) | Q(num_empleado='')
        )
        
        page = self.paginate_queryset(queryset)
        if page is not None:
            serializer = self.get_serializer(page, many=True)
            return self.get_paginated_response(serializer.data)

        serializer = self.get_serializer(queryset, many=True)
        return Response(serializer.data)

    @action(detail=False, methods=['get'], url_path='datatableImprimir')
    def datatableImprimir(self, request):
        """
        Endpoint para alimentar la data table del front.
        Columnas: Num Empleado, RFC, CURP, nombres, adscripcion, puesto.
        Filtros: foto y firma no nulos.
        """
        queryset = self.get_queryset().exclude(
            Q(foto__isnull=True) | Q(foto=b'') |
            Q(firma__isnull=True) | Q(firma=b'') 
        )
        
        page = self.paginate_queryset(queryset)
        if page is not None:
            serializer = EnrolamientoDataTableSerializer(page, many=True)
            return self.get_paginated_response(serializer.data)

        serializer = EnrolamientoDataTableSerializer(queryset, many=True)
        return Response(serializer.data)
    
    @action(detail=False, methods=['get'])
    def pendientes(self, request):
        """
        Muestra solo los registros que faltan de foto O firma.
        Una vez que ambos campos tengan datos, el registro dejará de salir en esta lista.
        """
        # Filtramos donde foto es nula/vacía O firma es nula/vacía
        queryset = self.get_queryset().filter(
            Q(foto__isnull=True) | Q(foto=b'') | 
            Q(firma__isnull=True) | Q(firma=b'')
        )
        
        page = self.paginate_queryset(queryset)
        if page is not None:
            serializer = self.get_serializer(page, many=True)
            return self.get_paginated_response(serializer.data)

        serializer = self.get_serializer(queryset, many=True)
        return Response(serializer.data)

    @action(detail=False, methods=['get', 'post'])
    def sincronizar(self, request):
        """
        Copia los datos de SicreTblSig a Enrolamiento.
        """
        registros_sig = SicreTblSig.objects.all()
        creados = 0
        actualizados = 0

        for sig in registros_sig:
            obj, created = Enrolamiento.objects.update_or_create(
                rfc=sig.rfc,
                defaults={
                    'num_empleado': sig.num_empleado,
                    'curp': sig.curp,
                    'nombre': sig.nombre,
                    'paterno': sig.paterno,
                    'materno': sig.materno,
                    'apellidos': sig.apellidos,
                    'puesto': sig.puesto,
                    'adscripcion': sig.adscripcion,
                    'inicio_vig': sig.inicio_vig,
                    'fin_vig': sig.fin_vig,
                    'eladia': sig.eladia,
                    'foto': sig.foto,
                    'activo': 1
                }
            )
          
            if not created:
                obj.save() 

            if created:
                creados += 1
            else:
                actualizados += 1

            

        return Response({
            'status': 'Sincronización completada',
            'creados': creados,
            'actualizados': actualizados
        }, status=status.HTTP_200_OK)

class SigViewSet(viewsets.ReadOnlyModelViewSet):
    queryset = SicreTblSig.objects.all().order_by('nombre')
    serializer_class = SigSerializer
    
    # Configuración del Buscador
    filter_backends = [filters.SearchFilter]
    search_fields = ['rfc', 'nombres']



class CustomLoginView(APIView):
    def post(self, request):
        serializer = LoginSerializer(data=request.data)
        if serializer.is_valid():
            email = serializer.validated_data['email']
            password = serializer.validated_data['password']
            
      
            user = authenticate(request, username=email, password=password)

            if user is not None:
                if not user.is_active:
                    return Response({
                        'status': 403,
                        'message': 'Usuario inactivo'
                    }, status=status.HTTP_200_OK)

                token, created = Token.objects.get_or_create(user=user)

                return Response({
                    'status': 200,
                    'message': 'Login exitoso',
                    'model': {
                        'token': token.key,
                        'idUsuario': user.id,
                        'nombreCompleto': f"{user.first_name} {user.last_name}",
                        'email': user.email,
                        'unidadAdscripcion': getattr(user, 'adscripcion', ''), 
                        'area': 'Área del usuario' 
                    }
                }, status=status.HTTP_200_OK)
            else:
                return Response({
                    'status': 401,
                    'message': 'Credenciales incorrectas'
                }, status=status.HTTP_200_OK) 
        
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)