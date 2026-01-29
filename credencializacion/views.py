# enrolamiento/views.py
from rest_framework import viewsets, filters, status
from rest_framework.decorators import action
from rest_framework.response import Response
from .models import Enrolamiento, SicreTblSig
from .serializers import EnrolamientoSerializer, SigSerializer, EnrolamientoDataTableSerializer, ArchivoExcelSerializer, LoginSerializer
from django.db.models import Q
from django.contrib.auth import authenticate
from rest_framework.authtoken.models import Token
from rest_framework.views import APIView
import pandas as pd
import base64
import os
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO


def extraer_imagenes_de_excel(archivo, columna_foto='L'):
    """
    Extrae las imágenes incrustadas en Excel y las asocia con el número de fila.
    Retorna un diccionario: {numero_fila: bytes_imagen}
    """
    wb = load_workbook(archivo)
    ws = wb.active
    imagenes_por_fila = {}
    
    for image in ws._images:
        # Obtener la celda ancla donde está la imagen
        row = image.anchor._from.row + 1  # +1 porque openpyxl usa índice 0
        
        # Convertir la imagen a bytes
        if hasattr(image, '_data'):
            imagenes_por_fila[row] = image._data()
        elif hasattr(image, 'ref'):
            img_stream = BytesIO()
            image.ref.save(img_stream)
            imagenes_por_fila[row] = img_stream.getvalue()
    
    wb.close()
    return imagenes_por_fila


def procesar_foto_desde_excel(valor_celda, fila_numero=None, imagenes_excel=None):
    """
    Procesa el valor de la columna 'foto' del Excel y lo convierte a bytes.
    Prioridad:
    1. Imagen incrustada en Excel (si existe)
    2. Ruta de archivo
    3. Datos en base64
    4. None/vacío
    """
    # Prioridad 1: Imagen incrustada en el Excel
    if imagenes_excel and fila_numero and fila_numero in imagenes_excel:
        return imagenes_excel[fila_numero]
    
    if pd.isnull(valor_celda) or valor_celda == '':
        return None
    
    valor_str = str(valor_celda).strip()
    
    # Prioridad 2: Es una ruta de archivo (probar con y sin normalización)
    rutas_a_probar = [
        valor_str,
        os.path.abspath(valor_str),
        valor_str.replace('\\', '/'),
        valor_str.replace('/', '\\')
    ]
    
    for ruta in rutas_a_probar:
        if os.path.exists(ruta) and os.path.isfile(ruta):
            try:
                with open(ruta, 'rb') as f:
                    return f.read()
            except Exception as e:
                print(f"Error leyendo archivo {ruta}: {e}")
                continue
    
    # Prioridad 3: Es base64
    try:
        if 'base64,' in valor_str:
            valor_str = valor_str.split('base64,')[1]
        return base64.b64decode(valor_str)
    except Exception:
        return None


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
        Filtros: foto y firma no nulos, y que no estén marcados como impresos (impreso != 1).
        """
        queryset = self.get_queryset().exclude(
            Q(foto__isnull=True) | Q(foto=b'') |
            Q(firma__isnull=True) | Q(firma=b'') 
        ).filter(
            Q(impreso__isnull=True) | Q(impreso=0) | ~Q(impreso=1)
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
    
    @action(detail=False, methods=['get'], url_path='obtener-folio-maximo')
    def obtener_folio_maximo(self, request):
        """
        Retorna el folio máximo actual de la tabla enrolamiento.
        El front puede usar este valor para asignar el siguiente consecutivo.
        Respeta el formato con ceros a la izquierda (ej: 000001 -> 000002).
        """
        try:
            # Filtrar folios no nulos y obtener todos
            folios = Enrolamiento.objects.exclude(
                Q(folio__isnull=True) | Q(folio='')
            ).values_list('folio', flat=True)
            
            if not folios:
                # Si no hay folios, empezar desde 000001 (6 dígitos por defecto)
                return Response({
                    'status': 'success',
                    'folio_maximo': '000000',
                    'siguiente_folio': '000001'
                }, status=status.HTTP_200_OK)
            
            # Encontrar el folio con mayor valor numérico
            folio_max_valor = 0
            folio_max_str = ''
            longitud_formato = 6  # Default
            
            for folio in folios:
                folio_str = str(folio).strip()
                try:
                    valor_numerico = int(folio_str)
                    if valor_numerico > folio_max_valor:
                        folio_max_valor = valor_numerico
                        folio_max_str = folio_str
                        longitud_formato = len(folio_str)
                except (ValueError, TypeError):
                    continue
            
            if folio_max_valor == 0:
                # No se encontraron folios válidos
                return Response({
                    'status': 'success',
                    'folio_maximo': '000000',
                    'siguiente_folio': '000001'
                }, status=status.HTTP_200_OK)
            
            # Calcular siguiente folio manteniendo el formato
            siguiente_valor = folio_max_valor + 1
            siguiente_folio = str(siguiente_valor).zfill(longitud_formato)
            
            return Response({
                'status': 'success',
                'folio_maximo': folio_max_str,
                'siguiente_folio': siguiente_folio
            }, status=status.HTTP_200_OK)
            
        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al obtener folio máximo: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
    
    @action(detail=True, methods=['post'], url_path='marcar-impreso')
    def marcar_impreso(self, request, pk=None):
        """
        Marca un registro como impreso (impreso = 1) y actualiza la fecha de expedición.
        Se llama después de generar exitosamente el PDF de la credencial.
        """
        try:
            enrolamiento = self.get_object()
            
            # Marcar como impreso y actualizar fecha
            enrolamiento.impreso = 1
            enrolamiento.fecha_expedicion = request.data.get('fecha_expedicion')  # Opcional: enviar desde front
            enrolamiento.save()
            
            return Response({
                'status': 'success',
                'mensaje': f'Credencial {enrolamiento.folio or enrolamiento.id_enrolamiento} marcada como impresa',
                'data': {
                    'id_enrolamiento': enrolamiento.id_enrolamiento,
                    'folio': enrolamiento.folio,
                    'impreso': enrolamiento.impreso,
                    'fecha_expedicion': enrolamiento.fecha_expedicion
                }
            }, status=status.HTTP_200_OK)
            
        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al marcar como impreso: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
    
    @action(detail=False, methods=['get', 'post'], url_path='busqueda-avanzada')
    def busqueda_avanzada(self, request):
        """
        Endpoint de búsqueda avanzada para ag-grid.
        Permite filtrar por múltiples campos y retorna todos los detalles.
        Soporta filtros por: num_empleado, rfc, curp, nombre, paterno, materno,
        puesto, adscripcion, folio, impreso, fecha_expedicion_desde, fecha_expedicion_hasta.
        """
        queryset = self.get_queryset()
        
        # Obtener parámetros de filtro (soporta GET y POST)
        params = request.query_params if request.method == 'GET' else request.data
        
        # Filtros de texto
        if params.get('num_empleado'):
            queryset = queryset.filter(num_empleado__icontains=params['num_empleado'])
        
        if params.get('rfc'):
            queryset = queryset.filter(rfc__icontains=params['rfc'])
        
        if params.get('curp'):
            queryset = queryset.filter(curp__icontains=params['curp'])
        
        if params.get('nombre'):
            queryset = queryset.filter(nombre__icontains=params['nombre'])
        
        if params.get('paterno'):
            queryset = queryset.filter(paterno__icontains=params['paterno'])
        
        if params.get('materno'):
            queryset = queryset.filter(materno__icontains=params['materno'])
        
        if params.get('puesto'):
            queryset = queryset.filter(puesto__icontains=params['puesto'])
        
        if params.get('adscripcion'):
            queryset = queryset.filter(adscripcion__icontains=params['adscripcion'])
        
        if params.get('folio'):
            queryset = queryset.filter(folio__icontains=params['folio'])
        
        # Filtro por estado de impresión
        if params.get('impreso') is not None:
            impreso_val = params['impreso']
            if impreso_val == '1' or impreso_val == 1:
                queryset = queryset.filter(impreso=1)
            elif impreso_val == '0' or impreso_val == 0:
                queryset = queryset.filter(Q(impreso=0) | Q(impreso__isnull=True))
        
        # Filtros por rango de fechas de expedición
        if params.get('fecha_expedicion_desde'):
            queryset = queryset.filter(fecha_expedicion__gte=params['fecha_expedicion_desde'])
        
        if params.get('fecha_expedicion_hasta'):
            queryset = queryset.filter(fecha_expedicion__lte=params['fecha_expedicion_hasta'])
        
        # Filtros por rango de fechas de vigencia
        if params.get('inicio_vig_desde'):
            queryset = queryset.filter(inicio_vig__gte=params['inicio_vig_desde'])
        
        if params.get('fin_vig_hasta'):
            queryset = queryset.filter(fin_vig__lte=params['fin_vig_hasta'])
        
        # Filtros por fecha de registro
        if params.get('fecha_registro_desde'):
            queryset = queryset.filter(fecha_registro__date__gte=params['fecha_registro_desde'])
        
        if params.get('fecha_registro_hasta'):
            queryset = queryset.filter(fecha_registro__date__lte=params['fecha_registro_hasta'])
        
        # Filtro por fecha de registro específica
        if params.get('fecha_registro'):
            queryset = queryset.filter(fecha_registro__date=params['fecha_registro'])
        
        # Filtro por fecha de expedición específica
        if params.get('fecha_expedicion'):
            queryset = queryset.filter(fecha_expedicion=params['fecha_expedicion'])
        
        # Filtro por estado de completitud (con foto y firma)
        if params.get('solo_completos') == 'true' or params.get('solo_completos') == True:
            queryset = queryset.exclude(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            )
        
        # Filtro por registros activos
        if params.get('solo_activos') == 'true' or params.get('solo_activos') == True:
            queryset = queryset.filter(activo=1)
        
        # Paginación para ag-grid
        page = self.paginate_queryset(queryset)
        if page is not None:
            serializer = self.get_serializer(page, many=True)
            return self.get_paginated_response(serializer.data)

        serializer = self.get_serializer(queryset, many=True)
        return Response(serializer.data)
    
    @action(detail=False, methods=['get'], url_path='estadisticas')
    def estadisticas(self, request):
        """
        Retorna estadísticas generales del sistema de credencialización.
        Incluye: total de credenciales, impresas, pendientes, impresas hoy,
        por adscripción, y tendencias.
        Soporta filtrado por rango de fechas (fecha_desde, fecha_hasta).
        """
        from django.db.models import Count, Q
        from datetime import datetime, timedelta
        
        try:
            # Fecha de hoy
            hoy = datetime.now().date()
            
            # Obtener parámetros de filtrado por fechas
            fecha_desde = request.query_params.get('fecha_desde')
            fecha_hasta = request.query_params.get('fecha_hasta')
            
            # Crear queryset base con filtros opcionales
            queryset_base = Enrolamiento.objects.all()
            if fecha_desde:
                queryset_base = queryset_base.filter(fecha_registro__date__gte=fecha_desde)
            if fecha_hasta:
                queryset_base = queryset_base.filter(fecha_registro__date__lte=fecha_hasta)
            
            # Total de registros
            total_credenciales = queryset_base.count()
            
            # Credenciales completas (con foto y firma)
            credenciales_completas = queryset_base.exclude(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            ).count()
            
            # Credenciales impresas
            credenciales_impresas = queryset_base.filter(impreso=1).count()
            
            # Credenciales impresas hoy
            credenciales_impresas_hoy = queryset_base.filter(
                impreso=1,
                fecha_expedicion=hoy
            ).count()
            
            # Credenciales pendientes de impresión (completas pero no impresas)
            credenciales_pendientes = queryset_base.exclude(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            ).filter(
                Q(impreso__isnull=True) | Q(impreso=0) | ~Q(impreso=1)
            ).count()
            
            # Lista detallada de pendientes (para mostrar al dar click en el banner)
            lista_pendientes = queryset_base.exclude(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            ).filter(
                Q(impreso__isnull=True) | Q(impreso=0) | ~Q(impreso=1)
            ).values('id_enrolamiento', 'num_empleado', 'rfc', 'nombre', 'paterno', 'materno', 'folio', 'adscripcion')
            
            # Credenciales incompletas (sin foto o sin firma)
            credenciales_incompletas = queryset_base.filter(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            ).count()
            
            # Lista detallada de incompletas
            lista_incompletas = queryset_base.filter(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            ).values('id_enrolamiento', 'num_empleado', 'rfc', 'nombre', 'paterno', 'materno', 'adscripcion')
            
            # Credenciales por adscripción (top 10)
            por_adscripcion = queryset_base.values('adscripcion').annotate(
                total=Count('id_enrolamiento')
            ).order_by('-total')[:10]
            
            # Estadísticas de los últimos 7 días
            hace_7_dias = hoy - timedelta(days=7)
            credenciales_ultima_semana = queryset_base.filter(
                fecha_registro__gte=hace_7_dias
            ).count()
            
            impresas_ultima_semana = queryset_base.filter(
                impreso=1,
                fecha_expedicion__gte=hace_7_dias
            ).count()
            
            # Estadísticas del mes actual
            inicio_mes = hoy.replace(day=1)
            credenciales_mes_actual = queryset_base.filter(
                fecha_registro__gte=inicio_mes
            ).count()
            
            impresas_mes_actual = queryset_base.filter(
                impreso=1,
                fecha_expedicion__gte=inicio_mes
            ).count()
            
            # Porcentajes
            porcentaje_impresas = round((credenciales_impresas / total_credenciales * 100), 2) if total_credenciales > 0 else 0
            porcentaje_completas = round((credenciales_completas / total_credenciales * 100), 2) if total_credenciales > 0 else 0
            porcentaje_pendientes = round((credenciales_pendientes / credenciales_completas * 100), 2) if credenciales_completas > 0 else 0
            
            return Response({
                'status': 'success',
                'totales': {
                    'total_credenciales': total_credenciales,
                    'credenciales_completas': credenciales_completas,
                    'credenciales_impresas': credenciales_impresas,
                    'credenciales_pendientes': credenciales_pendientes,
                    'credenciales_incompletas': credenciales_incompletas
                },
                'hoy': {
                    'credenciales_impresas_hoy': credenciales_impresas_hoy
                },
                'ultima_semana': {
                    'credenciales_registradas': credenciales_ultima_semana,
                    'credenciales_impresas': impresas_ultima_semana
                },
                'mes_actual': {
                    'credenciales_registradas': credenciales_mes_actual,
                    'credenciales_impresas': impresas_mes_actual
                },
                'porcentajes': {
                    'porcentaje_impresas': porcentaje_impresas,
                    'porcentaje_completas': porcentaje_completas,
                    'porcentaje_pendientes': porcentaje_pendientes
                },
                'por_adscripcion': list(por_adscripcion),
                'detalle_pendientes': list(lista_pendientes),
                'detalle_incompletas': list(lista_incompletas)
            }, status=status.HTTP_200_OK)
            
        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al obtener estadísticas: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

class SigViewSet(viewsets.ReadOnlyModelViewSet):
    queryset = SicreTblSig.objects.all().order_by('nombre')
    serializer_class = SigSerializer
    
    # Configuración del Buscador
    filter_backends = [filters.SearchFilter]
    search_fields = ['rfc', 'nombres']

    # --- PASO 1: PREVISUALIZAR EXCEL ---
    @action(detail=False, methods=['POST'], serializer_class=ArchivoExcelSerializer)
    def previsualizar_excel(self, request):
        """
        Lee el Excel y retorna una vista previa de los datos sin guardar nada.
        """
        serializer = ArchivoExcelSerializer(data=request.data)
        if serializer.is_valid():
            archivo = serializer.validated_data['archivo']
            try:
                # Leer el Excel
                df = pd.read_excel(archivo)
                total_registros = len(df)
                
                # Extraer imágenes para saber qué filas tienen foto
                imagenes_excel = extraer_imagenes_de_excel(archivo)
                
                # Convertir a lista de diccionarios para el front
                registros_preview = []
                for index, row in df.iterrows():
                    fila_excel = index + 2
                    tiene_foto = fila_excel in imagenes_excel or pd.notnull(row.get('foto'))
                    
                    registro = {
                        'num_empleado': row.get('num_empleado'),
                        'rfc': row.get('rfc') or row.get('RFC'),
                        'curp': row.get('curp'),
                        'nombre': row.get('nombre'),
                        'paterno': row.get('paterno'),
                        'materno': row.get('materno'),
                        'puesto': row.get('puesto'),
                        'adscripcion': row.get('adscripcion'),
                        'inicio_vig': row.get('Inicio Vigencia') or row.get('inicio_vig'),
                        'fin_vig': row.get('Fin Vigencia') or row.get('fin_vig'),
                        'eladia': row.get('eladia'),
                        'foto': tiene_foto
                    }
                    registros_preview.append(registro)
                
                return Response({
                    "status": "success",
                    "mensaje": "Vista previa generada correctamente",
                    "total_registros": total_registros,
                    "registros": registros_preview
                }, status=status.HTTP_200_OK)
                
            except Exception as e:
                return Response({
                    "status": "error",
                    "mensaje": f"Error al leer el archivo: {str(e)}"
                }, status=status.HTTP_400_BAD_REQUEST)
        
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    # --- PASO 2: VALIDAR Y CARGAR (TODO EN UNO) ---
    @action(detail=False, methods=['POST'], serializer_class=ArchivoExcelSerializer)
    def subir_excel(self, request):
        serializer = ArchivoExcelSerializer(data=request.data)
        if serializer.is_valid():
            archivo = serializer.validated_data['archivo']
            try:
                # 1. Extraer imágenes incrustadas del Excel
                imagenes_excel = extraer_imagenes_de_excel(archivo)
                
                # 2. Leemos el Excel
                df = pd.read_excel(archivo)
                total_registros = len(df)
                
                # 2. Limpieza y validación de duplicados
                df['curp_limpia'] = df['curp'].apply(lambda x: str(x).strip().upper() if pd.notnull(x) else '')
                # Filtrar CURPs vacías para validación
                curps_excel = [c for c in df['curp_limpia'].tolist() if c]

                curps_duplicadas_en_bd = list(Enrolamiento.objects.filter(curp__in=curps_excel).values_list('curp', flat=True))
                cantidad_duplicados = len(curps_duplicadas_en_bd)

                # --- ESCENARIO A: HAY DUPLICADOS (BLOQUEO) ---
                if cantidad_duplicados > 0:
                    detalles_duplicados = df[df['curp_limpia'].isin(curps_duplicadas_en_bd)][['curp', 'nombre', 'paterno']].to_dict('records')
                    return Response({
                        "status": "error",
                        "mensaje": "Se encontraron registros duplicados. Por favor, quítalos del Excel para continuar.",
                        "resumen": {
                            "total_registros": total_registros,
                            "registros_validos": total_registros - cantidad_duplicados,
                            "registros_duplicados": cantidad_duplicados

                        },
                        "lista_duplicados": detalles_duplicados
                    }, status=status.HTTP_400_BAD_REQUEST)

                # --- ESCENARIO B: TODO LIMPIO (INSERTAR EN AMBAS TABLAS) ---
                sig_objs = []
                enrolamiento_objs = []

                for index, row in df.iterrows():
                    # Lógica de fechas
                    raw_inicio = row.get('Inicio Vigencia') or row.get('inicio_vig') or row.get('inicio vigencia')
                    raw_fin = row.get('Fin Vigencia') or row.get('fin_vig') or row.get('fin vigencia')
                    
                    inicio_vig = str(raw_inicio).split(' ')[0] if pd.notnull(raw_inicio) else None
                    fin_vig = str(raw_fin).split(' ')[0] if pd.notnull(raw_fin) else None

                    rfc = row.get('rfc') or row.get('RFC')
                    apellidos = f"{row.get('paterno') or ''} {row.get('materno') or ''}".strip()
                    
                    # Procesar foto (prioriza imagen incrustada en Excel)
                    # +2 porque: pandas usa índice 0, Excel empieza en 1, y tenemos header
                    fila_excel = index + 2
                    foto_bytes = procesar_foto_desde_excel(row.get('foto'), fila_excel, imagenes_excel)

                    # Objeto para SicreTblSig
                    sig = SicreTblSig(
                        num_empleado=row.get('num_empleado'),
                        rfc=rfc,
                        curp=row.get('curp_limpia'),
                        nombre=row.get('nombre'),
                        paterno=row.get('paterno'),
                        materno=row.get('materno'),
                        apellidos=apellidos,
                        puesto=row.get('puesto'),
                        adscripcion=row.get('adscripcion'),
                        inicio_vig=inicio_vig,
                        fin_vig=fin_vig,
                        eladia=row.get('eladia'),
                        foto=foto_bytes
                    )
                    sig_objs.append(sig)

                    # Objeto para Enrolamiento
                    enrolamiento = Enrolamiento(
                        rfc=rfc,
                        curp=row.get('curp_limpia'),
                        num_empleado=row.get('num_empleado'),
                        nombre=row.get('nombre'),
                        paterno=row.get('paterno'),
                        materno=row.get('materno'),
                        apellidos=apellidos,
                        puesto=row.get('puesto'),
                        adscripcion=row.get('adscripcion'),
                        folio=row.get('folio'),
                        inicio_vig=inicio_vig,
                        fin_vig=fin_vig,
                        foto=foto_bytes,
                        eladia=row.get('eladia'),
                        activo=1
                    )
                    enrolamiento_objs.append(enrolamiento)
                
                # Guardado masivo (Muy rápido)
                SicreTblSig.objects.bulk_create(sig_objs)
                Enrolamiento.objects.bulk_create(enrolamiento_objs)
                
                return Response({
                    "status": "success",
                    "mensaje": "Carga de registros exitosa",
                    "resumen": {
                        "total_procesados": len(enrolamiento_objs)
                    }
                }, status=status.HTTP_201_CREATED)

            except Exception as e:
                return Response({"error": "Error interno: " + str(e)}, status=status.HTTP_400_BAD_REQUEST)
        
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


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