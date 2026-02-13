# enrolamiento/views.py
from rest_framework import viewsets, filters, status
from rest_framework.decorators import action
from rest_framework.response import Response
from .models import Enrolamiento, SicreTblSig
from .serializers import EnrolamientoSerializer, SigSerializer, EnrolamientoDataTableSerializer, ArchivoExcelSerializer, LoginSerializer
from django.db.models import Q, Count, Min
from django.db import models
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


def extraer_imagenes_de_excel(archivo):
    """
    Versión 5.0 (Final): Soporte DUAL para FOTOS (Col L) y FIRMAS (Col M).
    Mantiene el nombre original para compatibilidad.
    """
    from openpyxl import load_workbook
    from io import BytesIO
    import zipfile
    import posixpath
    import xml.etree.ElementTree as ET
    
    print(f"--- INICIANDO EXTRACCIÓN DUAL: {archivo} ---")
    fotos_por_fila = {}
    firmas_por_fila = {}
    
    # --- INTENTO 1: MÉTODO ESTÁNDAR (OpenPyXL) ---
    try:
        if hasattr(archivo, 'seek'): archivo.seek(0)
        # data_only=False es vital para ver las imágenes
        wb = load_workbook(archivo, data_only=False)
        ws = wb.active
        
        imagenes = getattr(ws, '_images', []) or getattr(ws, 'images', [])
        print(f"   -> Método Estándar detectó: {len(imagenes)} imágenes")

        for image in imagenes:
            try:
                # Coordenadas (Excel base 1)
                # .anchor._from.row es índice 0, por eso sumamos 1
                row = image.anchor._from.row + 1
                col = image.anchor._from.col + 1 
                
                # Extraer bytes
                img_bytes = None
                if hasattr(image, '_data'):
                    img_bytes = image._data()
                elif hasattr(image, 'ref'):
                    buf = BytesIO()
                    image.ref.save(buf)
                    img_bytes = buf.getvalue()
                
                if img_bytes:
                    # Lógica de Columnas: 
                    # Columna 12 = L (Foto)
                    # Columna 13 = M (Firma)
                    if col == 12: 
                        fotos_por_fila[row] = img_bytes
                    elif col == 13:
                        firmas_por_fila[row] = img_bytes
                        
            except Exception as e:
                print(f"   -> Error leyendo imagen estándar: {e}")
        
        wb.close()
    except Exception as e:
        print(f"   -> Falló método estándar: {e}")

    # --- INTENTO 2: PLAN B ROBUSTO (ZIP + DRAWING RELS) ---
    # Si el método estándar falló y no detectó NADA, resolvemos anclas reales
    # desde drawing.xml y sus relaciones (NO por orden alfabético de archivos).
    if not fotos_por_fila and not firmas_por_fila:
        print("⚠️ ACTIVANDO EXTRACCIÓN ROBUSTA DUAL (PLAN B ZIP + RELS)")
        try:
            if hasattr(archivo, 'seek'): archivo.seek(0)
            if zipfile.is_zipfile(archivo):
                with zipfile.ZipFile(archivo, 'r') as z:
                    ns_main = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    ns_rel = {'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
                    ns_pkg_rel = {'pr': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                    ns_draw = {
                        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                    }

                    def _norm_zip_path(base_dir, target):
                        # target puede venir absoluto (/xl/...) o relativo (../drawings/...)
                        if target.startswith('/'):
                            target = target[1:]
                        else:
                            target = posixpath.normpath(posixpath.join(base_dir, target))
                        return target

                    # 1) Localizar worksheet activo real a partir de workbook.xml
                    wb_root = ET.fromstring(z.read('xl/workbook.xml'))
                    sheets = wb_root.find('m:sheets', ns_main)
                    first_sheet = sheets.find('m:sheet', ns_main) if sheets is not None else None
                    if first_sheet is None:
                        raise ValueError('No se encontró ninguna hoja en workbook.xml')
                    sheet_rid = first_sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

                    wb_rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                    rel_node = None
                    for rel in wb_rels_root.findall('pr:Relationship', ns_pkg_rel):
                        if rel.attrib.get('Id') == sheet_rid:
                            rel_node = rel
                            break
                    if rel_node is None:
                        raise ValueError('No se pudo resolver la relación de la hoja activa')

                    sheet_path = _norm_zip_path('xl', rel_node.attrib.get('Target', ''))
                    sheet_dir = posixpath.dirname(sheet_path)

                    # 2) Leer drawing r:id desde la hoja activa
                    sheet_root = ET.fromstring(z.read(sheet_path))
                    drawing_tag = sheet_root.find('m:drawing', {**ns_main, **ns_rel})
                    if drawing_tag is None:
                        print('   -> La hoja activa no tiene drawing asociado')
                    else:
                        drawing_rid = drawing_tag.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

                        # 3) Resolver drawing.xml desde sheet rels
                        sheet_rels_path = posixpath.join(sheet_dir, '_rels', posixpath.basename(sheet_path) + '.rels')
                        sheet_rels_root = ET.fromstring(z.read(sheet_rels_path))

                        drawing_rel = None
                        for rel in sheet_rels_root.findall('pr:Relationship', ns_pkg_rel):
                            if rel.attrib.get('Id') == drawing_rid:
                                drawing_rel = rel
                                break

                        if drawing_rel is None:
                            print('   -> No se encontró relación al drawing en la hoja activa')
                        else:
                            drawing_path = _norm_zip_path(sheet_dir, drawing_rel.attrib.get('Target', ''))
                            drawing_dir = posixpath.dirname(drawing_path)
                            drawing_rels_path = posixpath.join(drawing_dir, '_rels', posixpath.basename(drawing_path) + '.rels')

                            # 4) Mapa rId -> target_path (más tolerante)
                            drawing_rels_root = ET.fromstring(z.read(drawing_rels_path))
                            media_by_rid = {}
                            for rel in drawing_rels_root.findall('pr:Relationship', ns_pkg_rel):
                                rel_id = rel.attrib.get('Id')
                                target = rel.attrib.get('Target', '')
                                if rel_id and target:
                                    media_by_rid[rel_id] = _norm_zip_path(drawing_dir, target)

                            print(f"   -> Relaciones en drawing: {len(media_by_rid)}")

                            # 5) Recorrer anchors con posición real y resolver bytes
                            drawing_root = ET.fromstring(z.read(drawing_path))
                            anchors = (
                                drawing_root.findall('.//xdr:twoCellAnchor', ns_draw)
                                + drawing_root.findall('.//xdr:oneCellAnchor', ns_draw)
                                + drawing_root.findall('.//xdr:absoluteAnchor', ns_draw)
                            )
                            print(f"   -> Anchors detectados en drawing: {len(anchors)}")

                            # Guardamos candidatos para asignar por fila/posición horizontal
                            candidates = []

                            for anchor in anchors:
                                from_node = anchor.find('xdr:from', ns_draw)
                                pic_node = anchor.find('.//xdr:pic', ns_draw)
                                if from_node is None or pic_node is None:
                                    continue

                                col_node = from_node.find('xdr:col', ns_draw)
                                row_node = from_node.find('xdr:row', ns_draw)
                                if col_node is None or row_node is None:
                                    continue

                                col = int(col_node.text) + 1
                                row = int(row_node.text) + 1

                                blip = pic_node.find('.//a:blip', ns_draw)
                                if blip is None:
                                    continue

                                rid_embed = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                rid_link = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link')
                                rid = rid_embed or rid_link
                                if not rid:
                                    continue

                                media_path = media_by_rid.get(rid)
                                if not media_path:
                                    continue

                                # Solo nos interesan recursos internos de media
                                if 'xl/media/' not in media_path.replace('\\', '/'):
                                    continue

                                if media_path not in z.namelist():
                                    continue

                                img_bytes = z.read(media_path)
                                if not img_bytes:
                                    continue

                                candidates.append((row, col, media_path, img_bytes))

                            print(f"   -> Candidatos de imagen resueltos: {len(candidates)}")

                            # 6) Asignar foto/firma por fila usando orden horizontal (izq->der)
                            # Esto evita depender de si Excel ancla en K/L o L/M.
                            by_row = {}
                            for row, col, media_path, img_bytes in candidates:
                                by_row.setdefault(row, []).append((col, media_path, img_bytes))

                            for row, items in by_row.items():
                                items.sort(key=lambda x: x[0])

                                if len(items) == 1:
                                    col, media_path, img_bytes = items[0]
                                    # regla simple por columna aproximada
                                    if col <= 12:
                                        fotos_por_fila[row] = img_bytes
                                        print(f"   -> [ZIP-RELS] {media_path} es FOTO de Fila {row} (col={col})")
                                    else:
                                        firmas_por_fila[row] = img_bytes
                                        print(f"   -> [ZIP-RELS] {media_path} es FIRMA de Fila {row} (col={col})")
                                else:
                                    # Izquierda = Foto, Derecha = Firma
                                    col_f, media_f, bytes_f = items[0]
                                    col_s, media_s, bytes_s = items[1]
                                    fotos_por_fila[row] = bytes_f
                                    firmas_por_fila[row] = bytes_s
                                    print(f"   -> [ZIP-RELS] {media_f} es FOTO de Fila {row} (col={col_f})")
                                    print(f"   -> [ZIP-RELS] {media_s} es FIRMA de Fila {row} (col={col_s})")

        except Exception as e:
            print(f"❌ Falló Plan B Dual (RELS): {e}")

    # Rebobinar siempre al final para que Pandas no falle después
    if hasattr(archivo, 'seek'): archivo.seek(0)
    
    return fotos_por_fila, firmas_por_fila


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

                # Normalizar cabeceras para evitar problemas de mayúsculas/espacios
                df.columns = df.columns.str.strip().str.lower()

                total_registros = len(df)

                # Extraer imágenes incrustadas separadas por columna (foto/firma)
                fotos_excel, firmas_excel = extraer_imagenes_de_excel(archivo)

                # Convertir a lista de diccionarios para el front
                registros_preview = []
                for index, row in df.iterrows():
                    fila_excel = index + 2  # header + índice base 0

                    # Detectar foto/firma con la misma lógica robusta de carga
                    val_foto = row.get('foto')
                    val_firma = row.get('firma')

                    foto_bytes = procesar_foto_desde_excel(val_foto, fila_excel, fotos_excel)
                    firma_bytes = procesar_foto_desde_excel(val_firma, fila_excel, firmas_excel)

                    tiene_foto = foto_bytes is not None and len(foto_bytes) > 0
                    tiene_firma = firma_bytes is not None and len(firma_bytes) > 0

                    registro = {
                        'num_empleado': row.get('num_empleado'),
                        'rfc': row.get('rfc'),
                        'curp': row.get('curp'),
                        'nombre': row.get('nombre'),
                        'paterno': row.get('paterno'),
                        'materno': row.get('materno'),
                        'puesto': row.get('puesto'),
                        'adscripcion': row.get('adscripcion'),
                        'inicio_vig': row.get('inicio_vig') or row.get('inicio vigencia'),
                        'fin_vig': row.get('fin_vig') or row.get('fin vigencia'),
                        'eladia': row.get('eladia'),

                        # Compatibilidad con front actual + explícitos
                        'foto': tiene_foto,
                        'firma': tiene_firma,
                        'tiene_foto': tiene_foto,
                        'tiene_firma': tiene_firma,
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
                # ---------------------------------------------------------
                # PASO 1: EXTRAER IMÁGENES Y FIRMAS (Plan B incluido)
                # ---------------------------------------------------------
                fotos_excel, firmas_excel = extraer_imagenes_de_excel(archivo)
                
                # ---------------------------------------------------------
                # PASO 2: LEER DATOS Y LIMPIEZA
                # ---------------------------------------------------------
                df = pd.read_excel(archivo)
                
                # Normalizar cabeceras (quita espacios y convierte a minúsculas)
                df.columns = df.columns.str.strip().str.lower()

                # --- 🛡️ FILTRO ANTI-FANTASMAS (LA SOLUCIÓN A TU ERROR) ---
                # 1. Eliminamos filas donde la CURP sea completamente nula (NaN)
                df = df.dropna(subset=['curp'])
                # 2. Eliminamos filas donde la CURP sea texto vacío ('') o espacios ('   ')
                df = df[df['curp'].astype(str).str.strip() != '']
                # ---------------------------------------------------------

                total_registros = len(df)

                # ---------------------------------------------------------
                # PASO 3: VALIDACIÓN DE DUPLICADOS EN BD
                # ---------------------------------------------------------
                # Preparamos CURPs para buscar (Mayúsculas y sin espacios)
                df['curp_limpia'] = df['curp'].apply(lambda x: str(x).strip().upper())
                curps_excel = df['curp_limpia'].tolist()

                # Buscamos si alguna de estas CURPs ya existe en Enrolamiento
                curps_duplicadas_en_bd = list(Enrolamiento.objects.filter(curp__in=curps_excel).values_list('curp', flat=True))
                
                if len(curps_duplicadas_en_bd) > 0:
                    return Response({
                        "status": "error",
                        "mensaje": "Se encontraron registros duplicados (CURP ya existe).",
                        "lista_duplicados": curps_duplicadas_en_bd
                    }, status=status.HTTP_400_BAD_REQUEST)

                # ---------------------------------------------------------
                # PASO 4: CALCULAR LOTE Y FECHAS
                # ---------------------------------------------------------
                from datetime import datetime
                
                # Calcular siguiente Lote (001, 002...)
                lote_maximo = SicreTblSig.objects.filter(lote__isnull=False).order_by('-lote').first()
                if lote_maximo and lote_maximo.lote:
                    try:
                        siguiente_numero = int(lote_maximo.lote) + 1
                    except:
                        siguiente_numero = 1
                else:
                    siguiente_numero = 1
                siguiente_lote = str(siguiente_numero).zfill(3)
                
                fecha_carga_actual = datetime.now()
                
                sig_objs = []
                enrolamiento_objs = []

                # ---------------------------------------------------------
                # PASO 5: ITERAR Y CREAR OBJETOS
                # ---------------------------------------------------------
                for index, row in df.iterrows():
                    # Calculamos la fila real en Excel (Header es 1 + Index 0 = Fila 2)
                    fila_excel = index + 2
                    
                    # Limpieza de datos básicos
                    rfc = str(row.get('rfc') or row.get('RFC') or '').strip().upper()
                    curp = str(row.get('curp_limpia')).strip().upper()
                    num_empleado = str(row.get('num_empleado') or '').strip()
                    
                    # Nombres
                    nombre = str(row.get('nombre') or '').strip()
                    paterno = str(row.get('paterno') or '').strip()
                    materno = str(row.get('materno') or '').strip()
                    apellidos = f"{paterno} {materno}".strip()
                    
                    # Otros datos
                    puesto = str(row.get('puesto') or '').strip()
                    adscripcion = str(row.get('adscripcion') or '').strip()
                    eladia = str(row.get('eladia') or '').strip()
                    
                    # Fechas (Manejo de formatos variados)
                    raw_inicio = row.get('inicio_vig') or row.get('inicio vigencia')
                    raw_fin = row.get('fin_vig') or row.get('fin vigencia')
                    inicio_vig = str(raw_inicio).split(' ')[0] if pd.notnull(raw_inicio) else None
                    fin_vig = str(raw_fin).split(' ')[0] if pd.notnull(raw_fin) else None

                    # --- PROCESAR FOTO (Columna L) ---
                    # Busca en columna 'foto' o usa la fila para buscar en el diccionario
                    val_foto = row.get('foto')
                    foto_bytes = procesar_foto_desde_excel(val_foto, fila_excel, fotos_excel)
                    
                    # --- PROCESAR FIRMA (Columna M) ---
                    val_firma = row.get('firma') or row.get('FIRMA')
                    firma_bytes = procesar_foto_desde_excel(val_firma, fila_excel, firmas_excel)

                    # 1. Objeto para Histórico (SicreTblSig)
                    sig = SicreTblSig(
                        num_empleado=num_empleado,
                        rfc=rfc,
                        curp=curp,
                        nombre=nombre,
                        paterno=paterno,
                        materno=materno,
                        apellidos=apellidos,
                        puesto=puesto,
                        adscripcion=adscripcion,
                        inicio_vig=inicio_vig,
                        fin_vig=fin_vig,
                        eladia=eladia,
                        foto=foto_bytes,
                        firma=firma_bytes,  # Campo Firma
                        lote=siguiente_lote,
                        fecha_carga=fecha_carga_actual
                    )
                    sig_objs.append(sig)

                    # 2. Objeto para Sistema Actual (Enrolamiento)
                    enrolamiento = Enrolamiento(
                        num_empleado=num_empleado,
                        rfc=rfc,
                        curp=curp,
                        nombre=nombre,
                        paterno=paterno,
                        materno=materno,
                        apellidos=apellidos,
                        puesto=puesto,
                        adscripcion=adscripcion,
                        inicio_vig=inicio_vig,
                        fin_vig=fin_vig,
                        eladia=eladia,
                        foto=foto_bytes,
                        firma=firma_bytes, # Campo Firma
                        activo=1
                    )
                    enrolamiento_objs.append(enrolamiento)
                
                # ---------------------------------------------------------
                # PASO 6: GUARDADO MASIVO (BULK INSERT)
                # ---------------------------------------------------------
                SicreTblSig.objects.bulk_create(sig_objs)
                # Usamos ignore_conflicts=True por seguridad en Enrolamiento si usas Postgres, 
                # en MySQL bulk_create es directo. Si quieres actualizar existentes, 
                # bulk_create no sirve, tendrías que usar update_or_create, 
                # pero para carga inicial masiva esto es lo mejor.
                Enrolamiento.objects.bulk_create(enrolamiento_objs)
                
                return Response({
                    "status": "success",
                    "mensaje": "Carga completada exitosamente.",
                    "resumen": {
                        "registros_procesados": len(enrolamiento_objs),
                        "lote_generado": siguiente_lote,
                        "fecha": fecha_carga_actual
                    }
                }, status=status.HTTP_201_CREATED)

            except Exception as e:
                # Log del error en consola para que veas qué pasó
                print(f"❌ ERROR EN SUBIR_EXCEL: {str(e)}")
                return Response({"error": f"Error procesando el archivo: {str(e)}"}, status=status.HTTP_400_BAD_REQUEST)
        
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    @action(detail=False, methods=['get'], url_path='historial-cargas')
    def historial_cargas(self, request):
        """
        Retorna un resumen de todas las cargas masivas realizadas.
        Incluye: lote, fecha_carga, total_registros, y nombres de los primeros 10 registros.
        """
        from django.db.models import Count, Min
        
        try:
            # Obtener todos los lotes distintos con información agregada
            lotes = SicreTblSig.objects.values('lote', 'fecha_carga').annotate(
                total_registros=Count('rfc'),
                primera_carga=Min('fecha_carga')
            ).filter(lote__isnull=False).order_by('-lote')
            
            resultado = []
            for lote_info in lotes:
                lote_num = lote_info['lote']
                
                # Obtener los primeros 10 registros de este lote
                primeros_10 = SicreTblSig.objects.filter(lote=lote_num).values(
                    'nombre', 'paterno', 'materno'
                )[:10]
                
                # Construir nombres completos
                nombres_preview = [
                    f"{r['nombre']} {r['paterno']} {r['materno'] or ''}".strip()
                    for r in primeros_10
                ]
                
                resultado.append({
                    'lote': lote_num,
                    'fecha_carga': lote_info['primera_carga'],
                    'total_registros': lote_info['total_registros'],
                    'preview_nombres': nombres_preview
                })
            
            return Response({
                'status': 'success',
                'total_cargas': len(resultado),
                'cargas': resultado
            }, status=status.HTTP_200_OK)
            
        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al obtener historial: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
    
    @action(detail=False, methods=['get'], url_path='detalle-lote/(?P<lote_id>[0-9]+)')
    def detalle_lote(self, request, lote_id=None):
        """
        Retorna todos los registros de un lote específico.
        Parámetro: lote_id (entero) - Número de lote a consultar
        """
        try:
            # Validar que el lote existe
            if not SicreTblSig.objects.filter(lote=lote_id).exists():
                return Response({
                    'status': 'error',
                    'mensaje': f'El lote {lote_id} no existe'
                }, status=status.HTTP_404_NOT_FOUND)
            
            # Obtener todos los registros del lote
            registros = SicreTblSig.objects.filter(lote=lote_id).order_by('rfc')
            serializer = SigSerializer(registros, many=True)
            
            # Información del lote
            info_lote = SicreTblSig.objects.filter(lote=lote_id).aggregate(
                total_registros=Count('rfc'),
                fecha_carga=Min('fecha_carga')
            )
            
            return Response({
                'status': 'success',
                'lote': int(lote_id),
                'fecha_carga': info_lote['fecha_carga'],
                'total_registros': info_lote['total_registros'],
                'registros': serializer.data
            }, status=status.HTTP_200_OK)
            
        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al obtener detalle del lote: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


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