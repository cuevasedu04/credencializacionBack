# enrolamiento/views.py
from rest_framework import viewsets, filters, status
from rest_framework.decorators import action, api_view
from rest_framework.response import Response
from rest_framework.pagination import PageNumberPagination
from .models import Enrolamiento, SicreTblSig, EnrolamientoFamiliar
from .serializers import EnrolamientoSerializer, SigSerializer, EnrolamientoDataTableSerializer, ArchivoExcelSerializer, LoginSerializer, EnrolamientoFamiliarSerializer
from django.db.models import Q, Count, Min
from django.db import models, transaction, connection
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
from django.utils import timezone


def _cargo_a_nivel(cargo: str) -> str:
    """
    Determina el nivel de credencial (nivel_credencial) a partir del campo CARGO/PUESTO.
    Niveles válidos:
        TITULAR, DIRECTOR_GENERAL, DIRECTOR_CENTRAL, DIRECTOR_DE_AREA,
        SUBDIRECTOR, JEFE_DE_DEPARTAMENTO, ENLACE, SEGURIDAD_INSTITUCIONAL
    """
    if not cargo:
        return 'ENLACE'
    c = str(cargo).strip().upper()

    if 'DIRECTOR GENERAL' in c or 'ADMINISTRADOR GENERAL' in c or 'TITULAR' in c:
        return 'TITULAR'
    if 'DIRECTOR CENTRAL' in c:
        return 'DIRECTOR_CENTRAL'
    if c.startswith('DIRECTOR DE AREA') or c.startswith('DIRECTOR DE ÁREA'):
        return 'DIRECTOR_DE_AREA'
    if c.startswith('DIRECTOR'):
        return 'DIRECTOR_DE_AREA'
    if 'SUBDIRECTOR' in c:
        return 'SUBDIRECTOR'
    if 'JEFE DE DEPARTAMENTO' in c or 'JEFE DEL DEPARTAMENTO' in c:
        return 'JEFE_DE_DEPARTAMENTO'
    if 'ENLACE' in c:
        return 'ENLACE'
    if 'OPERATIVO' in c or 'SEGURIDAD INSTITUCIONAL' in c or 'SEGURIDAD_INSTITUCIONAL' in c:
        return 'SEGURIDAD_INSTITUCIONAL'
    return 'ENLACE'


def extraer_imagenes_de_excel(archivo):
    """
    Versión 5.0 (Final): Soporte DUAL para FOTOS (Col L) y FIRMAS (Col M).
    Wrapper que delega en extraer_imagenes_de_excel_real.
    """
    return extraer_imagenes_de_excel_real(archivo)


def _get_num_empleados_con_foto_firma():
    """
    Devuelve un set de valores de num_empleado que ya tienen FOTO y FIRMA
    en safirho_db.NW_EMPL_FOTO_ANAM (en todas sus variantes de formato).
    """
    try:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT EMPLID FROM safirho_db.NW_EMPL_FOTO_ANAM
                WHERE FOTO IS NOT NULL AND LENGTH(FOTO) > 1
                  AND FIRMA IS NOT NULL AND LENGTH(FIRMA) > 1
                """
            )
            rows = cursor.fetchall()
        result = set()
        for (emplid,) in rows:
            emplid_str = str(emplid).strip()
            result.add(emplid_str)                          # original  (00020232019)
            result.add(emplid_str.lstrip('0') or '0')      # sin ceros  (20232019)
        return result
    except Exception:
        return set()


def extraer_imagenes_de_excel_real(archivo):
    """\n    Versión 5.0 (Final): Soporte DUAL para FOTOS (Col L) y FIRMAS (Col M).
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
        Filtros: foto y firma no nulos (incluyendo flag 1 o registro en NW_EMPL_FOTO_ANAM),
        y que no estén marcados como impresos (impreso != 1).
        """
        completos = _get_num_empleados_con_foto_firma()

        queryset = self.get_queryset().exclude(
            Q(foto__isnull=True) | Q(foto=b'') |
            Q(firma__isnull=True) | Q(firma=b'')
        ).filter(
            Q(impreso__isnull=True) | Q(impreso=0) | ~Q(impreso=1)
        )

        # También incluir empleados con foto+firma en NW_EMPL_FOTO_ANAM
        if completos:
            qs_externos = self.get_queryset().filter(
                num_empleado__in=completos
            ).filter(
                Q(impreso__isnull=True) | Q(impreso=0) | ~Q(impreso=1)
            )
            from itertools import chain
            queryset = list(chain(queryset, qs_externos.exclude(pk__in=queryset.values_list('pk', flat=True))))
        else:
            queryset = list(queryset)

        queryset_familiares = EnrolamientoFamiliar.objects.exclude(
            Q(foto__isnull=True) | Q(foto=b'') |
            Q(firma__isnull=True) | Q(firma=b'')
        ).filter(
            Q(impreso__isnull=True) | Q(impreso=0) | ~Q(impreso=1)
        )

        serializer_enrolamiento = EnrolamientoDataTableSerializer(queryset, many=True)
        serializer_familiar = EnrolamientoFamiliarSerializer(queryset_familiares, many=True)

        data_enrolamiento = []
        for item in serializer_enrolamiento.data:
            registro = dict(item)
            registro['source_table'] = 'enrolamiento'
            if int(registro.get('nuevo_laredo') or 0) == 1:
                registro['tipo_credencial'] = 'provisional'
            else:
                registro['tipo_credencial'] = 'anam'
            data_enrolamiento.append(registro)

        data_familiares = []
        for item in serializer_familiar.data:
            registro = dict(item)
            registro['source_table'] = 'familiar'
            registro['tipo_credencial'] = 'familiar'
            registro['folio'] = registro.get('folio_familiares')
            data_familiares.append(registro)

        data = sorted(
            data_enrolamiento + data_familiares,
            key=lambda x: x.get('fecha_enrolamiento') or x.get('fecha_registro') or '',
            reverse=True
        )
        
        page = self.paginate_queryset(data)
        if page is not None:
            return self.get_paginated_response(page)
        return Response(data)
    
    @action(detail=False, methods=['get'])
    def pendientes(self, request):
        """
        Muestra solo los registros que faltan de foto O firma,
        y que adicionalmente NO tienen ambas en safirho_db.NW_EMPL_FOTO_ANAM.
        """
        completos = _get_num_empleados_con_foto_firma()

        queryset = self.get_queryset().filter(
            Q(foto__isnull=True) | Q(foto=b'') |
            Q(firma__isnull=True) | Q(firma=b'')
        )

        if completos:
            queryset = queryset.exclude(num_empleado__in=completos)

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
            nombre = (sig.nombres or '').strip()
            paterno = (sig.primer_apellido or '').strip()
            materno = (sig.segundo_apellido or '').strip()
            apellidos = f"{paterno} {materno}".strip()
            rfc_ref = (sig.empleado_anam or sig.no_empleado or sig.curp or '').strip()

            obj, created = Enrolamiento.objects.update_or_create(
                curp=sig.curp,
                defaults={
                    'rfc': rfc_ref,
                    'num_empleado': sig.no_empleado,
                    'curp': sig.curp,
                    'nombre': nombre,
                    'paterno': paterno,
                    'materno': materno,
                    'apellidos': apellidos,
                    'puesto': sig.cargo,
                    'adscripcion': sig.area,
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
        Permite separar series por nuevo_laredo con query param ?nuevo_laredo=0|1
        """
        try:
            nuevo_laredo_param = request.query_params.get('nuevo_laredo', None)
            folios_queryset = Enrolamiento.objects.exclude(
                Q(folio__isnull=True) | Q(folio='')
            )

            if nuevo_laredo_param in ['0', '1']:
                nuevo_laredo_flag = int(nuevo_laredo_param)
                if nuevo_laredo_flag == 1:
                    folios_queryset = folios_queryset.filter(nuevo_laredo=1)
                else:
                    folios_queryset = folios_queryset.filter(
                        Q(nuevo_laredo=0) | Q(nuevo_laredo__isnull=True)
                    )

            # Filtrar folios no nulos y obtener todos
            folios = folios_queryset.values_list('folio', flat=True)
            
            if not folios:
                # Si no hay folios, empezar desde 000001 (6 dígitos por defecto)
                return Response({
                    'status': 'success',
                    'folio_maximo': '000000',
                    'siguiente_folio': '000001',
                    'serie_nuevo_laredo': nuevo_laredo_param
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
                    'siguiente_folio': '000001',
                    'serie_nuevo_laredo': nuevo_laredo_param
                }, status=status.HTTP_200_OK)
            
            # Calcular siguiente folio manteniendo el formato
            siguiente_valor = folio_max_valor + 1
            siguiente_folio = str(siguiente_valor).zfill(longitud_formato)
            
            return Response({
                'status': 'success',
                'folio_maximo': folio_max_str,
                'siguiente_folio': siguiente_folio,
                'serie_nuevo_laredo': nuevo_laredo_param
            }, status=status.HTTP_200_OK)
            
        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al obtener folio máximo: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    @action(detail=True, methods=['post'], url_path='guardar-foto-firma')
    def guardar_foto_firma(self, request, pk=None):
        """
        Guarda foto y firma en safirho_db.NW_EMPL_FOTO_ANAM (INSERT o UPDATE).
        Marca enrolamiento.foto y enrolamiento.firma con flag b'\\x01' para excluirlo
        del endpoint pendientes sin guardar el blob en sicre_db.
        """
        enrolamiento = self.get_object()
        foto_b64 = request.data.get('foto')
        firma_b64 = request.data.get('firma')

        if not foto_b64 or not firma_b64:
            return Response({'status': 'error', 'mensaje': 'Faltan foto o firma'}, status=400)

        def clean_b64(data):
            if isinstance(data, str) and 'base64,' in data:
                data = data.split('base64,')[1]
            return base64.b64decode(data)

        try:
            foto_bytes = clean_b64(foto_b64)
            firma_bytes = clean_b64(firma_b64)
        except Exception as e:
            return Response({'status': 'error', 'mensaje': f'Error al decodificar imagen: {e}'}, status=400)

        emplid = str(enrolamiento.num_empleado or '').strip()
        if not emplid:
            return Response({'status': 'error', 'mensaje': 'El empleado no tiene número de empleado'}, status=400)

        emplid_padded = emplid.zfill(11)
        emplid_stripped = emplid.lstrip('0') or '0'
        variantes = list({emplid, emplid_padded, emplid_stripped})
        placeholders = ','.join(['%s'] * len(variantes))

        try:
            with connection.cursor() as cursor:
                cursor.execute(
                    f"SELECT EMPLID FROM safirho_db.NW_EMPL_FOTO_ANAM WHERE EMPLID IN ({placeholders}) LIMIT 1",
                    variantes
                )
                existing = cursor.fetchone()
                if existing:
                    cursor.execute(
                        "UPDATE safirho_db.NW_EMPL_FOTO_ANAM "
                        "SET FOTO=%s, FIRMA=%s, TipoMime='image/jpeg', updated_at=NOW() "
                        "WHERE EMPLID=%s",
                        [foto_bytes, firma_bytes, existing[0]]
                    )
                else:
                    cursor.execute(
                        "INSERT INTO safirho_db.NW_EMPL_FOTO_ANAM "
                        "(EMPLID, FOTO, FIRMA, TipoMime, updated_at) VALUES (%s, %s, %s, 'image/jpeg', NOW())",
                        [emplid_padded, foto_bytes, firma_bytes]
                    )
        except Exception as e:
            return Response({'status': 'error', 'mensaje': f'Error al guardar en BD externa: {e}'}, status=500)

        enrolamiento.foto = b'\x01'
        enrolamiento.firma = b'\x01'
        enrolamiento.save(update_fields=['foto', 'firma'])

        return Response({'status': 'success', 'mensaje': 'Foto y firma guardadas correctamente'})

    @action(detail=False, methods=['get', 'post'], url_path='busqueda-avanzada')
    def busqueda_avanzada(self, request):
        """
        Endpoint de búsqueda avanzada para ag-grid.
        Permite filtrar por múltiples campos y retorna todos los detalles.
        Soporta filtros por: num_empleado, rfc, curp, nombre, paterno, materno,
        puesto, adscripcion, folio, impreso, fecha_expedicion_desde, fecha_expedicion_hasta.
        """
        queryset = self.get_queryset()
        queryset_familiares = EnrolamientoFamiliar.objects.all().order_by('-id_enrolamiento')
        
        # Obtener parámetros de filtro (soporta GET y POST)
        params = request.query_params if request.method == 'GET' else request.data
        
        # Filtros de texto
        if params.get('num_empleado'):
            queryset = queryset.filter(num_empleado__icontains=params['num_empleado'])
            queryset_familiares = queryset_familiares.filter(num_empleado__icontains=params['num_empleado'])
        
        if params.get('rfc'):
            queryset = queryset.filter(rfc__icontains=params['rfc'])
            queryset_familiares = queryset_familiares.filter(rfc__icontains=params['rfc'])
        
        if params.get('curp'):
            queryset = queryset.filter(curp__icontains=params['curp'])
            queryset_familiares = queryset_familiares.filter(curp__icontains=params['curp'])
        
        if params.get('nombre'):
            queryset = queryset.filter(nombre__icontains=params['nombre'])
            queryset_familiares = queryset_familiares.filter(nombre__icontains=params['nombre'])
        
        if params.get('paterno'):
            queryset = queryset.filter(paterno__icontains=params['paterno'])
            queryset_familiares = queryset_familiares.filter(paterno__icontains=params['paterno'])
        
        if params.get('materno'):
            queryset = queryset.filter(materno__icontains=params['materno'])
            queryset_familiares = queryset_familiares.filter(materno__icontains=params['materno'])
        
        if params.get('puesto'):
            queryset = queryset.filter(puesto__icontains=params['puesto'])
            queryset_familiares = queryset_familiares.filter(puesto__icontains=params['puesto'])
        
        if params.get('adscripcion'):
            queryset = queryset.filter(adscripcion__icontains=params['adscripcion'])
            queryset_familiares = queryset_familiares.filter(adscripcion__icontains=params['adscripcion'])
        
        if params.get('folio'):
            queryset = queryset.filter(folio__icontains=params['folio'])
            queryset_familiares = queryset_familiares.filter(folio_familiares__icontains=params['folio'])
        
        # Filtro por estado de impresión
        if params.get('impreso') is not None:
            impreso_val = params['impreso']
            if impreso_val == '1' or impreso_val == 1:
                queryset = queryset.filter(impreso=1)
                queryset_familiares = queryset_familiares.filter(impreso=1)
            elif impreso_val == '0' or impreso_val == 0:
                queryset = queryset.filter(Q(impreso=0) | Q(impreso__isnull=True))
                queryset_familiares = queryset_familiares.filter(Q(impreso=0) | Q(impreso__isnull=True))
        
        # Filtros por rango de fechas de expedición
        if params.get('fecha_expedicion_desde'):
            queryset = queryset.filter(fecha_expedicion__gte=params['fecha_expedicion_desde'])
            queryset_familiares = queryset_familiares.filter(fecha_expedicion__gte=params['fecha_expedicion_desde'])
        
        if params.get('fecha_expedicion_hasta'):
            queryset = queryset.filter(fecha_expedicion__lte=params['fecha_expedicion_hasta'])
            queryset_familiares = queryset_familiares.filter(fecha_expedicion__lte=params['fecha_expedicion_hasta'])
        
        # Filtros por rango de fechas de vigencia
        if params.get('inicio_vig_desde'):
            queryset = queryset.filter(inicio_vig__gte=params['inicio_vig_desde'])
            queryset_familiares = queryset_familiares.filter(inicio_vig__gte=params['inicio_vig_desde'])
        
        if params.get('fin_vig_hasta'):
            queryset = queryset.filter(fin_vig__lte=params['fin_vig_hasta'])
            queryset_familiares = queryset_familiares.filter(fin_vig__lte=params['fin_vig_hasta'])
        
        # Filtros por fecha de registro
        if params.get('fecha_registro_desde'):
            queryset = queryset.filter(fecha_registro__date__gte=params['fecha_registro_desde'])
            queryset_familiares = queryset_familiares.filter(fecha_registro__date__gte=params['fecha_registro_desde'])
        
        if params.get('fecha_registro_hasta'):
            queryset = queryset.filter(fecha_registro__date__lte=params['fecha_registro_hasta'])
            queryset_familiares = queryset_familiares.filter(fecha_registro__date__lte=params['fecha_registro_hasta'])
        
        # Filtro por fecha de registro específica
        if params.get('fecha_registro'):
            queryset = queryset.filter(fecha_registro__date=params['fecha_registro'])
            queryset_familiares = queryset_familiares.filter(fecha_registro__date=params['fecha_registro'])
        
        # Filtro por fecha de expedición específica
        if params.get('fecha_expedicion'):
            queryset = queryset.filter(fecha_expedicion=params['fecha_expedicion'])
            queryset_familiares = queryset_familiares.filter(fecha_expedicion=params['fecha_expedicion'])
        
        # Filtro por estado de completitud (con foto y firma)
        if params.get('solo_completos') == 'true' or params.get('solo_completos') == True:
            queryset = queryset.exclude(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            )
            queryset_familiares = queryset_familiares.exclude(
                Q(foto__isnull=True) | Q(foto=b'') |
                Q(firma__isnull=True) | Q(firma=b'')
            )
        
        # Filtro por registros activos
        if params.get('solo_activos') == 'true' or params.get('solo_activos') == True:
            queryset = queryset.filter(activo=1)
            queryset_familiares = queryset_familiares.filter(activo=1)

        serializer_enrolamiento = self.get_serializer(queryset, many=True)
        serializer_familiar = EnrolamientoFamiliarSerializer(queryset_familiares, many=True)

        data_enrolamiento = []
        for item in serializer_enrolamiento.data:
            registro = dict(item)
            registro['source_table'] = 'enrolamiento'
            if int(registro.get('nuevo_laredo') or 0) == 1:
                registro['tipo_credencial'] = 'provisional'
            elif int(registro.get('provisional') or 0) == 1:
                registro['tipo_credencial'] = 'anam'
            else:
                registro['tipo_credencial'] = 'enrolamiento'
            data_enrolamiento.append(registro)

        data_familiares = []
        for item in serializer_familiar.data:
            registro = dict(item)
            registro['source_table'] = 'familiar'
            registro['tipo_credencial'] = 'familiar'
            registro['folio'] = registro.get('folio_familiares')
            data_familiares.append(registro)

        data = sorted(
            data_enrolamiento + data_familiares,
            key=lambda x: x.get('fecha_registro') or x.get('fecha_enrolamiento') or '',
            reverse=True
        )

        page = self.paginate_queryset(data)
        if page is not None:
            return self.get_paginated_response(page)

        return Response(data)
    
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


class EnrolamientoFamiliarViewSet(viewsets.ModelViewSet):
    queryset = EnrolamientoFamiliar.objects.all().order_by('-id_enrolamiento')
    serializer_class = EnrolamientoFamiliarSerializer

    filter_backends = [filters.SearchFilter]
    search_fields = ['rfc', 'num_empleado', 'nombre', 'paterno', 'folio_familiares']

    @action(detail=False, methods=['get'], url_path='obtener-folio-maximo-familiares')
    def obtener_folio_maximo_familiares(self, request):
        """
        Retorna el folio máximo actual de familiares con consecutivo independiente.
        Usa el campo folio_familiares y mantiene ceros a la izquierda.
        """
        try:
            folios = EnrolamientoFamiliar.objects.exclude(
                Q(folio_familiares__isnull=True) | Q(folio_familiares='')
            ).values_list('folio_familiares', flat=True)

            if not folios:
                return Response({
                    'status': 'success',
                    'folio_maximo': '000000',
                    'siguiente_folio': '000001'
                }, status=status.HTTP_200_OK)

            folio_max_valor = 0
            folio_max_str = ''
            longitud_formato = 6

            for folio in folios:
                folio_str = str(folio).strip()
                solo_digitos = ''.join(ch for ch in folio_str if ch.isdigit())
                if not solo_digitos:
                    continue

                valor = int(solo_digitos)
                if valor > folio_max_valor:
                    folio_max_valor = valor
                    folio_max_str = folio_str
                    longitud_formato = max(len(folio_str), 6)

            if folio_max_valor == 0:
                return Response({
                    'status': 'success',
                    'folio_maximo': '000000',
                    'siguiente_folio': '000001'
                }, status=status.HTTP_200_OK)

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
                'mensaje': f'Error al obtener folio máximo de familiares: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    @action(detail=True, methods=['post'], url_path='marcar-impreso')
    def marcar_impreso(self, request, pk=None):
        try:
            enrolamiento = self.get_object()

            enrolamiento.impreso = 1
            fecha_expedicion_payload = request.data.get('fecha_expedicion')

            if not enrolamiento.fecha_expedicion:
                enrolamiento.fecha_expedicion = fecha_expedicion_payload or timezone.now().date()
            enrolamiento.save()

            return Response({
                'status': 'success',
                'mensaje': f'Credencial familiar {enrolamiento.folio_familiares or enrolamiento.id_enrolamiento} marcada como impresa',
                'data': {
                    'id_enrolamiento': enrolamiento.id_enrolamiento,
                    'folio_familiares': enrolamiento.folio_familiares,
                    'impreso': enrolamiento.impreso,
                    'fecha_expedicion': enrolamiento.fecha_expedicion
                }
            }, status=status.HTTP_200_OK)

        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al marcar familiar como impreso: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

class SigPagination(PageNumberPagination):
    page_size = 20
    page_size_query_param = 'page_size'
    max_page_size = 200


class SigViewSet(viewsets.ReadOnlyModelViewSet):
    queryset = SicreTblSig.objects.all().order_by('nombres')
    serializer_class = SigSerializer
    pagination_class = SigPagination

    filter_backends = [filters.SearchFilter]
    search_fields = [
        'empleado_anam', 'no_empleado', 'curp', 'nombres',
        'primer_apellido', 'segundo_apellido', 'area', 'cargo',
        'estatus', 'estado_hum', 'estado_nom', 'firma_drh', 'cargo_drh',
    ]

    COLUMN_ALIASES = {
        'EMPLEADO_ANAM': ['EMPLEADO_ANAM', 'EMPLEADO ANAM'],
        'NO_EMPLEADO': ['NO_EMPLEADO', 'NO EMPLEADO', 'NUM_EMPLEADO'],
        'CURP': ['CURP'],
        'NOMBRES': ['NOMBRES', 'NOMBRE'],
        'PRIMER_APELLIDO': ['PRIMER_APELLIDO', 'PRIMER APELLIDO', 'PRIMER_APE', 'PATERNO'],
        'SEGUNDO_APELLIDO': ['SEGUNDO_APELLIDO', 'SEGUNDO APELLIDO', 'SEGUNDO_APE', 'MATERNO'],
        'AREA': ['AREA', 'ADSCRIPCION'],
        'CARGO': ['CARGO', 'PUESTO'],
        'FECHA_EXPEDICION': ['FECHA_EXPEDICION', 'FECHA EXPEDICION'],
        'FIRMA_DRH': ['FIRMA_DRH', 'FIRMA DRH'],
        'CARGO_DRH': ['CARGO_DRH', 'CARGO DRH'],
        'QR': ['QR'],
        'ESTATUS': ['ESTATUS', 'STATUS'],
        'ESTADO_HUM': ['ESTADO_HUM', 'ESTADO HUM'],
        'ESTADO_NOM': ['ESTADO_NOM', 'ESTADO NOM'],
    }

    REQUIRED_COLUMNS = list(COLUMN_ALIASES.keys())

    @staticmethod
    def _normalize_header(value):
        return str(value or '').strip().upper().replace('.', '').replace('-', '_').replace(' ', '_')

    @staticmethod
    def _clean_text(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        text = str(value).strip()
        return text or None

    @staticmethod
    def _clean_identifier(value, upper=False):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None

        if isinstance(value, float) and value.is_integer():
            text = str(int(value))
        else:
            text = str(value).strip()

        if not text:
            return None

        if text.endswith('.0') and text.replace('.', '', 1).isdigit():
            text = text[:-2]

        return text.upper() if upper else text

    @staticmethod
    def _parse_date(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        parsed = pd.to_datetime(value, errors='coerce')
        if pd.isna(parsed):
            return None
        return parsed.date()

    def _build_dataframe(self, archivo):
        if hasattr(archivo, 'seek'):
            archivo.seek(0)

        df = pd.read_excel(archivo)
        if df.empty:
            return df, []

        normalized_map = {self._normalize_header(col): col for col in df.columns}
        selected_columns = {}
        missing = []

        for canonical, aliases in self.COLUMN_ALIASES.items():
            source_col = None
            for alias in aliases:
                normalized_alias = self._normalize_header(alias)
                if normalized_alias in normalized_map:
                    source_col = normalized_map[normalized_alias]
                    break

            if source_col is None:
                missing.append(canonical)
            else:
                selected_columns[canonical] = source_col

        if missing:
            return None, missing

        normalized_df = pd.DataFrame({
            canonical: df[source]
            for canonical, source in selected_columns.items()
        })

        normalized_df['EMPLEADO_ANAM'] = normalized_df['EMPLEADO_ANAM'].apply(lambda x: self._clean_identifier(x) or '')
        normalized_df['CURP'] = normalized_df['CURP'].apply(lambda x: self._clean_identifier(x, upper=True) or '')

        normalized_df = normalized_df[normalized_df['EMPLEADO_ANAM'] != '']
        normalized_df = normalized_df.drop_duplicates(subset=['EMPLEADO_ANAM'], keep='last')

        return normalized_df, []

    def _sig_payload_from_row(self, row):
        return {
            'empleado_anam': self._clean_identifier(row.get('EMPLEADO_ANAM')),
            'no_empleado': self._clean_identifier(row.get('NO_EMPLEADO')),
            'curp': self._clean_identifier(row.get('CURP'), upper=True),
            'nombres': self._clean_text(row.get('NOMBRES')),
            'primer_apellido': self._clean_text(row.get('PRIMER_APELLIDO')),
            'segundo_apellido': self._clean_text(row.get('SEGUNDO_APELLIDO')),
            'area': self._clean_text(row.get('AREA')),
            'cargo': self._clean_text(row.get('CARGO')),
            'fecha_expedicion': self._parse_date(row.get('FECHA_EXPEDICION')),
            'firma_drh': self._clean_text(row.get('FIRMA_DRH')),
            'cargo_drh': self._clean_text(row.get('CARGO_DRH')),
            'qr': self._clean_text(row.get('QR')),
            'estatus': self._clean_text(row.get('ESTATUS')),
            'estado_hum': self._clean_text(row.get('ESTADO_HUM')),
            'estado_nom': self._clean_text(row.get('ESTADO_NOM')),
        }

    @action(detail=False, methods=['POST'], serializer_class=ArchivoExcelSerializer)
    def previsualizar_excel(self, request):
        serializer = ArchivoExcelSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)

        try:
            archivo = serializer.validated_data['archivo']
            df, missing = self._build_dataframe(archivo)

            if missing:
                return Response({
                    'status': 'error',
                    'mensaje': 'El archivo no contiene todas las columnas requeridas de SIG.',
                    'columnas_faltantes': missing
                }, status=status.HTTP_400_BAD_REQUEST)

            total_registros = 0 if df is None else len(df)
            preview_limit = min(max(int(request.query_params.get('preview_limit', 200)), 1), 500)
            preview_df = df.head(preview_limit) if df is not None else pd.DataFrame()

            registros_preview = []
            for _, row in preview_df.iterrows():
                registro = self._sig_payload_from_row(row)
                registros_preview.append(registro)

            return Response({
                'status': 'success',
                'mensaje': 'Vista previa generada correctamente',
                'total_registros': total_registros,
                'registros': registros_preview,
                'preview_limit': preview_limit,
                'preview_registros': len(registros_preview),
                'preview_parcial': total_registros > len(registros_preview)
            }, status=status.HTTP_200_OK)

        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al leer el archivo: {str(e)}'
            }, status=status.HTTP_400_BAD_REQUEST)

    @action(detail=False, methods=['POST'], serializer_class=ArchivoExcelSerializer)
    def subir_excel(self, request):
        serializer = ArchivoExcelSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)

        try:
            archivo = serializer.validated_data['archivo']
            df, missing = self._build_dataframe(archivo)

            if missing:
                return Response({
                    'status': 'error',
                    'mensaje': 'El archivo no contiene todas las columnas requeridas de SIG.',
                    'columnas_faltantes': missing
                }, status=status.HTTP_400_BAD_REQUEST)

            if df is None or df.empty:
                return Response({
                    'status': 'error',
                    'mensaje': 'No hay registros válidos para procesar.'
                }, status=status.HTTP_400_BAD_REQUEST)

            rows = [self._sig_payload_from_row(row) for _, row in df.iterrows()]

            # --- Validación de duplicados en el archivo ---
            curp_count = {}
            num_emp_count = {}
            for r in rows:
                c = (r.get('curp') or '').strip()
                n = (r.get('no_empleado') or '').strip()
                if c:
                    curp_count[c] = curp_count.get(c, 0) + 1
                if n:
                    num_emp_count[n] = num_emp_count.get(n, 0) + 1

            curps_dup = [c for c, cnt in curp_count.items() if cnt > 1]
            nums_dup  = [n for n, cnt in num_emp_count.items() if cnt > 1]

            if curps_dup or nums_dup:
                errores = {}
                if curps_dup:
                    errores['curps_duplicadas'] = curps_dup
                if nums_dup:
                    errores['numeros_empleado_duplicados'] = nums_dup
                return Response({
                    'status': 'error',
                    'mensaje': 'El archivo contiene registros duplicados. No puede haber más de un empleado con la misma CURP o el mismo número de empleado.',
                    **errores
                }, status=status.HTTP_400_BAD_REQUEST)
            # --- Fin validación duplicados ---

            rows_by_empleado = {}
            for row in rows:
                empleado_key = row.get('empleado_anam')
                if not empleado_key:
                    continue
                rows_by_empleado[empleado_key] = row

            rows = list(rows_by_empleado.values())
            empleados_anam = [row['empleado_anam'] for row in rows if row.get('empleado_anam')]
            curps = [row['curp'] for row in rows if row.get('curp')]

            sig_create = []
            sig_update = []
            enr_create = []
            enr_update = []

            sig_existing = SicreTblSig.objects.in_bulk(empleados_anam)
            sig_existing_by_curp_qs = SicreTblSig.objects.filter(curp__in=curps)
            sig_existing_by_curp = {}
            for obj in sig_existing_by_curp_qs:
                if obj.curp and obj.curp not in sig_existing_by_curp:
                    sig_existing_by_curp[obj.curp] = obj

            enr_existing_qs = Enrolamiento.objects.filter(curp__in=curps).order_by('-id_enrolamiento')
            enr_existing = {}
            for obj in enr_existing_qs:
                if obj.curp and obj.curp not in enr_existing:
                    enr_existing[obj.curp] = obj

            max_id = SicreTblSig.objects.aggregate(max_id=models.Max('id')).get('max_id') or 0
            next_id = int(max_id)
            now_ts = timezone.now()

            def normalized_text(value):
                if value is None:
                    return ''
                return str(value).strip()

            def same_text(current, incoming):
                return normalized_text(current) == normalized_text(incoming)

            def same_date(current, incoming):
                return current == incoming

            for row in rows:
                empleado_anam = row['empleado_anam']
                curp = row.get('curp')
                if not empleado_anam:
                    continue

                sig_obj = sig_existing.get(empleado_anam)
                if not sig_obj and curp:
                    sig_obj = sig_existing_by_curp.get(curp)
                if sig_obj:
                    sig_changed = False

                    if sig_obj.id is None:
                        next_id += 1
                        sig_obj.id = next_id
                        sig_changed = True

                    sig_new_values = {
                        'no_empleado': row['no_empleado'],
                        'curp': row['curp'],
                        'nombres': row['nombres'],
                        'primer_apellido': row['primer_apellido'],
                        'segundo_apellido': row['segundo_apellido'],
                        'area': row['area'],
                        'cargo': row['cargo'],
                        'fecha_expedicion': row['fecha_expedicion'],
                        'firma_drh': row['firma_drh'],
                        'cargo_drh': row['cargo_drh'],
                        'qr': row['qr'],
                        'estatus': row['estatus'],
                        'estado_hum': row['estado_hum'],
                        'estado_nom': row['estado_nom'],
                    }

                    for field_name, new_value in sig_new_values.items():
                        current_value = getattr(sig_obj, field_name)
                        is_same = same_date(current_value, new_value) if field_name == 'fecha_expedicion' else same_text(current_value, new_value)
                        if not is_same:
                            setattr(sig_obj, field_name, new_value)
                            sig_changed = True

                    if sig_changed:
                        sig_obj.fecha_actualizacion = now_ts
                        sig_update.append(sig_obj)
                else:
                    next_id += 1
                    sig_create.append(SicreTblSig(
                        id=next_id,
                        fecha_actualizacion=now_ts,
                        **row
                    ))

                nombre = (row.get('nombres') or '').strip()
                paterno = (row.get('primer_apellido') or '').strip()
                materno = (row.get('segundo_apellido') or '').strip()
                apellidos = f"{paterno} {materno}".strip()
                estatus = (row.get('estatus') or '').lower()
                activo_flag = 0 if 'baja' in estatus else 1
                rfc_ref = (row.get('empleado_anam') or row.get('no_empleado') or curp or '').strip()

                if not curp:
                    continue

                enr_obj = enr_existing.get(curp)
                if enr_obj:
                    enr_changed = False
                    enr_new_values = {
                        'rfc': rfc_ref,
                        'num_empleado': row.get('no_empleado'),
                        'nombre': nombre,
                        'paterno': paterno,
                        'materno': materno,
                        'apellidos': apellidos,
                        'puesto': row.get('cargo'),
                        'adscripcion': row.get('area'),
                        'activo': activo_flag,
                        'nivel_credencial': _cargo_a_nivel(row.get('cargo') or ''),
                        'layout_credencial': 'ANAM_2025',
                    }

                    for field_name, new_value in enr_new_values.items():
                        current_value = getattr(enr_obj, field_name)
                        is_same = current_value == new_value if field_name == 'activo' else same_text(current_value, new_value)
                        if not is_same:
                            setattr(enr_obj, field_name, new_value)
                            enr_changed = True

                    if enr_changed:
                        enr_obj.fecha_modificacion = timezone.now()
                        enr_update.append(enr_obj)
                else:
                    enr_create.append(Enrolamiento(
                        rfc=rfc_ref,
                        num_empleado=row.get('no_empleado'),
                        curp=curp,
                        nombre=nombre,
                        paterno=paterno,
                        materno=materno,
                        apellidos=apellidos,
                        puesto=row.get('cargo'),
                        adscripcion=row.get('area'),
                        activo=activo_flag,
                        fecha_enrolamiento=timezone.now(),
                        nivel_credencial=_cargo_a_nivel(row.get('cargo') or ''),
                        layout_credencial='ANAM_2025'
                    ))

            sig_update = list({obj.empleado_anam: obj for obj in sig_update}.values())
            enr_update = list({obj.id_enrolamiento: obj for obj in enr_update}.values())

            if not sig_create and not sig_update and not enr_create and not enr_update:
                resumen_sin_cambios = {
                    'registros_archivo': len(rows),
                    'sig_creados': 0,
                    'sig_actualizados': 0,
                    'enrolamiento_creados': 0,
                    'enrolamiento_actualizados': 0
                }
                return Response({
                    'status': 'success',
                    'mensaje': 'No se encontraron nuevos empleados ni actualizaciones de registros.',
                    'resumen': resumen_sin_cambios
                }, status=status.HTTP_200_OK)

            with transaction.atomic():
                if sig_create:
                    SicreTblSig.objects.bulk_create(sig_create, batch_size=1000)
                if sig_update:
                    SicreTblSig.objects.bulk_update(
                        sig_update,
                        ['id', 'no_empleado', 'curp', 'nombres', 'primer_apellido', 'segundo_apellido', 'area', 'cargo', 'fecha_expedicion', 'firma_drh', 'cargo_drh', 'qr', 'estatus', 'estado_hum', 'estado_nom', 'fecha_actualizacion'],
                        batch_size=1000
                    )

                if enr_create:
                    Enrolamiento.objects.bulk_create(enr_create, batch_size=1000)
                if enr_update:
                    Enrolamiento.objects.bulk_update(
                        enr_update,
                        ['rfc', 'num_empleado', 'nombre', 'paterno', 'materno', 'apellidos', 'puesto', 'adscripcion', 'activo', 'nivel_credencial', 'layout_credencial', 'fecha_modificacion'],
                        batch_size=1000
                    )

            resumen = {
                'registros_archivo': len(rows),
                'sig_creados': len(sig_create),
                'sig_actualizados': len(sig_update),
                'enrolamiento_creados': len(enr_create),
                'enrolamiento_actualizados': len(enr_update),
                'detalle_empleados': [
                    {
                        'no_empleado': obj.no_empleado or '',
                        'nombre': f"{obj.nombres or ''} {obj.primer_apellido or ''} {obj.segundo_apellido or ''}".strip(),
                        'tipo': 'Nuevo ingreso'
                    }
                    for obj in sig_create
                ] + [
                    {
                        'no_empleado': obj.no_empleado or '',
                        'nombre': f"{obj.nombres or ''} {obj.primer_apellido or ''} {obj.segundo_apellido or ''}".strip(),
                        'tipo': 'Actualización'
                    }
                    for obj in sig_update
                ]
            }

            sin_cambios = (
                resumen['sig_creados'] == 0 and
                resumen['sig_actualizados'] == 0 and
                resumen['enrolamiento_creados'] == 0 and
                resumen['enrolamiento_actualizados'] == 0
            )

            return Response({
                'status': 'success',
                'mensaje': 'No se encontraron nuevos empleados ni actualizaciones de registros.' if sin_cambios else 'Consulta de SIG actualizada exitosamente.',
                'resumen': resumen
            }, status=status.HTTP_200_OK if sin_cambios else status.HTTP_201_CREATED)

        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error procesando el archivo: {str(e)}'
            }, status=status.HTTP_400_BAD_REQUEST)

    @action(detail=False, methods=['get'], url_path='historial-cargas')
    def historial_cargas(self, request):
        try:
            resumen_estatus = list(
                SicreTblSig.objects.values('estatus').annotate(total=Count('empleado_anam')).order_by('-total')[:10]
            )
            preview = list(
                SicreTblSig.objects.values('empleado_anam', 'curp', 'nombres', 'primer_apellido', 'segundo_apellido', 'fecha_actualizacion')[:10]
            )

            return Response({
                'status': 'success',
                'total_registros': SicreTblSig.objects.count(),
                'por_estatus': resumen_estatus,
                'preview': preview
            }, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({
                'status': 'error',
                'mensaje': f'Error al obtener historial: {str(e)}'
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    @action(detail=False, methods=['get'], url_path='detalle-lote/(?P<lote_id>[0-9]+)')
    def detalle_lote(self, request, lote_id=None):
        return Response({
            'status': 'error',
            'mensaje': 'Detalle por lote ya no aplica en el nuevo esquema SIG.'
        }, status=status.HTTP_410_GONE)


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


@api_view(['GET'])
def foto_firma_empleado(request, emplid):
    """
    Obtiene foto y firma de safirho_db.NW_EMPL_FOTO_ANAM para un EMPLID dado.
    Prueba con el valor tal cual, zero-padded a 11 dígitos y sin ceros a la izquierda.
    """
    emplid_str = str(emplid).strip()
    emplid_padded = emplid_str.zfill(11)
    emplid_stripped = emplid_str.lstrip('0') or '0'

    variantes = list({emplid_str, emplid_padded, emplid_stripped})
    placeholders = ','.join(['%s'] * len(variantes))

    try:
        with connection.cursor() as cursor:
            cursor.execute(
                f"SELECT FOTO, FIRMA, TipoMime FROM safirho_db.NW_EMPL_FOTO_ANAM WHERE EMPLID IN ({placeholders}) LIMIT 1",
                variantes
            )
            row = cursor.fetchone()

        if not row:
            return Response({'foto': None, 'firma': None, 'encontrado': False})

        foto_raw, firma_raw, tipo_mime = row
        mime = tipo_mime or 'image/jpeg'

        def blob_to_b64(raw):
            if raw is None:
                return None
            if isinstance(raw, memoryview):
                raw = bytes(raw)
            elif isinstance(raw, bytearray):
                raw = bytes(raw)
            elif not isinstance(raw, bytes):
                raw = bytes(raw)
            if not raw:
                return None
            return f'data:{mime};base64,' + base64.b64encode(raw).decode('utf-8')

        return Response({
            'foto': blob_to_b64(foto_raw),
            'firma': blob_to_b64(firma_raw),
            'encontrado': True,
        })

    except Exception as e:
        return Response({'foto': None, 'firma': None, 'encontrado': False, 'error': str(e)}, status=200)