from django.db import models
import re
import unicodedata

from . import media_utils


class EnrolamientoCredencial(models.Model):
    """
    Registro de "qué pasó al expedir una credencial" -- NO es una copia de los
    datos del empleado. Esos (nombre, curp, puesto, adscripcion, etc.) ya viven
    en `sicre_tbl_sig` (sincronizada cada 30 min por Celery en el proyecto
    Control_De_Plazas_Backend) y se consultan ahi mismo cruzando por
    `num_empleado` <-> `sicre_tbl_sig.NO_EMPLEADO` -- guardarlos aqui tambien
    los duplicaria y los dejaria obsoletos en cuanto el roster cambie.

    Tampoco guarda rutas de foto/firma: se resuelven en el momento por
    convencion de nombre de archivo (`media_utils.resolver_foto/resolver_firma`,
    buscan `<num_empleado>.<ext>` dentro de MEDIA_ROOT). Guardar la ruta aqui
    la duplicaria innecesariamente y podria desincronizarse si el archivo se
    vuelve a capturar.
    """
    id_enrolamiento = models.AutoField(primary_key=True)
    # Llave de cruce contra sicre_tbl_sig.NO_EMPLEADO (forma corta, sin ceros a
    # la izquierda -- la misma que usan los ~14,300 archivos de /media).
    num_empleado = models.CharField(max_length=20, blank=True, null=True, db_index=True)

    fin_vig = models.DateField(blank=True, null=True)
    folio = models.CharField(max_length=50, blank=True, null=True)
    fecha_expedicion = models.DateField(blank=True, null=True)
    provisional = models.IntegerField(db_column='provisional', blank=True, null=True)

    # Clave (PlantillaCredencial.clave) de la plantilla con la que se expidio
    # esta credencial -- texto, no FK, por consistencia con el resto del
    # sistema (ver CLAUDE.md, gotcha #9: sin FKs, cruce por valor). Permite
    # reimprimir con la misma plantilla sin que el enrolador tenga que
    # volver a elegirla.
    # Clave de la PlantillaCredencial usada. Se guarda la CLAVE y no un FK
    # (el modelo de datos no usa FKs, ver CLAUDE.md) ni una copia del diseno:
    # lo que interesa es "con que plantilla se le imprimio a esta persona".
    plantilla_credencial = models.CharField(max_length=50, blank=True, null=True)

    # Copia EXACTA del lienzo impreso, SIEMPRE -- no solo cuando hubo ajustes.
    #
    # Guardar unicamente la clave de la plantilla no basta para auditar: las
    # plantillas son editables desde `/plantillas`, asi que reconstruir una
    # credencial de hace seis meses con la plantilla de hoy mostraria un
    # documento que nunca existio. El snapshot congela texto, posiciones,
    # tipografias y fondo tal como salieron impresos.
    #
    # Las imagenes (foto, firma, fondo) NO se embeben en el JSON: se archivan
    # aparte en `media/historico/` nombradas por el hash de su contenido, y
    # aqui solo queda la referencia. Ver `media_utils.archivar_medio`. Medido:
    # embeberlas costaria ~250 KB por fila (~4 GB en 16 400 impresiones)
    # contra ~47 KB referenciandolas, y el archivo deduplica las reimpresiones.
    #
    # Los objetos conservan su `data.binding`, asi que al REIMPRIMIR se vuelven
    # a poblar con datos frescos (folio y fecha nuevos) sin perder las
    # posiciones ajustadas. La auditoria, en cambio, dibuja el JSON tal cual,
    # sin repoblar nada.
    canvas_frente = models.JSONField(blank=True, null=True)
    canvas_reverso = models.JSONField(blank=True, null=True)

    # ¿El operador movio/edito algo a mano antes de imprimir?
    #
    # Antes esto se deducia de "hay canvas guardado", pero desde que el
    # snapshot se guarda SIEMPRE esa señal dejo de existir y hubo que hacerla
    # explicita. Importa porque "Imprimir credenciales" solo restaura el
    # lienzo guardado cuando hubo ajustes: si lo restaurara siempre, una
    # credencial impresa limpia dejaria de seguir a su plantilla y los
    # cambios de diseño nunca llegarian a las reimpresiones.
    con_ajustes = models.BooleanField(default=False)

    # Auditoria
    fecha_registro = models.DateTimeField(auto_now_add=True, blank=True, null=True)
    id_usuario_registra = models.IntegerField(blank=True, null=True)
    fecha_modificacion = models.DateTimeField(blank=True, null=True)
    id_usuario_modifica = models.IntegerField(blank=True, null=True)
    activo = models.IntegerField(blank=True, null=True)

    class Meta:
        managed = True
        db_table = 'sicre_tbl_enrolamiento_credencial'
        # Una fila POR IMPRESION: el historial es cronologico y auditable.
        ordering = ['-fecha_registro']
        verbose_name = 'Enrolamiento (credencial)'
        verbose_name_plural = 'Enrolamientos (credencial)'

    def __str__(self):
        return f"{self.num_empleado or 's/n'} (folio {self.folio or 's/n'})"

    # ---- Ayudas de medios (resueltas por convencion, nunca persistidas) ----

    @property
    def foto_url(self):
        ruta = media_utils.resolver_foto(self.num_empleado)
        return media_utils.url_publica(ruta) if ruta else None

    @property
    def firma_url(self):
        ruta = media_utils.resolver_firma(self.num_empleado)
        return media_utils.url_publica(ruta) if ruta else None


class PlantillaCredencial(models.Model):
    """
    Plantilla de credencial disenada en el editor tipo canvas.

    `canvas_frente` / `canvas_reverso` guardan el JSON nativo de Fabric.js, de modo
    que editar, previsualizar e imprimir usen exactamente las mismas coordenadas.
    """
    id_plantilla = models.AutoField(primary_key=True)
    clave = models.CharField(max_length=50, unique=True)
    nombre = models.CharField(max_length=150)
    descripcion = models.CharField(max_length=255, blank=True, null=True)

    # Rutas relativas en MEDIA_ROOT (carpeta plantillas/).
    fondo_frente = models.CharField(max_length=255, blank=True, null=True)
    fondo_reverso = models.CharField(max_length=255, blank=True, null=True)

    # Serializacion Fabric.js de cada cara.
    canvas_frente = models.JSONField(blank=True, null=True)
    canvas_reverso = models.JSONField(blank=True, null=True)

    # Espacio de diseno en px (debe coincidir editor / preview / PDF).
    ancho_px = models.IntegerField(default=638)
    alto_px = models.IntegerField(default=1016)

    # Tamano fisico de impresion en milimetros (CR80 vertical por defecto).
    ancho_mm = models.DecimalField(max_digits=6, decimal_places=2, default=54)
    alto_mm = models.DecimalField(max_digits=6, decimal_places=2, default=86)

    activo = models.BooleanField(default=True)
    # Plantilla que se carga por omision en "Imprimir credenciales". La
    # exclusividad (solo una en true) la garantiza el action
    # `marcar-por-defecto` del viewset, no una constraint de BD: MySQL no
    # soporta indices unicos parciales (WHERE por_defecto = 1), y un UNIQUE
    # normal impediria tener varias plantillas en false.
    por_defecto = models.BooleanField(default=False)
    fecha_registro = models.DateTimeField(auto_now_add=True)
    fecha_modificacion = models.DateTimeField(auto_now=True)
    id_usuario_registra = models.IntegerField(blank=True, null=True)
    id_usuario_modifica = models.IntegerField(blank=True, null=True)

    class Meta:
        managed = True
        db_table = 'sicre_tbl_plantilla_credencial'
        verbose_name = 'Plantilla de credencial'
        verbose_name_plural = 'Plantillas de credencial'
        ordering = ['nombre']

    def __str__(self):
        return f"{self.clave} - {self.nombre}"

    @property
    def fondo_frente_url(self):
        return media_utils.url_publica(self.fondo_frente)

    @property
    def fondo_reverso_url(self):
        return media_utils.url_publica(self.fondo_reverso)


class Enrolamiento(models.Model):
    id_enrolamiento = models.AutoField(primary_key=True)
    num_empleado = models.CharField(max_length=20, blank=True, null=True)
    rfc = models.CharField(max_length=18)
    curp = models.CharField(max_length=18, blank=True, null=True)
    nombre = models.CharField(max_length=100, blank=True, null=True)
    paterno = models.CharField(max_length=100, blank=True, null=True)
    materno = models.CharField(max_length=100, blank=True, null=True)
    apellidos = models.CharField(max_length=200, blank=True, null=True)
    puesto = models.CharField(max_length=100, blank=True, null=True)
    adscripcion = models.CharField(max_length=100, blank=True, null=True)
    inicio_vig = models.DateField(blank=True, null=True)
    fin_vig = models.DateField(blank=True, null=True)
    eladia = models.CharField(max_length=100, blank=True, null=True)
    
    foto = models.BinaryField(blank=True, null=True)
    firma = models.BinaryField(blank=True, null=True)
    
    folio = models.CharField(max_length=50, blank=True, null=True)
    impreso = models.IntegerField(blank=True, null=True)
    fecha_expedicion = models.DateField(blank=True, null=True)
    provisional = models.IntegerField(db_column='provisional', blank=True, null=True)
    id_usuario_registra = models.IntegerField(blank=True, null=True)
    fecha_registro = models.DateTimeField(auto_now_add=True, blank=True, null=True)
    fecha_enrolamiento = models.DateTimeField(blank=True, null=True)
    id_usuario_modifica = models.IntegerField(blank=True, null=True)
    fecha_modificacion = models.DateTimeField(blank=True, null=True)
    ultima_carga = models.DateTimeField(auto_now=True, blank=True, null=True)
    activo = models.IntegerField(blank=True, null=True)
    nuevo_laredo = models.IntegerField(blank=True, null=True)
    nivel_credencial = models.CharField(max_length=50, blank=True, null=True)
    layout_credencial = models.CharField(max_length=30, blank=True, null=True)

    class Meta:
        managed = True
        db_table = 'sicre_tbl_enrolamiento'
        verbose_name = 'Expediente'
        verbose_name_plural = 'Expedientes de Enrolamiento'

    def __str__(self):
        return f"{self.rfc} - {self.nombre}"


class EnrolamientoFamiliar(models.Model):
    id_enrolamiento = models.AutoField(primary_key=True)
    num_empleado = models.CharField(max_length=20, blank=True, null=True)
    rfc = models.CharField(max_length=18)
    curp = models.CharField(max_length=18, blank=True, null=True)
    nombre = models.CharField(max_length=100, blank=True, null=True)
    paterno = models.CharField(max_length=100, blank=True, null=True)
    materno = models.CharField(max_length=100, blank=True, null=True)
    apellidos = models.CharField(max_length=200, blank=True, null=True)
    puesto = models.CharField(max_length=100, blank=True, null=True)
    adscripcion = models.CharField(max_length=100, blank=True, null=True)
    inicio_vig = models.DateField(blank=True, null=True)
    fin_vig = models.DateField(blank=True, null=True)
    eladia = models.CharField(max_length=100, blank=True, null=True)

    foto = models.BinaryField(blank=True, null=True)
    firma = models.BinaryField(blank=True, null=True)

    folio = models.CharField(max_length=50, blank=True, null=True)
    folio_familiares = models.CharField(max_length=50, blank=True, null=True)
    impreso = models.IntegerField(blank=True, null=True)
    fecha_expedicion = models.DateField(blank=True, null=True)
    provisional = models.IntegerField(db_column='provisional', blank=True, null=True)
    nombre_reverso = models.CharField(max_length=200, blank=True, null=True)
    vivienda = models.CharField(max_length=100, blank=True, null=True)
    marca = models.CharField(max_length=100, blank=True, null=True)
    color = models.CharField(max_length=100, blank=True, null=True)
    modelo = models.CharField(max_length=100, blank=True, null=True)
    placas = models.CharField(max_length=100, blank=True, null=True)
    id_usuario_registra = models.IntegerField(blank=True, null=True)
    fecha_registro = models.DateTimeField(auto_now_add=True, blank=True, null=True)
    fecha_enrolamiento = models.DateTimeField(blank=True, null=True)
    id_usuario_modifica = models.IntegerField(blank=True, null=True)
    fecha_modificacion = models.DateTimeField(blank=True, null=True)
    ultima_carga = models.DateTimeField(auto_now=True, blank=True, null=True)
    activo = models.IntegerField(blank=True, null=True)
    nuevo_laredo = models.IntegerField(blank=True, null=True)

    class Meta:
        managed = True
        db_table = 'sicre_tbl_enrolamiento_familiar'
        verbose_name = 'Expediente Familiar'
        verbose_name_plural = 'Expedientes de Enrolamiento Familiar'

    def __str__(self):
        return f"{self.rfc} - {self.nombre}"
    


class SicreTblSig(models.Model):
    id = models.BigIntegerField(db_column='ID', blank=True, null=True, db_index=True)
    empleado_anam = models.CharField(db_column='EMPLEADO_ANAM', max_length=50, primary_key=True)
    no_empleado = models.CharField(db_column='NO_EMPLEADO', max_length=50, blank=True, null=True)
    curp = models.CharField(db_column='CURP', max_length=18, blank=True, null=True)
    nombres = models.CharField(db_column='NOMBRES', max_length=150, blank=True, null=True)
    primer_apellido = models.CharField(db_column='PRIMER_APELLIDO', max_length=100, blank=True, null=True)
    segundo_apellido = models.CharField(db_column='SEGUNDO_APELLIDO', max_length=100, blank=True, null=True)
    area = models.CharField(db_column='AREA', max_length=150, blank=True, null=True)
    cargo = models.CharField(db_column='CARGO', max_length=150, blank=True, null=True)
    fecha_expedicion = models.DateField(db_column='FECHA_EXPEDICION', blank=True, null=True)
    firma_drh = models.CharField(db_column='FIRMA_DRH', max_length=255, blank=True, null=True)
    cargo_drh = models.CharField(db_column='CARGO_DRH', max_length=255, blank=True, null=True)
    qr = models.TextField(db_column='QR', blank=True, null=True)
    estatus = models.CharField(db_column='ESTATUS', max_length=100, blank=True, null=True)
    estado_hum = models.CharField(db_column='ESTADO_HUM', max_length=100, blank=True, null=True)
    estado_nom = models.CharField(db_column='ESTADO_NOM', max_length=100, blank=True, null=True)
    # Se reescribe en CADA sincronizacion (cada 30 min), para TODOS los
    # registros -- no distingue "recien llegado" de "ya existia y se
    # refresco de rutina". Para eso ver fecha_primera_deteccion.
    fecha_actualizacion = models.DateTimeField(db_column='FECHA_ACTUALIZACION', blank=True, null=True)
    # A diferencia de fecha_actualizacion, esta SOLO se escribe la primera vez
    # que se ve a este empleado_anam; en sincronizaciones posteriores el
    # proyecto Control_De_Plazas_Backend la conserva tal cual (no la
    # sobreescribe). Es lo que permite distinguir altas reales de refrescos
    # rutinarios -- ver Control_De_Plazas_Backend/plantilla/tasks.py,
    # _obtener_fechas_primera_deteccion / _leer_csv_poblado_credenciales.
    fecha_primera_deteccion = models.DateTimeField(
        db_column='FECHA_PRIMERA_DETECCION', blank=True, null=True
    )

    class Meta:
        managed = True
        db_table = 'sicre_tbl_sig'
        verbose_name = 'SICRE-sig'
        verbose_name_plural = 'SICRE-SIG'

    def __str__(self):
        return f"{self.empleado_anam} - {self.nombres}"

class CargaMasiva(models.Model):
    # ==========================================
    # 1. CAMPOS ORIGINALES DE CARGA MASIVA
    # ==========================================
    id = models.AutoField(primary_key=True)
    rfc = models.CharField(max_length=18)
    nombre = models.CharField(max_length=200, blank=True, null=True)
    foto = models.BinaryField(blank=True, null=True)
    firma = models.BinaryField(blank=True, null=True)
    lote = models.CharField(max_length=100)
    fecha_enrolamiento = models.DateTimeField(auto_now_add=True)
    usuario_enrola = models.IntegerField(blank=True, null=True)
    activo = models.BooleanField(default=True) # Se mantiene Booleano (Soft Delete)

    # ==========================================
    # 2. CAMPOS HEREDADOS DE SICRE_TBL_SIG
    # ==========================================
    empleado_anam = models.CharField(max_length=50, blank=True, null=True)
    no_empleado = models.CharField(max_length=50, blank=True, null=True)
    curp = models.CharField(max_length=18, blank=True, null=True)
    nombres = models.CharField(max_length=150, blank=True, null=True)
    primer_apellido = models.CharField(max_length=100, blank=True, null=True)
    segundo_apellido = models.CharField(max_length=100, blank=True, null=True)
    area = models.CharField(max_length=150, blank=True, null=True)
    cargo = models.CharField(max_length=150, blank=True, null=True)
    fecha_expedicion = models.DateField(blank=True, null=True) # Compartido con Enrolamiento
    firma_drh = models.CharField(max_length=255, blank=True, null=True)
    cargo_drh = models.CharField(max_length=255, blank=True, null=True)
    qr = models.TextField(blank=True, null=True)
    estatus = models.CharField(max_length=100, blank=True, null=True)
    estado_hum = models.CharField(max_length=100, blank=True, null=True)
    estado_nom = models.CharField(max_length=100, blank=True, null=True)

    # ==========================================
    # 3. CAMPOS HEREDADOS DE ENROLAMIENTO
    # ==========================================
    inicio_vig = models.DateField(blank=True, null=True)
    fin_vig = models.DateField(blank=True, null=True)
    folio = models.CharField(max_length=50, blank=True, null=True)
    impreso = models.IntegerField(blank=True, null=True)
    provisional = models.IntegerField(db_column='provisional', blank=True, null=True)
    id_usuario_registra = models.IntegerField(blank=True, null=True)
    fecha_registro = models.DateTimeField(auto_now_add=True, blank=True, null=True)
    id_usuario_modifica = models.IntegerField(blank=True, null=True)
    fecha_modificacion = models.DateTimeField(blank=True, null=True)
    nuevo_laredo = models.IntegerField(blank=True, null=True)
    nivel_credencial = models.CharField(max_length=50, blank=True, null=True)
    layout_credencial = models.CharField(max_length=30, blank=True, null=True)

    class Meta:
        db_table = 'sicre_tbl_carga_masiva'

class ConsecutivoFolio(models.Model):
    """
    Contador de folios de credencial, persistido en servidor.

    Vive en BD y no en el navegador a proposito: el folio debe ser unico entre
    TODAS las estaciones de impresion. Guardarlo en localStorage haria que dos
    capturistas trabajando en paralelo emitieran credenciales con el mismo
    folio sin enterarse.

    `clave` permite mas de un contador por si en el futuro se separan series
    (ANAM / Nuevo Laredo / familiares), aunque hoy solo se usa 'credencial'.
    `valor` es el PROXIMO folio a emitir, no el ultimo emitido.
    """
    id_consecutivo = models.AutoField(primary_key=True)
    clave = models.CharField(max_length=50, unique=True, default='credencial')
    valor = models.IntegerField(default=1)
    # Ceros a la izquierda al formatear: 6 -> '000123'.
    longitud = models.IntegerField(default=6)
    fecha_modificacion = models.DateTimeField(auto_now=True)
    id_usuario_modifica = models.IntegerField(blank=True, null=True)

    class Meta:
        managed = True
        db_table = 'sicre_tbl_consecutivo_folio'

    def __str__(self):
        return f'{self.clave}: {self.formateado()}'

    def formateado(self, valor=None) -> str:
        numero = self.valor if valor is None else valor
        return str(numero).zfill(self.longitud or 6)


class UnidadAdministrativa(models.Model):
    """
    Catalogo de unidades administrativas con su nombre corto, para imprimir
    en la credencial.

    El campo `area` del roster SIG trae el nombre oficial completo, que llega
    a 90 caracteres ("Aduana del Aeropuerto Internacional de la Ciudad de
    Mexico con sede en la Ciudad de Mexico") y desborda o tapa otros campos
    del gafete. Aqui se guarda su equivalente compacto ("AICM").

    `nombre_normalizado` es la clave real de cruce y NO es redundante: el
    roster escribe el area en MAYUSCULAS ("UNIDAD DE ADMINISTRACION Y
    FINANZAS") mientras que el catalogo de origen viene en Title Case
    ("Unidad de Administracion y Finanzas"). Comprobado contra los datos
    reales: cruzando por el texto crudo empatan 0 de 16 431 empleados;
    normalizando, empatan los 16 431.
    """
    id_unidad = models.AutoField(primary_key=True)
    nombre = models.CharField(max_length=255)
    nombre_normalizado = models.CharField(max_length=255, unique=True, db_index=True)
    # Texto EXACTO que se imprime en la credencial. Las Direcciones Generales
    # llevan su acronimo (DGTI, UAF...) y las aduanas llevan 'ANAM'. Se guarda
    # ya resuelto, sin banderas ni reglas en codigo, para que el catalogo sea
    # editable desde la pantalla de Catalogos: si manana quieren que una
    # aduana imprima su propio nombre, lo escriben ahi y funciona, sin tocar
    # el sistema.
    nombre_compactado = models.CharField(max_length=100)
    activo = models.BooleanField(default=True)
    fecha_registro = models.DateTimeField(auto_now_add=True)
    fecha_modificacion = models.DateTimeField(auto_now=True)

    class Meta:
        managed = True
        db_table = 'sicre_cat_unidad_compactada'
        ordering = ['nombre']

    # Valor por omision para las unidades que no son Direccion General.
    NOMBRE_GENERICO = 'ANAM'

    def __str__(self):
        return f'{self.nombre} -> {self.nombre_compactado}'

    @classmethod
    def compactar(cls, area) -> str:
        """
        Nombre corto que se imprime para un area del roster.

        Si el area no esta en el catalogo devuelve el nombre completo: texto
        largo y apretado en la credencial es preferible a un campo vacio, y
        el hueco se detecta aparte en la pantalla de Catalogos.
        """
        texto = str(area or '').strip()
        if not texto:
            return ''
        registro = cls.objects.filter(
            nombre_normalizado=cls.normalizar(texto), activo=True
        ).only('nombre_compactado').first()
        return registro.nombre_compactado if registro else texto

    @staticmethod
    def normalizar(valor) -> str:
        """
        Clave de cruce: mayusculas, sin acentos y sin puntuacion.

        Quitar los acentos ademas de subir a mayusculas cubre las variantes
        del acervo ("MEXICO" vs "MEXICO"), y colapsar todo lo que no sea
        alfanumerico absorbe dobles espacios y puntos finales, que aparecen
        de forma inconsistente entre ambas fuentes.
        """
        texto = unicodedata.normalize('NFD', str(valor or '').upper())
        texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
        return re.sub(r'[^A-Z0-9]+', ' ', texto).strip()

    def save(self, *args, **kwargs):
        self.nombre_normalizado = self.normalizar(self.nombre)
        super().save(*args, **kwargs)
