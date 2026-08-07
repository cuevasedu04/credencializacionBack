"""
Utilidades para manejar fotos, firmas y fondos de plantillas en MEDIA_ROOT.

Reemplaza el esquema anterior donde foto/firma se guardaban como BLOB en MySQL
(o como flag de 1 byte apuntando a safirho_db). Ahora el archivo vive en disco
bajo MEDIA_ROOT y la base de datos solo guarda la ruta relativa.

El acervo historico ya existente esta nombrado por num_empleado con extensiones
heterogeneas (fotos: jpg/JPG/jpeg/png, firmas: png/PNG), por lo que la resolucion
de un archivo existente prueba varias extensiones antes de rendirse.
"""
import base64
import binascii
import hashlib
import os
import re
from pathlib import Path
from urllib.parse import unquote, urlsplit

from django.conf import settings

# Extensiones a probar al buscar un archivo historico, en orden de preferencia.
EXTENSIONES_FOTO = ['jpg', 'JPG', 'jpeg', 'JPEG', 'png', 'PNG', 'webp', 'bmp']
EXTENSIONES_FIRMA = ['png', 'PNG', 'jpg', 'JPG', 'jpeg', 'JPEG', 'webp', 'bmp']

# Mapa de magic bytes -> extension, para nombrar correctamente lo que llega en base64.
_FIRMAS_BINARIAS = [
    (b'\xff\xd8\xff', 'jpg'),
    (b'\x89PNG\r\n\x1a\n', 'png'),
    (b'GIF87a', 'gif'),
    (b'GIF89a', 'gif'),
    (b'BM', 'bmp'),
]

_RE_DATA_URI = re.compile(r'^data:(?P<mime>[\w/\-\.\+]+);base64,(?P<datos>.*)$', re.DOTALL)
# num_empleado se usa como nombre de archivo: solo permitimos caracteres seguros.
_RE_NOMBRE_SEGURO = re.compile(r'[^A-Za-z0-9_\-]')


class MediaError(Exception):
    """Error al leer o escribir un archivo de medios."""


def _media_root() -> Path:
    return Path(settings.MEDIA_ROOT)


def carpeta_fotos() -> str:
    return settings.MEDIA_DIR_FOTOS


def carpeta_firmas() -> str:
    return settings.MEDIA_DIR_FIRMAS


def carpeta_plantillas() -> str:
    return settings.MEDIA_DIR_PLANTILLAS


def nombre_seguro(valor) -> str:
    """Normaliza un identificador (num_empleado) para usarlo como nombre de archivo."""
    return _RE_NOMBRE_SEGURO.sub('', str(valor or '').strip())


def detectar_extension(contenido: bytes) -> str:
    for magic, ext in _FIRMAS_BINARIAS:
        if contenido.startswith(magic):
            return ext
    return 'png'


def es_data_uri(valor) -> bool:
    return isinstance(valor, str) and valor.startswith('data:')


def decodificar_base64(valor: str) -> bytes:
    """Acepta un data-URI completo o base64 pelon y regresa los bytes."""
    if not isinstance(valor, str):
        raise MediaError('El contenido de la imagen debe ser texto base64.')

    match = _RE_DATA_URI.match(valor.strip())
    datos = match.group('datos') if match else valor.strip()
    datos = re.sub(r'\s+', '', datos)

    try:
        return base64.b64decode(datos, validate=False)
    except (binascii.Error, ValueError) as exc:
        raise MediaError(f'No se pudo decodificar la imagen base64: {exc}') from exc


def resolver_existente(identificador, carpeta: str, extensiones) -> str | None:
    """
    Busca un archivo ya existente en MEDIA_ROOT/<carpeta> nombrado por identificador.

    Regresa la ruta relativa (ej. 'fotos/20222493.jpg') o None si no existe.
    """
    base = nombre_seguro(identificador)
    if not base:
        return None

    directorio = _media_root() / carpeta
    if not directorio.is_dir():
        return None

    for ext in extensiones:
        candidato = directorio / f'{base}.{ext}'
        if candidato.is_file():
            return f'{carpeta}/{base}.{ext}'

    return None


def resolver_foto(identificador) -> str | None:
    return resolver_existente(identificador, carpeta_fotos(), EXTENSIONES_FOTO)


def resolver_firma(identificador) -> str | None:
    return resolver_existente(identificador, carpeta_firmas(), EXTENSIONES_FIRMA)


# Los primeros 10 caracteres del RFC y del CURP de una misma persona son
# IDENTICOS: 4 letras derivadas del nombre + 6 digitos de la fecha de
# nacimiento (AAMMDD). Lo que cambia es la cola -- el RFC sigue con 3
# caracteres de homoclave, el CURP con sexo, entidad, consonantes internas,
# homoclave y digito verificador.
#
#   RFC : ZUAA771125 H4A                 (13, o 10 si se omite la homoclave)
#   CURP: ZUAA771125 HDFXXX01            (18)
#            ^^^^^^^^^^ mismo prefijo
#
# Ese prefijo comun es lo que permite cruzar las capturas de "Enrolamiento
# previo" (nombradas por RFC, porque a esa persona todavia no se le asigna
# numero de empleado) con el roster SIG, que solo trae CURP y nunca RFC.
LONGITUD_PREFIJO_IDENTIDAD = 10


def prefijo_identidad(valor) -> str:
    """Prefijo comun RFC/CURP (10 caracteres, normalizado a mayusculas)."""
    base = nombre_seguro(valor).upper()
    return base[:LONGITUD_PREFIJO_IDENTIDAD] if len(base) >= LONGITUD_PREFIJO_IDENTIDAD else ''


# Indice prefijo -> nombre de archivo, por carpeta. Recorrer el directorio en
# cada consulta cuesta ~14 ms con los ~14 300 archivos del acervo real, y
# resolver_por_prefijo() se invoca en la ruta interactiva de "Imprimir
# credenciales" (dos veces por empleado seleccionado: foto y firma). El indice
# se invalida solo comparando el mtime del directorio, que cambia en cuanto se
# agrega, renombra o borra un archivo -- justo lo que hacen guardar_imagen(),
# migrar_medios() y borrar().
_INDICE_PREFIJOS: dict = {}


def _indice_por_prefijo(carpeta: str, extensiones) -> dict:
    directorio = _media_root() / carpeta
    try:
        marca = directorio.stat().st_mtime_ns
    except OSError:
        return {}

    cacheado = _INDICE_PREFIJOS.get(carpeta)
    if cacheado and cacheado[0] == marca:
        return cacheado[1]

    extensiones_validas = {e.lower() for e in extensiones}
    indice: dict = {}

    try:
        entradas = list(os.scandir(directorio))
    except OSError:
        return {}

    for entrada in entradas:
        if not entrada.is_file():
            continue
        base, _, ext = entrada.name.rpartition('.')
        if not base or ext.lower() not in extensiones_validas:
            continue
        if len(base) < LONGITUD_PREFIJO_IDENTIDAD:
            continue

        clave = base[:LONGITUD_PREFIJO_IDENTIDAD].upper()
        previo = indice.get(clave)
        # Ante varios archivos con el mismo prefijo (en el acervo real solo
        # ocurre con duplicados de la MISMA persona: 'zalr641127' junto a
        # 'zalr641127-1'), se conserva el nombre mas corto y, a igual
        # longitud, el menor alfabeticamente -- es decir, el nombre base sin
        # sufijos, y siempre el mismo sin importar el orden del scandir.
        if previo is None or (len(entrada.name), entrada.name) < (len(previo), previo):
            indice[clave] = entrada.name

    _INDICE_PREFIJOS[carpeta] = (marca, indice)
    return indice


def resolver_por_prefijo(valor, carpeta: str, extensiones) -> str | None:
    """
    Busca un archivo cuyo nombre EMPIECE con el prefijo de identidad de
    `valor` (un CURP o un RFC).

    Necesario porque el nombre exacto del archivo no se puede reconstruir: el
    acervo tiene la foto guardada con RFC (con o sin homoclave) mientras que
    el roster SIG solo entrega el CURP, y la homoclave del RFC no es derivable
    del CURP. Solo el prefijo de 10 caracteres es comun a ambos.

    La comparacion es insensible a mayusculas porque el acervo historico
    mezcla ambos casos (zuaa771125.jpg y HEHA841115.jpg conviven), mientras
    que los CURP/RFC capturados llegan siempre en mayusculas.
    """
    prefijo = prefijo_identidad(valor)
    if not prefijo:
        return None

    nombre = _indice_por_prefijo(carpeta, extensiones).get(prefijo)
    return f'{carpeta}/{nombre}' if nombre else None


def resolver_con_respaldo(identificador_principal, identificador_respaldo, carpeta: str, extensiones):
    """
    Resuelve un medio probando, en orden:
      1. nombre exacto por num_empleado (el canonico),
      2. nombre exacto por CURP/RFC,
      3. prefijo de identidad de 10 caracteres (ver resolver_por_prefijo).

    El paso 3 es el que cruza las capturas de "Enrolamiento previo" -- hechas
    antes de que la persona tuviera numero asignado y por tanto nombradas con
    su RFC -- contra un roster que solo conoce el CURP.

    Regresa (ruta_relativa, origen) con origen en {'principal', 'respaldo',
    'prefijo'} o (None, None) si no se encontro nada.
    """
    ruta = resolver_existente(identificador_principal, carpeta, extensiones)
    if ruta:
        return ruta, 'principal'

    if identificador_respaldo:
        ruta = resolver_existente(identificador_respaldo, carpeta, extensiones)
        if ruta:
            return ruta, 'respaldo'

        ruta = resolver_por_prefijo(identificador_respaldo, carpeta, extensiones)
        if ruta:
            return ruta, 'prefijo'

    return None, None


def migrar_medios(identificador_destino, identificador_respaldo) -> dict:
    """
    Renombra foto y firma que hoy viven bajo un RFC/CURP para que pasen a
    llamarse por num_empleado, una vez que el movimiento de ingreso ya se
    aplico y el cruce quedo confirmado. Se dispara al imprimir la credencial.

    Si ya existe un archivo nombrado por num_empleado, ese se respeta y NO se
    sobreescribe: ya es el canonico. Regresa {'foto': bool, 'firma': bool}
    indicando que se migro realmente.
    """
    resultado = {'foto': False, 'firma': False}

    base_destino = nombre_seguro(identificador_destino)
    if not base_destino or not identificador_respaldo:
        return resultado

    for carpeta, extensiones, clave in (
        (carpeta_fotos(), EXTENSIONES_FOTO, 'foto'),
        (carpeta_firmas(), EXTENSIONES_FIRMA, 'firma'),
    ):
        # Ya existe el canonico: no se toca nada.
        if resolver_existente(identificador_destino, carpeta, extensiones):
            continue

        origen = (
            resolver_existente(identificador_respaldo, carpeta, extensiones)
            or resolver_por_prefijo(identificador_respaldo, carpeta, extensiones)
        )
        if not origen:
            continue

        ext = origen.rsplit('.', 1)[-1]
        try:
            (_media_root() / origen).rename(
                _media_root() / carpeta / f'{base_destino}.{ext}'
            )
            resultado[clave] = True
        except OSError:
            pass

    return resultado


def _borrar_variantes(directorio: Path, base: str, extensiones) -> None:
    """Elimina versiones previas del mismo identificador con otra extension."""
    for ext in extensiones:
        previo = directorio / f'{base}.{ext}'
        if previo.is_file():
            try:
                previo.unlink()
            except OSError:
                pass


def convertir_a_png(contenido: bytes) -> bytes:
    """
    Reescribe cualquier imagen como PNG, SIN redimensionar.

    PNG es sin perdida: reguardar una foto no la degrada cada vez que se
    corrige, cosa que con JPEG si pasa. Las firmas ademas dependen de ello --
    en el acervo real vienen en modo RGBA, y aplanar el canal alfa las dejaria
    con fondo negro sobre la credencial.

    Las dimensiones se conservan tal cual: la resolucion la define la camara,
    no el sistema.
    """
    try:
        from PIL import Image
    except ImportError:
        # Sin Pillow se guarda el original: peor formato, pero nunca perder
        # la captura que el operador acaba de tomar.
        return contenido

    from io import BytesIO
    try:
        with Image.open(BytesIO(contenido)) as imagen:
            # 'P' (paleta) puede llevar transparencia; se pasa a RGBA para no
            # perderla. El resto conserva su modo: RGB sigue RGB.
            if imagen.mode in ('P', 'LA'):
                imagen = imagen.convert('RGBA')
            elif imagen.mode not in ('RGB', 'RGBA', 'L'):
                imagen = imagen.convert('RGB')

            salida = BytesIO()
            imagen.save(salida, format='PNG', optimize=True)
            return salida.getvalue()
    except Exception:
        return contenido


def guardar_imagen(
    contenido_base64: str, identificador, carpeta: str, extensiones,
    forzar_png: bool = False,
) -> str:
    """
    Guarda una imagen base64 en MEDIA_ROOT/<carpeta>/<identificador>.<ext>.

    Sobrescribe cualquier version previa del mismo identificador (incluso si tenia
    otra extension) para que la ruta siga siendo deterministica. Regresa la ruta
    relativa guardada.
    """
    base = nombre_seguro(identificador)
    if not base:
        raise MediaError('Se requiere un identificador (num_empleado) para guardar la imagen.')

    contenido = decodificar_base64(contenido_base64)
    if not contenido:
        raise MediaError('La imagen recibida esta vacia.')

    # Solo se fuerza PNG donde hace falta (firmas). Ver guardar_foto/guardar_firma.
    if forzar_png:
        contenido = convertir_a_png(contenido)

    directorio = _media_root() / carpeta
    directorio.mkdir(parents=True, exist_ok=True)

    ext = detectar_extension(contenido)
    _borrar_variantes(directorio, base, extensiones)

    destino = directorio / f'{base}.{ext}'
    with open(destino, 'wb') as archivo:
        archivo.write(contenido)

    return f'{carpeta}/{base}.{ext}'


def guardar_foto(contenido_base64: str, identificador) -> str:
    """
    Guarda la foto RESPETANDO su formato de origen.

    No se convierte a PNG a proposito: el acervo real son 14 312 fotos JPEG de
    ~88 KB, y reguardarlas en PNG las multiplica por ~9 (34 KB -> 308 KB
    medido). La camara ya entrega el formato y la resolucion adecuados; el
    sistema solo tiene que no estropearlos.
    """
    return guardar_imagen(contenido_base64, identificador, carpeta_fotos(), EXTENSIONES_FOTO)


def guardar_firma(contenido_base64: str, identificador) -> str:
    """
    Guarda la firma SIEMPRE como PNG.

    Aqui si hace falta: las firmas del acervo vienen en modo RGBA y su fondo
    transparente es lo que permite montarlas sobre la credencial. Un JPEG no
    tiene canal alfa, asi que convertirlas a ese formato las dejaria con
    fondo negro sobre el gafete.
    """
    return guardar_imagen(
        contenido_base64, identificador, carpeta_firmas(), EXTENSIONES_FIRMA,
        forzar_png=True,
    )


# RFC de persona fisica: 4 letras + 6 digitos (AAMMDD) + 3 de homoclave.
# La homoclave con frecuencia no se captura, de ahi que sea opcional.
_RE_RFC = re.compile(r'^[A-ZÑ&]{4}[0-9]{6}([A-Z0-9]{3})?$', re.IGNORECASE)
# CURP: 4 letras + 6 digitos + H/M + 5 letras + homoclave + digito = 18.
_RE_CURP = re.compile(r'^[A-Z][AEIOUX][A-Z]{2}[0-9]{6}[HM][A-Z]{5}[A-Z0-9][0-9]$', re.IGNORECASE)


def es_rfc(valor) -> bool:
    return bool(_RE_RFC.match(str(valor or '').strip()))


def es_curp(valor) -> bool:
    return bool(_RE_CURP.match(str(valor or '').strip()))


def _es_identificador_previo(base: str) -> bool:
    """
    True si el nombre de archivo corresponde a una captura pendiente de
    cruzarse con un num_empleado, es decir, nombrada por RFC o CURP.

    Los num_empleado del roster SIG son SIEMPRE numericos, asi que basta con
    exigir que el nombre traiga letras. Esto incluye a proposito los ~950
    archivos historicos nombrados con RFC corto en minusculas
    (zuaa771125.jpg): NO son ruido, son justamente capturas previas que
    todavia no se han cruzado, y el prefijo comun RFC/CURP de 10 caracteres
    permite cruzarlas igual que las nuevas (ver resolver_por_prefijo).
    """
    return bool(base) and not base.isdigit()


def listar_medios_por_identificador(solo_numericos: bool) -> dict:
    """
    Recorre fotos/ y FIRMAS/ y agrupa los archivos por nombre.

    `solo_numericos=True` devuelve los nombrados por num_empleado; False, los
    nombrados por RFC/CURP (pendientes de cruce). Es la misma pasada de disco
    para ambos casos, solo cambia el filtro.

    Regresa {identificador: {'identificador', 'foto', 'firma', 'fecha'}} con
    URLs publicas ya resueltas.
    """
    registros = {}

    for carpeta, extensiones, clave in (
        (carpeta_fotos(), EXTENSIONES_FOTO, 'foto'),
        (carpeta_firmas(), EXTENSIONES_FIRMA, 'firma'),
    ):
        directorio = _media_root() / carpeta
        if not directorio.is_dir():
            continue

        validas = {e.lower() for e in extensiones}
        try:
            entradas = list(os.scandir(directorio))
        except OSError:
            continue

        for entrada in entradas:
            if not entrada.is_file():
                continue
            base, _, ext = entrada.name.rpartition('.')
            if not base or ext.lower() not in validas:
                continue
            # isdigit() distingue num_empleado (siempre numerico) de RFC/CURP
            # (siempre con letras). Es el mismo criterio que usa el cruce.
            if base.isdigit() != solo_numericos:
                continue

            item = registros.setdefault(
                base, {'identificador': base, 'foto': None, 'firma': None, 'fecha': None}
            )
            item[clave] = url_publica(f'{carpeta}/{entrada.name}', versionada=True)
            try:
                modificado = entrada.stat().st_mtime
            except OSError:
                modificado = None
            if modificado and (item['fecha'] is None or modificado > item['fecha']):
                item['fecha'] = modificado

    return registros


def listar_enrolamientos_previos() -> list:
    """
    Capturas nombradas por RFC/CURP que aun no se cruzan con un num_empleado.
    Ordenadas por fecha de captura descendente.
    """
    registros = listar_medios_por_identificador(solo_numericos=False)
    for item in registros.values():
        item['rfc'] = item.pop('identificador')
    return sorted(registros.values(), key=lambda r: r['fecha'] or 0, reverse=True)


def renombrar_medios(identificador_actual, identificador_nuevo) -> dict:
    """
    Renombra foto y firma de un identificador a otro.

    Sirve para corregir una captura guardada con el RFC mal escrito: mientras
    el nombre este mal, "Imprimir credenciales" no la encuentra al cruzar.

    NO sobreescribe: si el destino ya tiene archivo, se aborta y se informa.
    Pisar en silencio la foto de otra persona seria mucho peor que fallar.
    """
    origen_base = nombre_seguro(identificador_actual)
    destino_base = nombre_seguro(identificador_nuevo)

    if not origen_base or not destino_base:
        raise MediaError('Se requieren el nombre actual y el nuevo.')
    if origen_base == destino_base:
        raise MediaError('El nombre nuevo es igual al actual.')

    plan = []
    for carpeta, extensiones, clave in (
        (carpeta_fotos(), EXTENSIONES_FOTO, 'foto'),
        (carpeta_firmas(), EXTENSIONES_FIRMA, 'firma'),
    ):
        if resolver_existente(destino_base, carpeta, extensiones):
            raise MediaError(
                f'Ya existe un archivo de {clave} con el nombre "{destino_base}". '
                'Revisa ese registro antes de renombrar.'
            )
        origen = resolver_existente(origen_base, carpeta, extensiones)
        if origen:
            plan.append((carpeta, origen, clave))

    if not plan:
        raise MediaError(f'No hay archivos guardados con el nombre "{origen_base}".')

    resultado = {'foto': False, 'firma': False}
    for carpeta, origen, clave in plan:
        ext = origen.rsplit('.', 1)[-1]
        try:
            (_media_root() / origen).rename(
                _media_root() / carpeta / f'{destino_base}.{ext}'
            )
            resultado[clave] = True
        except OSError as exc:
            raise MediaError(f'No se pudo renombrar el archivo de {clave}: {exc}') from exc

    return resultado


def borrar_medios(identificador) -> dict:
    """Elimina foto y firma de un identificador. Regresa que se borro."""
    resultado = {'foto': False, 'firma': False}
    for carpeta, extensiones, clave in (
        (carpeta_fotos(), EXTENSIONES_FOTO, 'foto'),
        (carpeta_firmas(), EXTENSIONES_FIRMA, 'firma'),
    ):
        ruta = resolver_existente(identificador, carpeta, extensiones)
        if ruta:
            resultado[clave] = borrar(ruta)
    return resultado


def guardar_fondo_plantilla(contenido_base64: str, nombre_archivo: str) -> str:
    """Guarda un fondo de plantilla; el nombre lo define quien llama (no es por empleado)."""
    base = nombre_seguro(nombre_archivo)
    if not base:
        raise MediaError('Se requiere un nombre para el fondo de la plantilla.')

    contenido = decodificar_base64(contenido_base64)
    if not contenido:
        raise MediaError('El fondo recibido esta vacio.')

    directorio = _media_root() / carpeta_plantillas()
    directorio.mkdir(parents=True, exist_ok=True)

    ext = detectar_extension(contenido)
    destino = directorio / f'{base}.{ext}'
    with open(destino, 'wb') as archivo:
        archivo.write(contenido)

    return f'{carpeta_plantillas()}/{base}.{ext}'


def existe(ruta_relativa) -> bool:
    if not ruta_relativa:
        return False
    return (_media_root() / str(ruta_relativa)).is_file()


def url_publica(ruta_relativa, versionada=True) -> str | None:
    """
    Convierte 'fotos/123.jpg' en '/media/fotos/123.jpg?v=<mtime>'.

    El versionado va ACTIVADO por omision, no como opcion: los archivos se
    guardan siempre en la MISMA ruta determinista, asi que al reemplazar una
    foto la URL no cambia, el navegador la sirve de su cache y parece que el
    cambio no se guardo. Ya paso: con el respaldo desactivado por defecto se
    versionaba el inventario pero no el endpoint `medios/`, y la credencial
    seguia mostrando la firma vieja varios minutos.

    Solo la extension delataba el cambio -- una foto .jpg reemplazada por .png
    si cambiaba de URL, una firma .png -> .png nunca --, lo que hacia que el
    fallo pareciera aleatorio.

    Con la marca de tiempo la URL cambia justo cuando cambia el contenido, y
    sigue cacheando mientras no cambie.
    """
    if not ruta_relativa:
        return None

    url = f"{settings.MEDIA_URL}{str(ruta_relativa).lstrip('/')}"
    if not versionada:
        return url

    try:
        marca = int((_media_root() / str(ruta_relativa)).stat().st_mtime)
    except OSError:
        return url
    return f'{url}?v={marca}'


def leer_como_data_uri(ruta_relativa) -> str | None:
    """Lee un archivo de MEDIA_ROOT y lo regresa como data-URI base64."""
    if not ruta_relativa:
        return None

    ruta = _media_root() / str(ruta_relativa)
    if not ruta.is_file():
        return None

    try:
        contenido = ruta.read_bytes()
    except OSError:
        return None

    ext = detectar_extension(contenido)
    mime = 'image/jpeg' if ext in ('jpg', 'jpeg') else f'image/{ext}'
    return f'data:{mime};base64,{base64.b64encode(contenido).decode("ascii")}'


# ========================================================================
# Archivo historico de medios (para la auditoria de credenciales)
# ========================================================================

CARPETA_HISTORICO = 'historico'


def carpeta_historico() -> str:
    return CARPETA_HISTORICO


def ruta_relativa_desde_url(valor) -> str | None:
    """
    Extrae la ruta relativa a MEDIA_ROOT de una URL de medios.

    Acepta las tres formas que produce el frontend:
        'http://host:4200/media/fotos/123.png?v=1786056998'
        '/media/FIRMAS/123.png'
        'fotos/123.png'

    Regresa None para data-URIs y para cualquier URL que no cuelgue de
    MEDIA_URL -- de otro modo se intentaria archivar algo que no es un medio
    nuestro.
    """
    if not valor:
        return None

    texto = str(valor)
    if texto.startswith('data:'):
        return None

    # Quitar esquema+host y la cadena de version (?v=<mtime>).
    partes = urlsplit(texto)
    ruta = unquote(partes.path or texto)

    prefijo = settings.MEDIA_URL          # normalmente '/media/'
    if prefijo and prefijo in ruta:
        ruta = ruta.split(prefijo, 1)[1]
    elif ruta.startswith('/'):
        # Ruta absoluta que no cuelga de MEDIA_URL: no es un medio nuestro.
        return None

    ruta = ruta.lstrip('/')
    if not ruta or '..' in ruta.split('/'):
        return None
    return ruta


def archivar_medio(ruta_relativa) -> str | None:
    """
    Copia un archivo de MEDIA_ROOT al archivo historico y regresa su nueva
    ruta relativa, o None si no existe.

    El destino se nombra por el SHA-256 de su CONTENIDO
    (`historico/<aa>/<hash>.<ext>`), lo que da tres propiedades que aqui
    importan:

    - **Inmutabilidad**: la ruta identifica unos bytes concretos. Si manana
      reemplazan la foto del empleado, la credencial ya impresa sigue
      apuntando a la imagen con la que REALMENTE se expidio. Sin esto la
      auditoria mostraria la foto actual sobre un folio viejo, que es
      justamente la mentira que una auditoria no puede contar.
    - **Deduplicacion**: reimprimir a la misma persona con la misma foto no
      cuesta un byte extra, porque el destino ya existe.
    - **Idempotencia**: archivar dos veces es una sola copia.

    El subdirectorio de dos caracteres evita meter decenas de miles de
    archivos en un solo directorio, que degrada el listado en la mayoria de
    sistemas de archivos.
    """
    if not ruta_relativa:
        return None

    origen = _media_root() / str(ruta_relativa)
    if not origen.is_file():
        return None

    # Ya archivado: no volver a copiar (evita historico/<x>/historico/<y>).
    if str(ruta_relativa).startswith(f'{CARPETA_HISTORICO}/'):
        return str(ruta_relativa)

    try:
        contenido = origen.read_bytes()
    except OSError:
        return None

    digest = hashlib.sha256(contenido).hexdigest()
    ext = detectar_extension(contenido) or (origen.suffix.lstrip('.') or 'bin')

    relativa = f'{CARPETA_HISTORICO}/{digest[:2]}/{digest}.{ext}'
    destino = _media_root() / relativa

    if destino.is_file():
        return relativa

    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        # Escribir a un temporal y renombrar: si dos impresiones concurrentes
        # archivan el mismo archivo, ninguna lee uno a medio escribir.
        temporal = destino.with_suffix(destino.suffix + '.tmp')
        temporal.write_bytes(contenido)
        os.replace(temporal, destino)
    except OSError:
        return None

    return relativa


def borrar(ruta_relativa) -> bool:
    if not ruta_relativa:
        return False
    ruta = _media_root() / str(ruta_relativa)
    if not ruta.is_file():
        return False
    try:
        os.remove(ruta)
        return True
    except OSError:
        return False
