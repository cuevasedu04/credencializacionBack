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
import os
import re
from pathlib import Path

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


def guardar_imagen(contenido_base64: str, identificador, carpeta: str, extensiones) -> str:
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

    directorio = _media_root() / carpeta
    directorio.mkdir(parents=True, exist_ok=True)

    ext = detectar_extension(contenido)
    _borrar_variantes(directorio, base, extensiones)

    destino = directorio / f'{base}.{ext}'
    with open(destino, 'wb') as archivo:
        archivo.write(contenido)

    return f'{carpeta}/{base}.{ext}'


def guardar_foto(contenido_base64: str, identificador) -> str:
    return guardar_imagen(contenido_base64, identificador, carpeta_fotos(), EXTENSIONES_FOTO)


def guardar_firma(contenido_base64: str, identificador) -> str:
    return guardar_imagen(contenido_base64, identificador, carpeta_firmas(), EXTENSIONES_FIRMA)


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


def listar_enrolamientos_previos() -> list:
    """
    Lista las capturas guardadas por CURP que aun no se han cruzado con un
    num_empleado (ver migrar_medios). Sirve para que "Enrolamiento previo"
    sepa a quien ya se le tomo foto/firma, incluso despues de recargar la
    pagina -- no hay tabla en BD que lo registre, la unica fuente de verdad
    son los archivos en MEDIA_ROOT.

    Regresa una lista de dicts ordenada por fecha de captura descendente.
    """
    registros = {}

    for carpeta, extensiones, clave in (
        (carpeta_fotos(), EXTENSIONES_FOTO, 'foto'),
        (carpeta_firmas(), EXTENSIONES_FIRMA, 'firma'),
    ):
        directorio = _media_root() / carpeta
        if not directorio.is_dir():
            continue

        for archivo in directorio.iterdir():
            if not archivo.is_file():
                continue

            base = archivo.stem
            ext = archivo.suffix.lstrip('.')
            if ext not in extensiones or not _es_identificador_previo(base):
                continue

            entrada = registros.setdefault(
                base, {'rfc': base, 'foto': None, 'firma': None, 'fecha': None}
            )
            entrada[clave] = url_publica(f'{carpeta}/{archivo.name}')

            try:
                modificado = archivo.stat().st_mtime
            except OSError:
                modificado = None
            if modificado and (entrada['fecha'] is None or modificado > entrada['fecha']):
                entrada['fecha'] = modificado

    return sorted(registros.values(), key=lambda r: r['fecha'] or 0, reverse=True)


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


def url_publica(ruta_relativa) -> str | None:
    """Convierte 'fotos/123.jpg' en '/media/fotos/123.jpg'."""
    if not ruta_relativa:
        return None
    return f"{settings.MEDIA_URL}{str(ruta_relativa).lstrip('/')}"


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
