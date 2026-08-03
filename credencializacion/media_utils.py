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
