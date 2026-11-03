"""
spotify_playlists.py
====================
Organiza tus canciones de "Me gusta" de Spotify en playlists.

Comandos (en la terminal, dentro de la carpeta del script):

    python spotify_playlists.py token
        -> SOLO EN TU COMPUTADOR. Inicia sesión en Spotify y te muestra el
           "refresh token" que debes guardar como secret en GitHub.
           Repetirlo cada ~6 meses (cuando el token vence).

    python spotify_playlists.py actualizar
        -> Lo corre GitHub cada semana:
           1) Carga las canciones de sin_clasificar.xlsx que ya tengan la
              columna "Playlist" llena (y aprende esos artistas).
           2) Revisa las canciones NUEVAS de "Me gusta" y las mete en su
              playlist según el artista (reglas_playlists.json).
           3) Las de artistas desconocidos quedan en sin_clasificar.xlsx y
              te llegan por correo.

    python spotify_playlists.py exportar
        -> Crea mis_me_gusta.xlsx con TODAS tus canciones de "Me gusta".

    python spotify_playlists.py cargar archivo.xlsx
        -> Mete en las playlists las canciones del Excel según su columna
           "Playlist" (varias separadas por ";").

Requisitos:  pip install requests openpyxl
"""

import base64
import http.server
import json
import os
import secrets
import smtplib
import sys
import time
import urllib.parse
import webbrowser
from email.message import EmailMessage
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------------
CLIENT_ID = os.environ.get("SPOTIFY_CLIENT_ID", "2e76d5745cee40399c8a0f681f70c230")
CLIENT_SECRET = os.environ.get("SPOTIFY_CLIENT_SECRET", "")
REFRESH_TOKEN = os.environ.get("SPOTIFY_REFRESH_TOKEN", "")

# Correo (solo se usa en GitHub para avisarte de canciones sin clasificar)
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
CORREO_DESTINO = os.environ.get("CORREO_DESTINO", "")

# Debe ser EXACTAMENTE igual al Redirect URI que pusiste en Spotify
REDIRECT_URI = "http://127.0.0.1:8888/callback"
SCOPES = "user-library-read playlist-read-private playlist-modify-private playlist-modify-public"

CARPETA = Path(__file__).resolve().parent
REGLAS_FILE = CARPETA / "reglas_playlists.json"
PROCESADAS_FILE = CARPETA / "procesadas.txt"
SIN_CLASIFICAR = CARPETA / "sin_clasificar.xlsx"
REGISTRO = CARPETA / "registro_actualizaciones.txt"
ARCHIVO_EXPORT = CARPETA / "mis_me_gusta.xlsx"

API = "https://api.spotify.com/v1"
TOKEN_URL = "https://accounts.spotify.com/api/token"
AUTH_URL = "https://accounts.spotify.com/authorize"

ENCABEZADOS = ["uri", "Canción", "Artista(s)", "Álbum", "Año", "Agregada el", "Playlist"]

_token_actual = {"access_token": None, "expires_at": 0}


class ErrorSpotify(Exception):
    def __init__(self, status, mensaje):
        super().__init__(mensaje)
        self.status = status


def log(texto=""):
    """Imprime en pantalla y lo deja anotado en registro_actualizaciones.txt"""
    print(texto)
    with open(REGISTRO, "a", encoding="utf-8") as f:
        f.write(texto + "\n")


# ------------------------------------------------------------------
# 1. INICIO DE SESIÓN
# ------------------------------------------------------------------
def _cabecera_basica():
    par = f"{CLIENT_ID}:{CLIENT_SECRET}".encode()
    return {"Authorization": "Basic " + base64.b64encode(par).decode()}


class _Receptor(http.server.BaseHTTPRequestHandler):
    codigo = None
    error = None
    state = None

    def do_GET(self):
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        if "code" in q or "error" in q:
            _Receptor.codigo = q.get("code", [None])[0]
            _Receptor.error = q.get("error", [None])[0]
            _Receptor.state = q.get("state", [None])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        msg = ("Listo. Ya puedes cerrar esta pestaña y volver a la terminal."
               if _Receptor.codigo else "No se completó la autorización. Vuelve a la terminal.")
        self.wfile.write(f"<h2 style='font-family:sans-serif'>{msg}</h2>".encode("utf-8"))

    def log_message(self, *args):
        pass


def comando_token():
    """Inicia sesión en el navegador y muestra el refresh token para GitHub."""
    global CLIENT_SECRET
    if not CLIENT_SECRET:
        CLIENT_SECRET = input("Pega aquí tu Client Secret de Spotify y dale Enter: ").strip()

    state = secrets.token_urlsafe(16)
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
        "state": state,
        "show_dialog": "true",
    }
    url = AUTH_URL + "?" + urllib.parse.urlencode(params)
    servidor = http.server.HTTPServer(("127.0.0.1", 8888), _Receptor)
    servidor.timeout = 5

    print("\nSe va a abrir el navegador para que inicies sesión en Spotify.")
    print("Si no se abre solo, copia y pega este enlace:\n\n" + url + "\n")
    webbrowser.open(url)

    limite = time.time() + 300
    while _Receptor.codigo is None and _Receptor.error is None:
        if time.time() > limite:
            sys.exit("Se acabó el tiempo esperando el inicio de sesión.")
        servidor.handle_request()
    servidor.server_close()

    if _Receptor.error or _Receptor.state != state:
        sys.exit(f"No se pudo iniciar sesión ({_Receptor.error or 'state inválido'}).")

    r = requests.post(
        TOKEN_URL,
        headers=_cabecera_basica(),
        data={"grant_type": "authorization_code", "code": _Receptor.codigo, "redirect_uri": REDIRECT_URI},
        timeout=30,
    )
    if r.status_code != 200:
        sys.exit(f"Spotify rechazó el inicio de sesión. Revisa el Client Secret.\nDetalle: {r.text}")

    print("\n" + "=" * 60)
    print("TU REFRESH TOKEN (cópialo completo, es una sola línea):\n")
    print(r.json()["refresh_token"])
    print("=" * 60)
    print("\nPégalo en GitHub como el secret SPOTIFY_REFRESH_TOKEN.")


def obtener_token():
    if _token_actual["access_token"] and time.time() < _token_actual["expires_at"] - 60:
        return _token_actual["access_token"]
    if not (CLIENT_SECRET and REFRESH_TOKEN):
        sys.exit("Faltan SPOTIFY_CLIENT_SECRET o SPOTIFY_REFRESH_TOKEN (secrets de GitHub).")
    r = requests.post(
        TOKEN_URL,
        headers=_cabecera_basica(),
        data={"grant_type": "refresh_token", "refresh_token": REFRESH_TOKEN},
        timeout=30,
    )
    if r.status_code != 200:
        log(f"ERROR: Spotify no aceptó el refresh token ({r.status_code}). "
            "Probablemente venció: corre 'python spotify_playlists.py token' en tu "
            "computador y actualiza el secret SPOTIFY_REFRESH_TOKEN en GitHub.")
        sys.exit(1)
    datos = r.json()
    _token_actual["access_token"] = datos["access_token"]
    _token_actual["expires_at"] = time.time() + int(datos.get("expires_in", 3600))
    return datos["access_token"]


# ------------------------------------------------------------------
# 2. LLAMADAS A LA API (con reintentos)
# ------------------------------------------------------------------
def api(metodo, ruta, **kwargs):
    url = ruta if ruta.startswith("http") else API + ruta
    for _ in range(6):
        r = requests.request(metodo, url, headers={"Authorization": f"Bearer {obtener_token()}"},
                             timeout=30, **kwargs)
        if r.status_code == 429:
            time.sleep(int(r.headers.get("Retry-After", "5")) + 1)
            continue
        if r.status_code == 401:
            _token_actual["access_token"] = None
            continue
        if r.status_code >= 500:
            time.sleep(3)
            continue
        if r.status_code >= 400:
            raise ErrorSpotify(r.status_code, r.text)
        return r.json() if r.content else {}
    raise ErrorSpotify(0, "Spotify no respondió después de varios intentos.")


def paginar(ruta):
    url = ruta
    while url:
        datos = api("GET", url)
        yield from datos.get("items", [])
        url = datos.get("next")


# ------------------------------------------------------------------
# 3. LEER "ME GUSTA"
# ------------------------------------------------------------------
def leer_me_gusta():
    canciones = []
    for it in paginar("/me/tracks?limit=50"):
        t = it.get("track") or it.get("item")
        if not t or not t.get("uri") or t.get("is_local"):
            continue
        canciones.append([
            t["uri"],
            t.get("name", ""),
            ", ".join(a.get("name", "") for a in t.get("artists", [])),
            (t.get("album") or {}).get("name", ""),
            ((t.get("album") or {}).get("release_date") or "")[:4],
            (it.get("added_at") or "")[:10],
            "",
        ])
    return canciones


def _dar_formato(ws):
    relleno = PatternFill("solid", fgColor="1DB954")
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
    for i, ancho in enumerate([40, 40, 35, 35, 8, 13, 35], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def exportar():
    canciones = leer_me_gusta()
    wb = Workbook()
    ws = wb.active
    ws.title = "Me gusta"
    ws.append(ENCABEZADOS)
    for fila in canciones:
        ws.append(fila)
    _dar_formato(ws)
    wb.save(ARCHIVO_EXPORT)
    print(f"Listo: {len(canciones)} canciones guardadas en {ARCHIVO_EXPORT.name}")


# ------------------------------------------------------------------
# 4. REGLAS Y CANCIONES PROCESADAS
# ------------------------------------------------------------------
def leer_reglas():
    return json.loads(REGLAS_FILE.read_text(encoding="utf-8")) if REGLAS_FILE.exists() else {}


def guardar_reglas(reglas):
    REGLAS_FILE.write_text(json.dumps(reglas, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")


def leer_procesadas():
    return set(PROCESADAS_FILE.read_text(encoding="utf-8").split()) if PROCESADAS_FILE.exists() else set()


def marcar_procesadas(uris):
    with open(PROCESADAS_FILE, "a", encoding="utf-8") as f:
        for u in uris:
            f.write(u + "\n")


# ------------------------------------------------------------------
# 5. METER CANCIONES EN LAS PLAYLISTS
# ------------------------------------------------------------------
def mis_playlists(mi_id):
    resultado = {}
    for p in paginar("/me/playlists?limit=50"):
        if p and (p.get("owner") or {}).get("id") == mi_id:
            resultado.setdefault(p["name"].strip(), p["id"])
    return resultado


def crear_playlist(mi_id, nombre):
    cuerpo = {"name": nombre, "public": False, "description": "Creada a partir de mis Me gusta"}
    try:
        return api("POST", "/me/playlists", json=cuerpo)["id"]
    except ErrorSpotify as e:
        if e.status in (404, 405):
            return api("POST", f"/users/{mi_id}/playlists", json=cuerpo)["id"]
        raise


def canciones_en_playlist(pid):
    uris = set()
    for it in paginar(f"/playlists/{pid}/items?limit=100"):
        t = it.get("item") or it.get("track")
        if t and t.get("uri"):
            uris.add(t["uri"])
    return uris


def agregar(pid, uris):
    for i in range(0, len(uris), 100):
        lote = {"uris": uris[i:i + 100]}
        try:
            api("POST", f"/playlists/{pid}/items", json=lote)
        except ErrorSpotify as e:
            if e.status in (404, 405):
                api("POST", f"/playlists/{pid}/tracks", json=lote)
            else:
                raise


def cargar_grupos(grupos):
    """grupos = {nombre de playlist: [uris]}"""
    if not grupos:
        return
    mi_id = api("GET", "/me")["id"]
    existentes = mis_playlists(mi_id)
    for nombre, uris in grupos.items():
        pid = existentes.get(nombre)
        if pid:
            ya, accion = canciones_en_playlist(pid), "actualizada"
        else:
            pid, ya, accion = crear_playlist(mi_id, nombre), set(), "creada"
            existentes[nombre] = pid
        nuevas = [u for u in uris if u not in ya]
        agregar(pid, nuevas)
        log(f"  {nombre}: {accion}, {len(nuevas)} canciones agregadas")


def _leer_excel(archivo):
    """Devuelve (filas como listas de 7 columnas). Ignora filas vacías."""
    wb = load_workbook(archivo)
    ws = wb.active
    filas = []
    for fila in ws.iter_rows(min_row=2, values_only=True):
        fila = list(fila) + [None] * (7 - len(fila))
        if fila[0]:
            filas.append(fila[:7])
    return filas


def _cargar_filas(filas):
    """Carga filas con Playlist llena; aprende artistas. Devuelve las uris cargadas."""
    grupos, cargadas, votos = {}, [], {}
    for f in filas:
        if not f[6]:
            continue
        uri = str(f[0]).strip()
        nombres = [n.strip() for n in str(f[6]).split(";") if n.strip()]
        for n in nombres:
            if uri not in grupos.setdefault(n, []):
                grupos[n].append(uri)
        cargadas.append(uri)
        if f[2] and nombres:
            artista = str(f[2]).split(", ")[0].strip()
            votos.setdefault(artista, {}).setdefault(nombres[0], 0)
            votos[artista][nombres[0]] += 1

    cargar_grupos(grupos)

    reglas = leer_reglas()
    nuevos = 0
    for artista, cuenta in votos.items():
        if artista not in reglas:
            reglas[artista] = max(cuenta, key=cuenta.get)
            nuevos += 1
    guardar_reglas(reglas)
    marcar_procesadas(cargadas)
    if nuevos:
        log(f"  Aprendí {nuevos} artistas nuevos.")
    return cargadas


def cargar(archivo):
    if not archivo.exists():
        sys.exit(f"No encuentro el archivo {archivo.name}.")
    cargadas = _cargar_filas(_leer_excel(archivo))
    print(f"Listo: {len(cargadas)} canciones cargadas.")


# ------------------------------------------------------------------
# 6. CORREO
# ------------------------------------------------------------------
def enviar_correo(asunto, cuerpo, adjunto=None):
    if not (GMAIL_USER and GMAIL_APP_PASSWORD and CORREO_DESTINO):
        log("  (No se envió correo: faltan los secrets de Gmail)")
        return
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = GMAIL_USER
    msg["To"] = CORREO_DESTINO
    msg.set_content(cuerpo)
    if adjunto and adjunto.exists():
        msg.add_attachment(
            adjunto.read_bytes(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=adjunto.name,
        )
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        s.send_message(msg)
    log(f"  Correo enviado a {CORREO_DESTINO}")


# ------------------------------------------------------------------
# 7. ACTUALIZACIÓN SEMANAL
# ------------------------------------------------------------------
def actualizar():
    log(f"\n=== Actualización {time.strftime('%Y-%m-%d %H:%M')} (UTC) ===")
    reglas = leer_reglas()
    if not reglas:
        sys.exit("No encuentro reglas_playlists.json.")

    # 1) Canciones de sin_clasificar.xlsx que ya tienen Playlist
    pendientes = []
    if SIN_CLASIFICAR.exists():
        filas = _leer_excel(SIN_CLASIFICAR)
        llenas = [f for f in filas if f[6]]
        pendientes = [f for f in filas if not f[6]]
        if llenas:
            log(f"Cargando {len(llenas)} canciones que clasificaste en sin_clasificar.xlsx:")
            _cargar_filas(llenas)
            reglas = leer_reglas()

    # 2) Canciones nuevas de "Me gusta"
    procesadas = leer_procesadas()
    ya_pendientes = {str(f[0]) for f in pendientes}
    nuevas = [c for c in leer_me_gusta() if c[0] not in procesadas and c[0] not in ya_pendientes]

    grupos, sin = {}, []
    for c in nuevas:
        artistas = [a.strip() for a in c[2].split(", ")]
        destino = next((reglas[a] for a in artistas if a in reglas), None)
        if destino:
            grupos.setdefault(destino, []).append(c[0])
        else:
            sin.append(c)

    log(f"{len(nuevas)} canciones nuevas: {len(nuevas) - len(sin)} clasificadas, {len(sin)} de artistas nuevos.")
    cargar_grupos(grupos)
    marcar_procesadas(u for d in grupos.values() for u in d)

    # 3) Reescribir sin_clasificar.xlsx solo con lo que sigue pendiente
    pendientes += sin
    if pendientes:
        wb = Workbook()
        ws = wb.active
        ws.title = "Sin clasificar"
        ws.append(ENCABEZADOS)
        for f in pendientes:
            ws.append(list(f))
        _dar_formato(ws)
        wb.save(SIN_CLASIFICAR)
    elif SIN_CLASIFICAR.exists():
        SIN_CLASIFICAR.unlink()

    # 4) Avisar por correo si hay canciones nuevas sin clasificar
    if sin:
        lista = "\n".join(f"  - {c[1]} — {c[2]}" for c in sin)
        enviar_correo(
            f"Spotify: {len(sin)} canciones nuevas sin clasificar",
            "Estas canciones nuevas de tus 'Me gusta' son de artistas que todavía no tienen playlist:\n\n"
            f"{lista}\n\n"
            "Para ubicarlas: abre sin_clasificar.xlsx en tu repositorio de GitHub (también va adjunto), "
            "llena la columna 'Playlist' con el nombre exacto de la playlist y súbelo de nuevo al repositorio. "
            "La próxima actualización las carga y aprende esos artistas.",
            SIN_CLASIFICAR,
        )
    log("Listo.")


# ------------------------------------------------------------------
if __name__ == "__main__":
    accion = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    try:
        if accion == "token":
            comando_token()
        elif accion == "actualizar":
            actualizar()
        elif accion == "exportar":
            exportar()
        elif accion == "cargar" and len(sys.argv) > 2:
            archivo = Path(sys.argv[2])
            cargar(archivo if archivo.is_absolute() else CARPETA / archivo)
        else:
            print("Uso:\n  python spotify_playlists.py token\n  python spotify_playlists.py actualizar\n"
                  "  python spotify_playlists.py exportar\n  python spotify_playlists.py cargar archivo.xlsx")
    except ErrorSpotify as e:
        log(f"ERROR de Spotify ({e.status}): {e}")
        sys.exit(1)
