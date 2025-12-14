import time
import spotipy
from spotipy.oauth2 import SpotifyOAuth

print("=" * 60)
print("Spotify Playlist to Liked Songs")
print("=" * 60)
print()

# Запрашиваем учетные данные у пользователя
print("Введите данные для подключения к Spotify API:")
print("(Получите их на https://developer.spotify.com/dashboard)")
print()

CLIENT_ID = input("Client ID: ").strip()
if not CLIENT_ID:
    print("Ошибка: Client ID обязателен!")
    exit(1)

CLIENT_SECRET = input("Client Secret: ").strip()
if not CLIENT_SECRET:
    print("Ошибка: Client Secret обязателен!")
    exit(1)

REDIRECT_URI = input("Redirect URI [по умолчанию: http://127.0.0.1:8888/callback]: ").strip()
if not REDIRECT_URI:
    REDIRECT_URI = "http://127.0.0.1:8888/callback"

print()
PLAYLIST_ID_OR_URL = input("ID или URL плейлиста для копирования: ").strip()
if not PLAYLIST_ID_OR_URL:
    print("Ошибка: ID или URL плейлиста обязателен!")
    exit(1)

print()
print("=" * 60)
print()

# Настройки задержек (можно изменить при необходимости)
DELAY_BETWEEN_REQUESTS = 0.12  # Задержка между запросами (секунды)
ERROR_DELAY = 1.0              # Задержка при ошибке (секунды)

# Scopes: читаем плейлист и сохраняем треки в Liked Songs
scope = "playlist-read-private user-library-modify user-library-read"

sp = spotipy.Spotify(auth_manager=SpotifyOAuth(
    client_id=CLIENT_ID,
    client_secret=CLIENT_SECRET,
    redirect_uri=REDIRECT_URI,
    scope=scope
))

# Получаем все треки плейлиста (в порядке)
tracks = []
results = sp.playlist_items(PLAYLIST_ID_OR_URL, fields="items.track.id,items.track.name,next", additional_types=['track'], limit=100)
while results:
    for item in results['items']:
        track = item.get('track')
        if track and track.get('id'):
            tracks.append((track['id'], track.get('name')))
    if results.get('next'):
        results = sp.next(results)
    else:
        break

print(f"Найдено треков в плейлисте: {len(tracks)}")

# Добавляем по одному с конца, чтобы сохранить порядок
print(f"\nНачинаем добавление {len(tracks)} треков в 'Мне нравится'...\n")
for idx, (track_id, track_name) in enumerate(reversed(tracks), start=1):
    try:
        sp.current_user_saved_tracks_add([track_id])
        print(f"[{idx}/{len(tracks)}] Добавлено: {track_name}")
        time.sleep(DELAY_BETWEEN_REQUESTS)
    except Exception as e:
        print(f"Ошибка при добавлении {track_name} ({track_id}): {e}")
        time.sleep(ERROR_DELAY)

print("Готово — треки добавлены в 'Мне нравится' в нужном порядке.")
