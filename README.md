# Spotify-Playlist-Import-Export
Script Powershell pour exporter et importer ses playlists Spotify

Un outil simple pour **gérer et exporter vos playlists Spotify**.  
Fonctionnalités principales :
- 🎵 Afficher vos playlists personnelles
- 📑 Lister les titres contenus dans chaque playlist
- 📤 Exporter vos playlists vers un fichier CSV
- 👥 Gérer plusieurs profils Spotify
- 📥 Importer des playlists dans un nouveau profil Spotify
- 🔄 Rafraîchir automatiquement le token de connexion

---

## 🚀 Installation

1. **Télécharger le RAW File :**
Ouvrir une console Powershell, allez dans le dossier où vous avez téléchargé le fichier.
   ```bash
  sl .\Downloads\
  .\PlaylistManagerGUI.ps1

2. **Configurer l'application Spotify :**
- Connectez-vous sur Spotify for Developers (https://developer.spotify.com/).
- Créez une application (ex. SpotifyExport).
- Renseignez une URL de callback (ex. https://example.org/callback), ça n'a pas d'importance.
- Récupérez votre Client ID et Client Secret (⚠️ ne les partagez jamais publiquement).

3. **Lancer le script et renseignez les informations**
- Renseignez les ID.
- Entrez un nom de profil
- Sauvegardez
- Chargez la playlist
- Exportez
- Pour importez, avec votre nouveau compte refaites la manip 2.
- Entrez votre nouveau profil
- Sauvegardez
- Importez vos playlists

  Les playlists sont éditable avec Google sheet ou Excel ou notepad (par exemple pour changer le nom de la playlist)
  
