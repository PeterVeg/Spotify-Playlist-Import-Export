Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# ============================================================================
# FONCTIONS UTILITAIRES
# ============================================================================

function Get-SpotifyToken {
    param(
        [string]$ProfileName = "default"
    )
    
    $configPath = "C:\Temp\SpotifyOAuth_$ProfileName.json"
    
    if (-not (Test-Path $configPath)) {
        throw "Configuration non trouvée pour le profil '$ProfileName'. Veuillez d'abord configurer ce compte."
    }
    
    $conf = Get-Content $configPath | ConvertFrom-Json
    
    if (-not (Test-Path $conf.TokenPath)) {
        throw "Token non trouvé pour le profil '$ProfileName'. Veuillez d'abord vous authentifier."
    }
    
    $token = Get-Content $conf.TokenPath | ConvertFrom-Json
    $headers = @{ Authorization = "Bearer $($token.access_token)" }
    
    try {
        Invoke-RestMethod -Uri "https://api.spotify.com/v1/me" -Headers $headers -ErrorAction Stop | Out-Null
        return $token.access_token
    }
    catch {
        Write-Host "Access token expiré pour '$ProfileName', rafraîchissement en cours..."
        $body = @{
            grant_type    = "refresh_token"
            refresh_token = $token.refresh_token
            client_id     = $conf.ClientId
            client_secret = $conf.ClientSecret
        }
        
        try {
            $response = Invoke-RestMethod -Method Post -Uri "https://accounts.spotify.com/api/token" -Body $body
            $token.access_token = $response.access_token
            $token.expires_in   = $response.expires_in
            $token.token_type   = $response.token_type
            $token | ConvertTo-Json | Set-Content -Path $conf.TokenPath -Encoding UTF8
            return $token.access_token
        }
        catch {
            throw "Impossible de rafraîchir le token pour '$ProfileName'. Veuillez vous réauthentifier."
        }
    }
}

function Save-SpotifyConfig {
    param(
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$RedirectUri,
        [string]$ProfileName = "default"
    )
    
    New-Item -ItemType Directory -Path C:\Temp -Force | Out-Null
    
    $conf = @{
        ClientId     = $ClientId
        ClientSecret = $ClientSecret
        RedirectUri  = $RedirectUri
        ProfileName  = $ProfileName
        Scopes       = @("playlist-read-private", "playlist-read-collaborative", "playlist-modify-public", "playlist-modify-private")
        TokenPath    = "C:\Temp\spotify_token_$ProfileName.json"
    }
    
    $conf | ConvertTo-Json | Set-Content -Path "C:\Temp\SpotifyOAuth_$ProfileName.json" -Encoding UTF8
}

function Get-AuthorizationUrl {
    param([string]$ProfileName = "default")
    
    $conf = Get-Content "C:\Temp\SpotifyOAuth_$ProfileName.json" | ConvertFrom-Json
    $scopes = [System.Web.HttpUtility]::UrlEncode(($conf.Scopes -join " "))
    return "https://accounts.spotify.com/authorize?client_id=$($conf.ClientId)&response_type=code&redirect_uri=$([System.Web.HttpUtility]::UrlEncode($conf.RedirectUri))&scope=$scopes"
}

function Exchange-CodeForToken {
    param(
        [string]$Code,
        [string]$ProfileName = "default"
    )
    
    $conf = Get-Content "C:\Temp\SpotifyOAuth_$ProfileName.json" | ConvertFrom-Json
    $body = @{
        grant_type    = "authorization_code"
        code          = $Code
        redirect_uri  = $conf.RedirectUri
        client_id     = $conf.ClientId
        client_secret = $conf.ClientSecret
    }
    
    $response = Invoke-RestMethod -Method Post -Uri "https://accounts.spotify.com/api/token" -Body $body
    $response | ConvertTo-Json | Set-Content -Path $conf.TokenPath -Encoding UTF8
}

function Get-UserPlaylists {
    param([string]$ProfileName = "default")
    
    $headers = @{ Authorization = "Bearer $(Get-SpotifyToken -ProfileName $ProfileName)" }
    $allPlaylists = @()
    $offset = 0
    $limit = 50
    
    do {
        $url = "https://api.spotify.com/v1/me/playlists?limit=$limit&offset=$offset"
        $response = Invoke-RestMethod -Uri $url -Headers $headers
        $allPlaylists += $response.items
        $offset += $limit
    } while ($response.items.Count -gt 0)
    
    return $allPlaylists
}

function Get-UserProfile {
    param([string]$ProfileName = "default")
    
    $headers = @{ Authorization = "Bearer $(Get-SpotifyToken -ProfileName $ProfileName)" }
    $profile = Invoke-RestMethod -Uri "https://api.spotify.com/v1/me" -Headers $headers
    return $profile
}

function Get-PlaylistTracks {
    param(
        [string]$PlaylistId,
        [string]$ProfileName = "default"
    )
    
    $headers = @{ Authorization = "Bearer $(Get-SpotifyToken -ProfileName $ProfileName)" }
    $playlistInfo = Invoke-RestMethod -Uri "https://api.spotify.com/v1/playlists/$PlaylistId" -Headers $headers
    
    $allTracks = @()
    $offset = 0
    $limit = 100
    
    do {
        $url = "https://api.spotify.com/v1/playlists/$PlaylistId/tracks?limit=$limit&offset=$offset"
        $response = Invoke-RestMethod -Uri $url -Headers $headers
        $tracks = $response.items | ForEach-Object {
            $track = $_.track
            if ($track) {
                [PSCustomObject]@{
                    Playlist = $playlistInfo.name
                    Titre    = $track.name
                    Artiste  = ($track.artists | ForEach-Object { $_.name }) -join ", "
                    Album    = $track.album.name
                    URI      = $track.uri
                }
            }
        }
        
        $allTracks += $tracks
        $offset += $limit
    } while ($response.items.Count -gt 0)
    
    return @{
        Name = $playlistInfo.name
        Tracks = $allTracks
    }
}

function Export-PlaylistToCSV {
    param(
        [string]$PlaylistId,
        [string]$OutputPath,
        [string]$ProfileName = "default"
    )
    
    $data = Get-PlaylistTracks -PlaylistId $PlaylistId -ProfileName $ProfileName
    $cleanName = ($data.Name -replace '[\\/:*?"<>|]', '_')
    $csvPath = Join-Path $OutputPath "$cleanName.csv"
    
    $data.Tracks | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    return $csvPath
}

function Search-SpotifyTrack {
    param(
        [string]$TrackName,
        [string]$Artist,
        [string]$ProfileName = "default"
    )
    
    $headers = @{ Authorization = "Bearer $(Get-SpotifyToken -ProfileName $ProfileName)" }
    $query = [System.Web.HttpUtility]::UrlEncode("track:$TrackName artist:$Artist")
    $url = "https://api.spotify.com/v1/search?q=$query&type=track&limit=1"
    
    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers
        if ($response.tracks.items.Count -gt 0) {
            return $response.tracks.items[0].uri
        }
    }
    catch {
        return $null
    }
    
    return $null
}

function New-SpotifyPlaylist {
    param(
        [string]$PlaylistName,
        [string]$ProfileName = "default"
    )
    
    $headers = @{ 
        Authorization = "Bearer $(Get-SpotifyToken -ProfileName $ProfileName)"
        "Content-Type" = "application/json"
    }
    
    $profile = Get-UserProfile -ProfileName $ProfileName
    $userId = $profile.id
    
    $body = @{
        name = $PlaylistName
        description = "Importée depuis CSV"
        public = $false
    } | ConvertTo-Json
    
    $url = "https://api.spotify.com/v1/users/$userId/playlists"
    $response = Invoke-RestMethod -Method Post -Uri $url -Headers $headers -Body $body
    
    return $response.id
}

function Add-TracksToPlaylist {
    param(
        [string]$PlaylistId,
        [array]$TrackUris,
        [string]$ProfileName = "default"
    )
    
    $headers = @{ 
        Authorization = "Bearer $(Get-SpotifyToken -ProfileName $ProfileName)"
        "Content-Type" = "application/json"
    }
    
    # Spotify limite à 100 tracks par requête
    $batchSize = 100
    for ($i = 0; $i -lt $TrackUris.Count; $i += $batchSize) {
        $batch = $TrackUris[$i..[Math]::Min($i + $batchSize - 1, $TrackUris.Count - 1)]
        $body = @{ uris = $batch } | ConvertTo-Json
        $url = "https://api.spotify.com/v1/playlists/$PlaylistId/tracks"
        Invoke-RestMethod -Method Post -Uri $url -Headers $headers -Body $body | Out-Null
    }
}

function Get-AvailableProfiles {
    $profiles = Get-ChildItem -Path "C:\Temp\SpotifyOAuth_*.json" -ErrorAction SilentlyContinue | ForEach-Object {
        $profileName = $_.Name -replace 'SpotifyOAuth_', '' -replace '\.json$', ''
        $profileName
    }
    return $profiles
}

# ============================================================================
# INTERFACE GRAPHIQUE
# ============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Spotify Playlist Manager - Multi-Comptes"
$form.Size = New-Object System.Drawing.Size(900, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 18)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Variable globale pour le profil actif
$script:currentProfile = "default"

# ============================================================================
# PANEL PROFIL
# ============================================================================

$profilePanel = New-Object System.Windows.Forms.Panel
$profilePanel.Location = New-Object System.Drawing.Point(10, 10)
$profilePanel.Size = New-Object System.Drawing.Size(860, 60)
$profilePanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$profilePanel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Text = "PROFIL ACTIF:"
$lblProfile.Location = New-Object System.Drawing.Point(10, 10)
$lblProfile.Size = New-Object System.Drawing.Size(120, 25)
$lblProfile.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblProfile.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
$profilePanel.Controls.Add($lblProfile)

$txtProfileName = New-Object System.Windows.Forms.TextBox
$txtProfileName.Location = New-Object System.Drawing.Point(140, 12)
$txtProfileName.Size = New-Object System.Drawing.Size(200, 25)
$txtProfileName.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$txtProfileName.ForeColor = [System.Drawing.Color]::White
$txtProfileName.Text = "default"
$profilePanel.Controls.Add($txtProfileName)

$comboProfiles = New-Object System.Windows.Forms.ComboBox
$comboProfiles.Location = New-Object System.Drawing.Point(360, 12)
$comboProfiles.Size = New-Object System.Drawing.Size(180, 25)
$comboProfiles.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$comboProfiles.ForeColor = [System.Drawing.Color]::White
$comboProfiles.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$profilePanel.Controls.Add($comboProfiles)

$btnDeleteProfile = New-Object System.Windows.Forms.Button
$btnDeleteProfile.Text = "Supprimer"
$btnDeleteProfile.Location = New-Object System.Drawing.Point(550, 10)
$btnDeleteProfile.Size = New-Object System.Drawing.Size(100, 30)
$btnDeleteProfile.BackColor = [System.Drawing.Color]::FromArgb(180, 30, 30)
$btnDeleteProfile.ForeColor = [System.Drawing.Color]::White
$btnDeleteProfile.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDeleteProfile.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnDeleteProfile.Enabled = $false
$profilePanel.Controls.Add($btnDeleteProfile)

$lblProfileInfo = New-Object System.Windows.Forms.Label
$lblProfileInfo.Text = "Non connecté"
$lblProfileInfo.Location = New-Object System.Drawing.Point(670, 15)
$lblProfileInfo.Size = New-Object System.Drawing.Size(180, 25)
$lblProfileInfo.ForeColor = [System.Drawing.Color]::FromArgb(179, 179, 179)
$lblProfileInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$profilePanel.Controls.Add($lblProfileInfo)

$form.Controls.Add($profilePanel)

# ============================================================================
# PANEL CONFIGURATION
# ============================================================================

$configPanel = New-Object System.Windows.Forms.Panel
$configPanel.Location = New-Object System.Drawing.Point(10, 80)
$configPanel.Size = New-Object System.Drawing.Size(860, 150)
$configPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$configPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

$lblConfig = New-Object System.Windows.Forms.Label
$lblConfig.Text = "CONFIGURATION SPOTIFY API"
$lblConfig.Location = New-Object System.Drawing.Point(10, 10)
$lblConfig.Size = New-Object System.Drawing.Size(300, 25)
$lblConfig.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblConfig.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
$configPanel.Controls.Add($lblConfig)

$lblClientId = New-Object System.Windows.Forms.Label
$lblClientId.Text = "Client ID:"
$lblClientId.Location = New-Object System.Drawing.Point(10, 45)
$lblClientId.Size = New-Object System.Drawing.Size(100, 20)
$configPanel.Controls.Add($lblClientId)

$txtClientId = New-Object System.Windows.Forms.TextBox
$txtClientId.Location = New-Object System.Drawing.Point(120, 43)
$txtClientId.Size = New-Object System.Drawing.Size(300, 25)
$txtClientId.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$txtClientId.ForeColor = [System.Drawing.Color]::White
$configPanel.Controls.Add($txtClientId)

$lblClientSecret = New-Object System.Windows.Forms.Label
$lblClientSecret.Text = "Client Secret:"
$lblClientSecret.Location = New-Object System.Drawing.Point(10, 75)
$lblClientSecret.Size = New-Object System.Drawing.Size(100, 20)
$configPanel.Controls.Add($lblClientSecret)

$txtClientSecret = New-Object System.Windows.Forms.TextBox
$txtClientSecret.Location = New-Object System.Drawing.Point(120, 73)
$txtClientSecret.Size = New-Object System.Drawing.Size(300, 25)
$txtClientSecret.PasswordChar = '*'
$txtClientSecret.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$txtClientSecret.ForeColor = [System.Drawing.Color]::White
$configPanel.Controls.Add($txtClientSecret)

$lblRedirect = New-Object System.Windows.Forms.Label
$lblRedirect.Text = "Redirect URI:"
$lblRedirect.Location = New-Object System.Drawing.Point(10, 105)
$lblRedirect.Size = New-Object System.Drawing.Size(100, 20)
$configPanel.Controls.Add($lblRedirect)

$txtRedirect = New-Object System.Windows.Forms.TextBox
$txtRedirect.Location = New-Object System.Drawing.Point(120, 103)
$txtRedirect.Size = New-Object System.Drawing.Size(300, 25)
$txtRedirect.Text = "https://example.org/callback"
$txtRedirect.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$txtRedirect.ForeColor = [System.Drawing.Color]::White
$configPanel.Controls.Add($txtRedirect)

$btnSaveConfig = New-Object System.Windows.Forms.Button
$btnSaveConfig.Text = "Sauvegarder"
$btnSaveConfig.Location = New-Object System.Drawing.Point(450, 43)
$btnSaveConfig.Size = New-Object System.Drawing.Size(180, 35)
$btnSaveConfig.BackColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
$btnSaveConfig.ForeColor = [System.Drawing.Color]::Black
$btnSaveConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSaveConfig.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$configPanel.Controls.Add($btnSaveConfig)

$btnAuth = New-Object System.Windows.Forms.Button
$btnAuth.Text = "Se Connecter"
$btnAuth.Location = New-Object System.Drawing.Point(650, 43)
$btnAuth.Size = New-Object System.Drawing.Size(180, 35)
$btnAuth.BackColor = [System.Drawing.Color]::FromArgb(29, 185, 84)
$btnAuth.ForeColor = [System.Drawing.Color]::White
$btnAuth.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnAuth.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnAuth.Enabled = $false
$configPanel.Controls.Add($btnAuth)

$btnRefreshToken = New-Object System.Windows.Forms.Button
$btnRefreshToken.Text = "Rafraîchir Token"
$btnRefreshToken.Location = New-Object System.Drawing.Point(650, 93)
$btnRefreshToken.Size = New-Object System.Drawing.Size(180, 35)
$btnRefreshToken.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$btnRefreshToken.ForeColor = [System.Drawing.Color]::White
$btnRefreshToken.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefreshToken.Enabled = $false
$configPanel.Controls.Add($btnRefreshToken)

$form.Controls.Add($configPanel)

# ============================================================================
# PANEL PLAYLISTS
# ============================================================================

$playlistPanel = New-Object System.Windows.Forms.Panel
$playlistPanel.Location = New-Object System.Drawing.Point(10, 240)
$playlistPanel.Size = New-Object System.Drawing.Size(860, 450)
$playlistPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$playlistPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

$lblPlaylists = New-Object System.Windows.Forms.Label
$lblPlaylists.Text = "MES PLAYLISTS"
$lblPlaylists.Location = New-Object System.Drawing.Point(10, 10)
$lblPlaylists.Size = New-Object System.Drawing.Size(200, 25)
$lblPlaylists.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblPlaylists.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
$playlistPanel.Controls.Add($lblPlaylists)

$btnLoadPlaylists = New-Object System.Windows.Forms.Button
$btnLoadPlaylists.Text = "Charger Playlists"
$btnLoadPlaylists.Location = New-Object System.Drawing.Point(250, 8)
$btnLoadPlaylists.Size = New-Object System.Drawing.Size(150, 30)
$btnLoadPlaylists.BackColor = [System.Drawing.Color]::FromArgb(29, 185, 84)
$btnLoadPlaylists.ForeColor = [System.Drawing.Color]::White
$btnLoadPlaylists.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLoadPlaylists.Enabled = $false
$playlistPanel.Controls.Add($btnLoadPlaylists)

$listPlaylists = New-Object System.Windows.Forms.ListBox
$listPlaylists.Location = New-Object System.Drawing.Point(10, 50)
$listPlaylists.Size = New-Object System.Drawing.Size(600, 340)
$listPlaylists.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$listPlaylists.ForeColor = [System.Drawing.Color]::White
$listPlaylists.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$listPlaylists.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$playlistPanel.Controls.Add($listPlaylists)

$btnViewTracks = New-Object System.Windows.Forms.Button
$btnViewTracks.Text = "Afficher Titres"
$btnViewTracks.Location = New-Object System.Drawing.Point(630, 50)
$btnViewTracks.Size = New-Object System.Drawing.Size(200, 35)
$btnViewTracks.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$btnViewTracks.ForeColor = [System.Drawing.Color]::White
$btnViewTracks.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnViewTracks.Enabled = $false
$playlistPanel.Controls.Add($btnViewTracks)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Exporter en CSV"
$btnExport.Location = New-Object System.Drawing.Point(630, 100)
$btnExport.Size = New-Object System.Drawing.Size(200, 35)
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
$btnExport.ForeColor = [System.Drawing.Color]::Black
$btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExport.Enabled = $false
$playlistPanel.Controls.Add($btnExport)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = "Importer CSV"
$btnImport.Location = New-Object System.Drawing.Point(630, 150)
$btnImport.Size = New-Object System.Drawing.Size(200, 35)
$btnImport.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
$btnImport.ForeColor = [System.Drawing.Color]::White
$btnImport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnImport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnImport.Enabled = $false
$playlistPanel.Controls.Add($btnImport)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Prêt"
$lblStatus.Location = New-Object System.Drawing.Point(10, 400)
$lblStatus.Size = New-Object System.Drawing.Size(650, 40)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(179, 179, 179)
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$playlistPanel.Controls.Add($lblStatus)

$lblCredit = New-Object System.Windows.Forms.LinkLabel
$lblCredit.Text = "Open source Peter"
$lblCredit.Location = New-Object System.Drawing.Point(690, 410)
$lblCredit.Size = New-Object System.Drawing.Size(150, 20)
$lblCredit.LinkColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
$lblCredit.ActiveLinkColor = [System.Drawing.Color]::FromArgb(50, 255, 116)
$lblCredit.VisitedLinkColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
$lblCredit.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblCredit.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCredit.Add_LinkClicked({
    Start-Process "https://x.com/PeterVeg1"
})
$playlistPanel.Controls.Add($lblCredit)

$form.Controls.Add($playlistPanel)

# ============================================================================
# FONCTIONS AUXILIAIRES UI
# ============================================================================

function Update-ProfileCombo {
    $comboProfiles.Items.Clear()
    $profiles = Get-AvailableProfiles
    if ($profiles) {
        foreach ($profile in $profiles) {
            $comboProfiles.Items.Add($profile) | Out-Null
        }
        if ($comboProfiles.Items.Contains($script:currentProfile)) {
            $comboProfiles.SelectedItem = $script:currentProfile
        }
        elseif ($comboProfiles.Items.Count -gt 0) {
            $comboProfiles.SelectedIndex = 0
            $script:currentProfile = $comboProfiles.SelectedItem
        }
        $btnDeleteProfile.Enabled = ($comboProfiles.Items.Count -gt 0)
    }
    else {
        $btnDeleteProfile.Enabled = $false
    }
}

function Update-ProfileInfo {
    try {
        $profile = Get-UserProfile -ProfileName $script:currentProfile
        $lblProfileInfo.Text = "✓ $($profile.display_name)"
        $lblProfileInfo.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
        $btnLoadPlaylists.Enabled = $true
        $btnImport.Enabled = $true
    }
    catch {
        $lblProfileInfo.Text = "Non connecté"
        $lblProfileInfo.ForeColor = [System.Drawing.Color]::FromArgb(179, 179, 179)
        $btnLoadPlaylists.Enabled = $false
        $btnImport.Enabled = $false
    }
}

function Load-ProfileConfig {
    param([string]$ProfileName)
    
    $configPath = "C:\Temp\SpotifyOAuth_$ProfileName.json"
    if (Test-Path $configPath) {
        try {
            $conf = Get-Content $configPath | ConvertFrom-Json
            $txtClientId.Text = $conf.ClientId
            $txtClientSecret.Text = $conf.ClientSecret
            $txtRedirect.Text = $conf.RedirectUri
            $txtProfileName.Text = $ProfileName
            $btnAuth.Enabled = $true
            $btnRefreshToken.Enabled = $true
            Update-ProfileInfo
            return $true
        }
        catch {
            return $false
        }
    }
    else {
        # Nouveau profil - réinitialiser les champs
        $txtClientId.Clear()
        $txtClientSecret.Clear()
        $txtRedirect.Text = "https://example.org/callback"
        $btnAuth.Enabled = $false
        $btnRefreshToken.Enabled = $false
        $lblProfileInfo.Text = "Non connecté"
        $lblProfileInfo.ForeColor = [System.Drawing.Color]::FromArgb(179, 179, 179)
        return $false
    }
}

# ============================================================================
# GESTION DES ÉVÉNEMENTS
# ============================================================================

# Initialisation
Update-ProfileCombo
if ($comboProfiles.Items.Count -gt 0) {
    Load-ProfileConfig -ProfileName $script:currentProfile
}

# Changement de profil dans le TextBox
$txtProfileName.Add_TextChanged({
    $script:currentProfile = $txtProfileName.Text.Trim()
})

# Changement de profil dans la ComboBox
$comboProfiles.Add_SelectedIndexChanged({
    if ($comboProfiles.SelectedItem) {
        $script:currentProfile = $comboProfiles.SelectedItem.ToString()
        $txtProfileName.Text = $script:currentProfile
        $listPlaylists.Items.Clear()
        $btnViewTracks.Enabled = $false
        $btnExport.Enabled = $false
        
        if (Load-ProfileConfig -ProfileName $script:currentProfile) {
            $lblStatus.Text = "Profil '$script:currentProfile' chargé"
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
        }
        else {
            $lblStatus.Text = "Configuration non trouvée pour ce profil"
            $lblStatus.ForeColor = [System.Drawing.Color]::Orange
        }
    }
})

# Supprimer un profil
$btnDeleteProfile.Add_Click({
    if ($comboProfiles.SelectedItem) {
        $profileToDelete = $comboProfiles.SelectedItem.ToString()
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Voulez-vous vraiment supprimer le profil '$profileToDelete' ?`n`nCette action est irréversible.",
            "Confirmer la suppression",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $configPath = "C:\Temp\SpotifyOAuth_$profileToDelete.json"
                $tokenPath = "C:\Temp\spotify_token_$profileToDelete.json"
                
                if (Test-Path $configPath) { Remove-Item $configPath -Force }
                if (Test-Path $tokenPath) { Remove-Item $tokenPath -Force }
                
                $lblStatus.Text = "Profil '$profileToDelete' supprimé"
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
                
                # Mettre à jour la liste et sélectionner un autre profil
                Update-ProfileCombo
                
                if ($comboProfiles.Items.Count -gt 0) {
                    $comboProfiles.SelectedIndex = 0
                    $script:currentProfile = $comboProfiles.SelectedItem
                    $txtProfileName.Text = $script:currentProfile
                }
                else {
                    $script:currentProfile = "default"
                    $txtProfileName.Text = "default"
                    $txtClientId.Clear()
                    $txtClientSecret.Clear()
                    $btnAuth.Enabled = $false
                    $btnRefreshToken.Enabled = $false
                }
                
                $listPlaylists.Items.Clear()
                $btnViewTracks.Enabled = $false
                $btnExport.Enabled = $false
            }
            catch {
                $lblStatus.Text = "Erreur lors de la suppression: $($_.Exception.Message)"
                $lblStatus.ForeColor = [System.Drawing.Color]::Red
            }
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner un profil dans la liste", "Information", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnSaveConfig.Add_Click({
    try {
        $profileName = $txtProfileName.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($profileName)) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez entrer un nom de profil", "Erreur", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Nettoyer le nom du profil
        $profileName = $profileName -replace '[\\/:*?"<>|]', '_'
        $script:currentProfile = $profileName
        $txtProfileName.Text = $profileName
        
        if ([string]::IsNullOrWhiteSpace($txtClientId.Text) -or 
            [string]::IsNullOrWhiteSpace($txtClientSecret.Text) -or 
            [string]::IsNullOrWhiteSpace($txtRedirect.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez remplir tous les champs de configuration", "Erreur", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        Save-SpotifyConfig -ClientId $txtClientId.Text -ClientSecret $txtClientSecret.Text -RedirectUri $txtRedirect.Text -ProfileName $profileName
        $btnAuth.Enabled = $true
        Update-ProfileCombo
        $comboProfiles.SelectedItem = $profileName
        
        $lblStatus.Text = "Configuration sauvegardée pour le profil '$profileName'"
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
    }
    catch {
        $lblStatus.Text = "Erreur: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$btnAuth.Add_Click({
    try {
        $authUrl = Get-AuthorizationUrl -ProfileName $script:currentProfile
        
        # Créer une fenêtre d'instructions sans bouton OK
        $instructionForm = New-Object System.Windows.Forms.Form
        $instructionForm.Text = "Authentification en cours..."
        $instructionForm.Size = New-Object System.Drawing.Size(450, 200)
        $instructionForm.StartPosition = "Manual"
        $instructionForm.Location = New-Object System.Drawing.Point(
            ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width - 450),
            50
        )
        $instructionForm.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 18)
        $instructionForm.ForeColor = [System.Drawing.Color]::White
        $instructionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
        $instructionForm.TopMost = $true
        $instructionForm.ShowInTaskbar = $false
        $instructionForm.ControlBox = $true
        
        $lblInstructions = New-Object System.Windows.Forms.Label
        $lblInstructions.Text = @"
🔐 Instructions d'authentification :

1. Authentifiez-vous sur Spotify dans le navigateur 
   qui va s'ouvrir

2. Après validation, copiez l'URL COMPLÈTE de la page 
   d'erreur affichée (sélectionnez l'URL et Ctrl+C)

3. Le code sera automatiquement détecté

⏳ L'application surveille votre presse-papier...
"@
        $lblInstructions.Location = New-Object System.Drawing.Point(15, 15)
        $lblInstructions.Size = New-Object System.Drawing.Size(410, 165)
        $lblInstructions.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
        $instructionForm.Controls.Add($lblInstructions)
        
        $instructionForm.Show()
        $instructionForm.BringToFront()
        
        $lblStatus.Text = "Authentification en cours pour '$script:currentProfile'... Copiez l'URL de redirection"
        $lblStatus.ForeColor = [System.Drawing.Color]::Orange
        
        # Ouvrir le navigateur après un court délai
        $openBrowserTimer = New-Object System.Windows.Forms.Timer
        $openBrowserTimer.Interval = 500
        $openBrowserTimer.Add_Tick({
            $openBrowserTimer.Stop()
            Start-Process $authUrl
            $instructionForm.BringToFront()
        })
        $openBrowserTimer.Start()
        
        # Surveiller le presse-papier
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 500
        $script:clipboardTimer = $timer
        $script:instructionForm = $instructionForm
        $previousClipboard = ""
        
        $timer.Add_Tick({
            try {
                $clipboard = [System.Windows.Forms.Clipboard]::GetText()
                
                # Vérifier si le contenu a changé et contient "code="
                if ($clipboard -ne $previousClipboard -and $clipboard -match 'code=([^&\s]+)') {
                    $code = $matches[1]
                    $script:clipboardTimer.Stop()
                    $script:clipboardTimer.Dispose()
                    
                    # Fermer la fenêtre d'instructions
                    if ($script:instructionForm -ne $null) {
                        $script:instructionForm.Close()
                        $script:instructionForm.Dispose()
                    }
                    
                    $lblStatus.Text = "Code détecté! Échange en cours..."
                    $lblStatus.ForeColor = [System.Drawing.Color]::Orange
                    $form.Refresh()
                    
                    try {
                        Exchange-CodeForToken -Code $code -ProfileName $script:currentProfile
                        $lblStatus.Text = "✓ Authentification réussie pour '$script:currentProfile'!"
                        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
                        $btnLoadPlaylists.Enabled = $true
                        $btnRefreshToken.Enabled = $true
                        $btnImport.Enabled = $true
                        Update-ProfileInfo
                        [System.Windows.Forms.MessageBox]::Show("Authentification réussie!", "Succès", 
                            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                    catch {
                        $lblStatus.Text = "Erreur lors de l'échange du code: $($_.Exception.Message)"
                        $lblStatus.ForeColor = [System.Drawing.Color]::Red
                    }
                }
                
                $previousClipboard = $clipboard
            }
            catch {
                # Ignorer les erreurs de lecture du presse-papier
            }
        })
        
        # Gérer la fermeture manuelle de la fenêtre d'instructions
        $instructionForm.Add_FormClosing({
            if ($script:clipboardTimer -ne $null) {
                $script:clipboardTimer.Stop()
                $script:clipboardTimer.Dispose()
                $lblStatus.Text = "Authentification annulée"
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(179, 179, 179)
            }
        })
        
        $timer.Start()
    }
    catch {
        $lblStatus.Text = "Erreur: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$btnRefreshToken.Add_Click({
    try {
        $null = Get-SpotifyToken -ProfileName $script:currentProfile
        $lblStatus.Text = "Token rafraîchi pour '$script:currentProfile'"
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
        $btnLoadPlaylists.Enabled = $true
        $btnImport.Enabled = $true
        Update-ProfileInfo
    }
    catch {
        $lblStatus.Text = "Erreur: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$btnLoadPlaylists.Add_Click({
    try {
        $lblStatus.Text = "Chargement des playlists pour '$script:currentProfile'..."
        $lblStatus.ForeColor = [System.Drawing.Color]::Orange
        $form.Refresh()
        
        $playlists = Get-UserPlaylists -ProfileName $script:currentProfile
        $listPlaylists.Items.Clear()
        
        $script:playlistData = @{}
        foreach ($playlist in $playlists) {
            $displayText = "$($playlist.name) ($($playlist.tracks.total) titres)"
            $listPlaylists.Items.Add($displayText) | Out-Null
            $script:playlistData[$displayText] = $playlist.id
        }
        
        $lblStatus.Text = "$($playlists.Count) playlist(s) chargée(s) pour '$script:currentProfile'"
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
        $btnViewTracks.Enabled = $true
        $btnExport.Enabled = $true
    }
    catch {
        $lblStatus.Text = "Erreur: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$btnViewTracks.Add_Click({
    if ($listPlaylists.SelectedItem) {
        try {
            $selectedPlaylist = $listPlaylists.SelectedItem.ToString()
            $playlistId = $script:playlistData[$selectedPlaylist]
            
            $lblStatus.Text = "Chargement des titres..."
            $lblStatus.ForeColor = [System.Drawing.Color]::Orange
            $form.Refresh()
            
            $data = Get-PlaylistTracks -PlaylistId $playlistId -ProfileName $script:currentProfile
            
            # Créer une fenêtre pour afficher les titres
            $trackForm = New-Object System.Windows.Forms.Form
            $trackForm.Text = "Titres - $($data.Name) [$script:currentProfile]"
            $trackForm.Size = New-Object System.Drawing.Size(1000, 650)
            $trackForm.StartPosition = "CenterScreen"
            $trackForm.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 18)
            
            $trackList = New-Object System.Windows.Forms.DataGridView
            $trackList.Location = New-Object System.Drawing.Point(10, 10)
            $trackList.Size = New-Object System.Drawing.Size(960, 580)
            $trackList.BackgroundColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
            $trackList.ForeColor = [System.Drawing.Color]::White
            $trackList.GridColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
            $trackList.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
            $trackList.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
            $trackList.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
            $trackList.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
            $trackList.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            $trackList.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5)
            $trackList.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
            $trackList.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
            $trackList.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            $trackList.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $trackList.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5)
            $trackList.EnableHeadersVisualStyles = $false
            $trackList.ColumnHeadersHeight = 40
            $trackList.RowTemplate.Height = 35
            $trackList.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
            $trackList.ReadOnly = $true
            $trackList.AllowUserToAddRows = $false
            $trackList.AllowUserToDeleteRows = $false
            $trackList.AllowUserToResizeRows = $false
            $trackList.RowHeadersVisible = $false
            $trackList.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
            $trackList.MultiSelect = $false
            $trackList.AutoGenerateColumns = $true
            
            # Créer un tableau pour DataGridView
            $dt = New-Object System.Data.DataTable
            $dt.Columns.Add("Playlist", [string]) | Out-Null
            $dt.Columns.Add("Titre", [string]) | Out-Null
            $dt.Columns.Add("Artiste", [string]) | Out-Null
            $dt.Columns.Add("Album", [string]) | Out-Null
            $dt.Columns.Add("URI", [string]) | Out-Null
            
            foreach ($track in $data.Tracks) {
                $row = $dt.NewRow()
                $row["Playlist"] = $track.Playlist
                $row["Titre"] = $track.Titre
                $row["Artiste"] = $track.Artiste
                $row["Album"] = $track.Album
                $row["URI"] = $track.URI
                $dt.Rows.Add($row)
            }
            
            $trackList.DataSource = $dt
            
            # Attendre que les colonnes soient créées
            $trackForm.Add_Shown({
                if ($trackList.Columns.Count -gt 0) {
                    $trackList.Columns["Playlist"].Width = 150
                    $trackList.Columns["Titre"].Width = 280
                    $trackList.Columns["Artiste"].Width = 220
                    $trackList.Columns["Album"].Width = 220
                    $trackList.Columns["URI"].Visible = $false
                }
            })
            
            $trackForm.Controls.Add($trackList)
            $trackForm.ShowDialog()
            
            $lblStatus.Text = "$($data.Tracks.Count) titre(s) affiché(s)"
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
        }
        catch {
            $lblStatus.Text = "Erreur: $($_.Exception.Message)"
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'affichage des titres:`n$($_.Exception.Message)", "Erreur", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner une playlist", "Information", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnExport.Add_Click({
    if ($listPlaylists.SelectedItem) {
        try {
            $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderBrowser.Description = "Choisir le dossier d'export"
            $folderBrowser.SelectedPath = "C:\Temp"
            
            if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPlaylist = $listPlaylists.SelectedItem.ToString()
                $playlistId = $script:playlistData[$selectedPlaylist]
                
                $lblStatus.Text = "Export en cours..."
                $lblStatus.ForeColor = [System.Drawing.Color]::Orange
                $form.Refresh()
                
                $csvPath = Export-PlaylistToCSV -PlaylistId $playlistId -OutputPath $folderBrowser.SelectedPath -ProfileName $script:currentProfile
                
                $lblStatus.Text = "Export terminé: $csvPath"
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
                [System.Windows.Forms.MessageBox]::Show("Export terminé avec succès!`n`n$csvPath", "Succès", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
        catch {
            $lblStatus.Text = "Erreur: $($_.Exception.Message)"
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner une playlist", "Information", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnImport.Add_Click({
    try {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Fichiers CSV (*.csv)|*.csv"
        $openFileDialog.Title = "Sélectionner un fichier CSV à importer"
        $openFileDialog.InitialDirectory = "C:\Temp"
        
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $csvPath = $openFileDialog.FileName
            
            # Demander le nom de la nouvelle playlist
            $inputForm = New-Object System.Windows.Forms.Form
            $inputForm.Text = "Nom de la playlist"
            $inputForm.Size = New-Object System.Drawing.Size(400, 150)
            $inputForm.StartPosition = "CenterScreen"
            $inputForm.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 18)
            $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $inputForm.MaximizeBox = $false
            $inputForm.MinimizeBox = $false
            
            $lblInput = New-Object System.Windows.Forms.Label
            $lblInput.Text = "Entrez le nom de la nouvelle playlist:"
            $lblInput.Location = New-Object System.Drawing.Point(10, 10)
            $lblInput.Size = New-Object System.Drawing.Size(370, 25)
            $lblInput.ForeColor = [System.Drawing.Color]::White
            $inputForm.Controls.Add($lblInput)
            
            $txtPlaylistName = New-Object System.Windows.Forms.TextBox
            $txtPlaylistName.Location = New-Object System.Drawing.Point(10, 40)
            $txtPlaylistName.Size = New-Object System.Drawing.Size(360, 25)
            $txtPlaylistName.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
            $txtPlaylistName.ForeColor = [System.Drawing.Color]::White
            $txtPlaylistName.Text = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
            $inputForm.Controls.Add($txtPlaylistName)
            
            $btnOK = New-Object System.Windows.Forms.Button
            $btnOK.Text = "Créer"
            $btnOK.Location = New-Object System.Drawing.Point(190, 75)
            $btnOK.Size = New-Object System.Drawing.Size(90, 30)
            $btnOK.BackColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
            $btnOK.ForeColor = [System.Drawing.Color]::Black
            $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $inputForm.Controls.Add($btnOK)
            
            $btnCancel = New-Object System.Windows.Forms.Button
            $btnCancel.Text = "Annuler"
            $btnCancel.Location = New-Object System.Drawing.Point(290, 75)
            $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
            $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
            $btnCancel.ForeColor = [System.Drawing.Color]::White
            $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $inputForm.Controls.Add($btnCancel)
            
            $inputForm.AcceptButton = $btnOK
            $inputForm.CancelButton = $btnCancel
            
            if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $playlistName = $txtPlaylistName.Text.Trim()
                
                if ([string]::IsNullOrWhiteSpace($playlistName)) {
                    [System.Windows.Forms.MessageBox]::Show("Le nom de la playlist ne peut pas être vide", "Erreur", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }
                
                $lblStatus.Text = "Importation en cours... (Lecture du CSV)"
                $lblStatus.ForeColor = [System.Drawing.Color]::Orange
                $form.Refresh()
                
                # Lire le CSV
                $tracks = Import-Csv -Path $csvPath -Encoding UTF8
                
                $lblStatus.Text = "Création de la playlist '$playlistName'..."
                $form.Refresh()
                
                # Créer la playlist
                $newPlaylistId = New-SpotifyPlaylist -PlaylistName $playlistName -ProfileName $script:currentProfile
                
                # Fenêtre de progression
                $progressForm = New-Object System.Windows.Forms.Form
                $progressForm.Text = "Import en cours"
                $progressForm.Size = New-Object System.Drawing.Size(500, 180)
                $progressForm.StartPosition = "CenterScreen"
                $progressForm.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 18)
                $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $progressForm.MaximizeBox = $false
                $progressForm.MinimizeBox = $false
                $progressForm.ControlBox = $false
                
                $lblProgress = New-Object System.Windows.Forms.Label
                $lblProgress.Text = "Recherche des titres sur Spotify..."
                $lblProgress.Location = New-Object System.Drawing.Point(10, 10)
                $lblProgress.Size = New-Object System.Drawing.Size(470, 30)
                $lblProgress.ForeColor = [System.Drawing.Color]::White
                $lblProgress.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                $progressForm.Controls.Add($lblProgress)
                
                $progressBar = New-Object System.Windows.Forms.ProgressBar
                $progressBar.Location = New-Object System.Drawing.Point(10, 50)
                $progressBar.Size = New-Object System.Drawing.Size(470, 30)
                $progressBar.Maximum = $tracks.Count
                $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $progressForm.Controls.Add($progressBar)
                
                $lblDetails = New-Object System.Windows.Forms.Label
                $lblDetails.Text = "0 / $($tracks.Count) titres traités"
                $lblDetails.Location = New-Object System.Drawing.Point(10, 90)
                $lblDetails.Size = New-Object System.Drawing.Size(470, 50)
                $lblDetails.ForeColor = [System.Drawing.Color]::FromArgb(179, 179, 179)
                $lblDetails.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $progressForm.Controls.Add($lblDetails)
                
                $progressForm.Show()
                $progressForm.Refresh()
                
                # Rechercher et collecter les URIs
                $foundUris = @()
                $notFound = @()
                $counter = 0
                
                foreach ($track in $tracks) {
                    $counter++
                    $progressBar.Value = $counter
                    $lblDetails.Text = "$counter / $($tracks.Count) titres traités`nRecherche: $($track.Titre) - $($track.Artiste)"
                    $progressForm.Refresh()
                    
                    $uri = Search-SpotifyTrack -TrackName $track.Titre -Artist $track.Artiste -ProfileName $script:currentProfile
                    if ($uri) {
                        $foundUris += $uri
                    }
                    else {
                        $notFound += "$($track.Titre) - $($track.Artiste)"
                    }
                    
                    Start-Sleep -Milliseconds 100  # Éviter de surcharger l'API
                }
                
                # Ajouter les titres à la playlist
                if ($foundUris.Count -gt 0) {
                    $lblProgress.Text = "Ajout des titres à la playlist..."
                    $progressForm.Refresh()
                    
                    Add-TracksToPlaylist -PlaylistId $newPlaylistId -TrackUris $foundUris -ProfileName $script:currentProfile
                }
                
                $progressForm.Close()
                
                # Rapport
                $reportMessage = "Import terminé!`n`n"
                $reportMessage += "Playlist: $playlistName`n"
                $reportMessage += "Titres trouvés: $($foundUris.Count) / $($tracks.Count)`n"
                
                if ($notFound.Count -gt 0) {
                    $reportMessage += "`nTitres non trouvés ($($notFound.Count)):`n"
                    $reportMessage += ($notFound | Select-Object -First 10) -join "`n"
                    if ($notFound.Count -gt 10) {
                        $reportMessage += "`n... et $($notFound.Count - 10) autres"
                    }
                }
                
                [System.Windows.Forms.MessageBox]::Show($reportMessage, "Import terminé", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                $lblStatus.Text = "Import terminé: $($foundUris.Count)/$($tracks.Count) titres importés"
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 215, 96)
                
                # Recharger les playlists
                $btnLoadPlaylists.PerformClick()
            }
        }
    }
    catch {
        $lblStatus.Text = "Erreur lors de l'import: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'import:`n$($_.Exception.Message)", "Erreur", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ============================================================================
# AFFICHAGE
# ============================================================================

$form.ShowDialog() | Out-Null