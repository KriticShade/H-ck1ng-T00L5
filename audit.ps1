# =======================================================================
#  AUDIT WINDOWS - SCRIPT FINAL PROPRE, UNIFIÉ & MULTI-UTILISATEURS
# =======================================================================

# ------------------------------- Fonctions ------------------------------

function Convert-ToLongPath($path) {
    $fs = New-Object -ComObject Scripting.FileSystemObject
    try { return $fs.GetFile($path).Path }
    catch {
        try { return $fs.GetFolder($path).Path }
        catch { return $path }
    }
}

function Remove-SafeFile($path) {
    $p = Convert-ToLongPath $path
    $obj = New-Object System.IO.FileInfo($p)
    if ($obj.Exists) { try { $obj.Delete() } catch { } }
}

# =======================================================================
#  1. Dossier d'audit
# =======================================================================

# Nom obligatoire
do { $Name = Read-Host "Nom du fichier :" } 
while ([string]::IsNullOrWhiteSpace($Name))

$Name = $Name -replace '[\\/:*?"<>|]', '_'
$Date = Get-Date -Format "yyyyMMdd_HHmm"
$Output = "$PSScriptRoot\Audit_${Date}_$Name"

New-Item -ItemType Directory -Path $Output -Force | Out-Null

Write-Host "`n=== Début de l'audit : $FolderName ===`n"

# Sqlite dans le dossier du script
$sqliteExe = Join-Path $PSScriptRoot "sqlite3.exe"

# =======================================================================
#  2. Audit MULTI-UTILISATEURS (Chrome, Firefox, Applications locales)
# =======================================================================

Write-Host "`n=== Audit multi-utilisateurs ===`n"

$UserProfiles = Get-ChildItem "C:\Users" -Directory |
    Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") }

foreach ($u in $UserProfiles) {

    $UserName = $u.Name
    Write-Host "`n--- Utilisateur : $UserName ---"

    # ===================================================================
    #  Firefox (historique + extensions)
    # ===================================================================

    $ffBase = Join-Path $u.FullName "AppData\Roaming\Mozilla\Firefox\Profiles"
    if (Test-Path $ffBase) {

        $ffProfile = Get-ChildItem $ffBase -Directory |
            Where-Object { $_.Name -match "default-release" } |
            Select-Object -First 1

        if ($ffProfile) {

            # ------------------------ Historique -------------------------

            $places = Join-Path $ffProfile.FullName "places.sqlite"
            if (Test-Path $places) {

                $tmpPlaces = "$env:TEMP\places_$UserName.sqlite"
                Copy-Item $places $tmpPlaces -Force

                $out = Join-Path $Output "Firefox_Historique_$UserName.csv"

                $query = @"
.open $tmpPlaces
SELECT datetime(moz_historyvisits.visit_date/1000000,'unixepoch') AS date,
       moz_places.url AS url
FROM moz_places
JOIN moz_historyvisits ON moz_places.id = moz_historyvisits.place_id;
"@

                $query | & "$sqliteExe" -header -csv > $out
                Remove-Item $tmpPlaces -Force
                Write-Host "Firefox Historique -> $UserName"
            }

            # ------------------------ Extensions -------------------------

            $extJson = Join-Path $ffProfile.FullName "extensions.json"
            if (Test-Path $extJson) {

                $json = Get-Content $extJson -Raw | ConvertFrom-Json

                $exts = foreach ($ext in $json.addons) {

                    if ($ext.location -in ("app","system","builtin","app-system-defaults","app-builtin")) { continue }
                    if ($ext.rootURI -like "resource://*" -or $ext.rootURI -like "chrome://*") { continue }

                    [PSCustomObject]@{
                        Nom         = $ext.defaultLocale.name
                        Version     = $ext.version
                        ID          = $ext.id
                        Type        = $ext.type
                        Description = $ext.defaultLocale.description
                        Homepage    = $ext.homepageURL
                        Actif       = $ext.active
                        Utilisateur = $UserName
                        Source      = "Firefox"
                    }
                }

                $exts | Export-Csv (Join-Path $Output "Firefox_Extensions_$UserName.csv") -NoTypeInformation
                Write-Host "Firefox Extensions -> $UserName"
            }
        }
    }

    # ===================================================================
    #  Chrome (historique + extensions)
    # ===================================================================

    $chromeBase = Join-Path $u.FullName "AppData\Local\Google\Chrome\User Data\Default"
    if (Test-Path $chromeBase) {

        # -------------------- Historique Chrome ------------------------

        $chromeHist = Join-Path $chromeBase "History"
        if (Test-Path $chromeHist) {
            $tmp = "$env:TEMP\chrome_hist_$UserName.sqlite"
            Copy-Item $chromeHist $tmp -Force

            $outChrome = Join-Path $Output "Chrome_Historique_$UserName.csv"

            $query = @"
.open $tmp
SELECT datetime((visits.visit_time/1000000)-11644473600,'unixepoch') AS date,
       urls.url AS url
FROM urls
JOIN visits ON urls.id = visits.url;
"@

            $query | & "$sqliteExe" -header -csv > $outChrome
            Remove-Item $tmp -Force
            Write-Host "Chrome Historique -> $UserName"
        }

        # -------------------- Extensions Chrome ------------------------

        $chromeExtPath = Join-Path $chromeBase "Extensions"
        if (Test-Path $chromeExtPath) {

            $ChromeExts = foreach ($ext in Get-ChildItem $chromeExtPath -Directory) {

                $ver = Get-ChildItem $ext.FullName -Directory | Sort-Object -Descending Name | Select-Object -First 1
                if (!$ver) { continue }

                $manifest = Join-Path $ver.FullName "manifest.json"
                if (!(Test-Path $manifest)) { continue }

                $json = Get-Content $manifest -Raw | ConvertFrom-Json

                [PSCustomObject]@{
                    Nom         = $json.name
                    Version     = $json.version
                    ID          = $ext.Name
                    Type        = $json.manifest_version
                    Description = $json.description
                    Homepage    = $json.homepage_url
                    Actif       = $true
                    Utilisateur = $UserName
                    Source      = "Chrome"
                }
            }

            $ChromeExts | Export-Csv (Join-Path $Output "Chrome_Extensions_$UserName.csv") -NoTypeInformation
            Write-Host "Chrome Extensions -> $UserName"
        }
    }
}

# ===================================================================
#  Applications locales utilisateur
# ===================================================================

Write-Host "`n* Collecte des applications locales + globales..."

$AllApps = @()

# --- Applications locales ---
foreach ($u in Get-ChildItem "C:\Users" -Directory) {
    $path = Join-Path $u.FullName "AppData\Local\Programs"
    if (Test-Path $path) {

        $AllApps += Get-ChildItem $path -Recurse -File -Include *.exe -ErrorAction SilentlyContinue |
            ForEach-Object {

                $v = $_.VersionInfo
                $nom = if ($v.ProductName -and $v.ProductName.Trim() -ne "") { 
                            $v.ProductName 
                       } else { 
                            $_.BaseName 
                       }

                [PSCustomObject]@{
                    Nom          = $nom
                    Version      = $v.ProductVersion
                    Editeur      = $v.CompanyName
                    Localisation = $_.DirectoryName
                    Utilisateur  = $u.Name
                    Type         = "LocalUser"
                }
            }
    }
}

# --- Applications globales ---
$regPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$AllApps += foreach ($r in $regPaths) {
    Get-ItemProperty $r -ErrorAction SilentlyContinue |
        Where-Object DisplayName |
        ForEach-Object {
            [PSCustomObject]@{
                Nom          = $_.DisplayName
                Version      = $_.DisplayVersion
                Editeur      = $_.Publisher
                Localisation = "Registry"
                Utilisateur  = "N/A"
                Type         = "Global"
            }
        }
}

# --- Export ---
$AllApps | Sort-Object Nom, Version -Unique |
    Export-Csv (Join-Path $Output "Applications.csv") -NoTypeInformation -Delimiter ";"

Write-Host "OK -> Applications.csv"

# =======================================================================
#  4. Processus actifs
# =======================================================================

Write-Host "* Processus actifs..."

$SystemProc = @(
    "System", "Idle", "Registry", "smss", "csrss", "wininit", "services",
    "lsass", "svchost", "winlogon", "dwm", "fontdrvhost", "WmiPrvSE",
    "conhost", "SearchIndexer", "ShellExperienceHost", "StartMenuExperienceHost",
    "RuntimeBroker", "taskhostw", "audiodg", "spoolsv", "explorer",
    "sihost", "ctfmon", "SecurityHealthService", "SearchApp",
    "ApplicationFrameHost", "MicrosoftEdgeUpdate", "msedgewebview2"
)

$proc = Get-Process | Where-Object {
    -not ($SystemProc -contains $_.Name) -and
    $_.Name -notmatch "Microsoft|Windows|Runtime"
} | Select-Object Name,Id,CPU,
    @{Name="RAM_MB";Expression={[math]::Round($_.WS / 1MB,2)}},
    StartTime

$proc | Export-Csv (Join-Path $Output "Processus_Actifs.csv") -NoTypeInformation
Write-Host "OK -> Processus_Actifs.csv"

# =======================================================================
#  5. Journaux Windows (Security, PowerShell, Defender)
# =======================================================================

Write-Host "* Journaux Windows..."

Get-WinEvent -LogName Security -MaxEvents 1000 |
    Export-Csv (Join-Path $Output "Events_Security.csv") -NoTypeInformation

Get-WinEvent -LogName "Microsoft-Windows-Windows Defender/Operational" -ErrorAction SilentlyContinue -MaxEvents 1000 |
    Export-Csv (Join-Path $Output "Events_Defender.csv") -NoTypeInformation

# =======================================================================
#  6. Wi-Fi
# =======================================================================

Write-Host "* Wi-Fi..."

$wifi = @()
$profiles = netsh wlan show profiles |
    Select-String "Profil" |
    ForEach-Object {
        if ($_.ToString() -match ":\s*(.+)$") { $matches[1].Trim() }
    }

foreach ($p in $profiles) {

    $det = netsh wlan show profile name="$p" key=clear
    $kv  = @{}

    foreach ($l in $det) {
        if ($l -match "^\s*([^:]+)\s*:\s*(.*)$") {
            $kv[$matches[1].Trim()] = $matches[2].Trim()
        }
    }

    $wifi += [PSCustomObject]@{
        Profil      = $p
        SSID        = $kv["Nom du SSID"]
        Authen      = $kv["Authentification"]
        Chiffrement = $kv["Chiffrement"]
        MotDePasse  = $kv["Contenu de la cle"]
    }
}

$wifi | Export-Csv (Join-Path $Output "Wifi_Details_Complets.csv") -NoTypeInformation
Write-Host "OK -> Wifi_Details_Complets.csv"

# =======================================================================
#  FIN
# =======================================================================

Write-Host "`n=== Audit terminé ===`n"
Write-Host "Dossier créé : $Output"

# =======================================================================
#  GENERATION DU HTML FINAL (COMPACT + DATATABLES)
# =======================================================================

Write-Host "`n=== Génération du HTML avec tri + recherche ===`n"
$HtmlFile = Join-Path $Output "Audit_Summary.html"

# --- Convertit un CSV en table HTML + DataTables ------------------------
function Convert-CsvToHtmlTable {
    param($csvPath)
    if (!(Test-Path $csvPath)) { return "" }

    try { $data = Import-Csv $csvPath -ErrorAction Stop }
    catch { return "<p>Impossible de charger : $csvPath</p>" }

    if ($data.Count -eq 0) { return "<p>Aucune donnée</p>" }

    $id = "tbl_" + ([guid]::NewGuid().ToString().Replace("-", ""))
    $html = "<table id='$id' class='display compact stripe' style='width:100%'><thead><tr>"
    $html += ($data[0].psobject.Properties.Name | ForEach-Object { "<th>$_</th>" }) -join ""
    $html += "</tr></thead><tbody>"

    foreach ($row in $data) {
        $html += "<tr>"
        foreach ($col in $row.psobject.Properties.Name) {
            $html += "<td>" + ($row.$col -replace '<','&lt;' -replace '>','&gt;') + "</td>"
        }
        $html += "</tr>"
    }

    $html += @"
</tbody></table>
<script>
`$(document).ready(function () {
    `$('#$id').DataTable({
        pageLength: 100,
        lengthMenu: [[25,50,100,500,1000,-1],[25,50,100,500,1000,"All"]],
        pagingType: "full_numbers",
        dom: '<"top"lfp>rt<"bottom"lp><"clear">',
        responsive: false
    });
});
</script>
"@


    return $html
}

# --- Organisation des CSV par catégories --------------------------------
$sections = @{
    "Firefox" = @()
    "Chrome" = @()
    "Applications" = @()
    "Processus" = @()
    "Journaux Windows" = @()
    "WiFi" = @()
}

foreach ($csv in Get-ChildItem $Output -Filter *.csv) {
    $n = $csv.Name
    switch -regex ($n) {
        "Firefox"                { $sections["Firefox"] += $csv }
        "Chrome"                 { $sections["Chrome"] += $csv }
        "Applications"           { $sections["Applications"] += $csv }
        "Processus_Actifs"       { $sections["Processus"] += $csv }
        "Events_"                { $sections["Journaux Windows"] += $csv }
        "wifi"                   { $sections["WiFi"] += $csv }   # insensible à la casse
    }
}

# --- HTML HEADER ---------------------------------------------------------
$htmlHeader = @"
<html><head><meta charset='UTF-8'>
<title>Audit Windows - Résumé</title>

<link rel='stylesheet' href='https://cdn.datatables.net/2.0.1/css/dataTables.dataTables.min.css'>
<script src='https://code.jquery.com/jquery-3.7.1.min.js'></script>
<script src='https://cdn.datatables.net/2.0.1/js/dataTables.min.js'></script>

<style>
body{font-family:Arial;background:#f2f2f2;padding:20px;}
.tabs{display:flex;cursor:pointer;margin-bottom:20px;}
.tab{background:#ddd;padding:10px 15px;margin-right:5px;border-radius:5px;}
.tab.active{background:#0075FF;color:#fff;}
.tabcontent{display:none;padding:15px;background:#fff;border-radius:5px;}
.subtitle{font-weight:bold;margin-top:25px;font-size:17px;}
</style>

<script>
function openTab(i){
    let t=document.getElementsByClassName('tab'),
        c=document.getElementsByClassName('tabcontent');
    for(let n=0;n<t.length;n++){ t[n].classList.remove('active'); c[n].style.display='none'; }
    t[i].classList.add('active'); c[i].style.display='block';
}
</script>

</head><body>
<h1>Audit Windows - Résumé Interactif</h1>
<div class='tabs'>
"@

# --- Génération des onglets + contenu -----------------------------------
$htmlTabs = ""
$htmlContent = ""
$index = 0

foreach ($section in $sections.Keys) {

    $htmlTabs += "<div class='tab' onclick='openTab($index)'>$section</div>"

    $block = "<div class='tabcontent'><h2>$section</h2>"

    if ($sections[$section].Count -eq 0) {
        $block += "<p><i>Aucune donnée disponible pour cette section.</i></p>"
    } else {
        foreach ($csv in $sections[$section]) {
            $block += "<div class='subtitle'>$($csv.BaseName)</div>"
            $block += Convert-CsvToHtmlTable $csv.FullName
        }
    }

    $block += "</div>"
    $htmlContent += $block
    $index++
}

# --- FOOTER --------------------------------------------------------------
$htmlFooter = "<script>openTab(0);</script></body></html>"

# --- Ecriture du fichier final -------------------------------------------
Set-Content -Path $HtmlFile -Value ($htmlHeader + $htmlTabs + "</div>" + $htmlContent + $htmlFooter) -Encoding UTF8
Write-Host "HTML généré avec DataTables : $HtmlFile"