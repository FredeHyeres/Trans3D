# ==============================================================================
# Trans3D - installation automatique (MicroStation V8i SS3)
#
# Copie le projet VBA (Trans3D.mvba) dans Standards\vba du workspace, puis
# configure le chargement automatique du projet au demarrage.
#
# Usage : clic droit sur install.cmd > Executer  (ou .\install.ps1 en PowerShell)
# Relancable sans risque : chaque etape est ignoree si deja faite.
# ==============================================================================

$ErrorActionPreference = "Stop"
Write-Host "=== Installation Trans3D ===" -ForegroundColor Cyan

# --- 0. Emplacements ---------------------------------------------------------
$Source    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Workspace = Join-Path $env:USERPROFILE "Documents\MicroStV8i\WorkSpace"

if (-not (Test-Path $Workspace)) {
    # Autre emplacement possible : installation par defaut de V8i
    $Alt = "C:\ProgramData\Bentley\MicroStation V8i (SELECTseries)\WorkSpace"
    if (Test-Path $Alt) { $Workspace = $Alt }
    else {
        Write-Host "ERREUR : workspace MicroStation introuvable." -ForegroundColor Red
        Write-Host "Cherche : $Workspace"
        Write-Host "Modifiez la variable `$Workspace en tete de script si votre"
        Write-Host "workspace est ailleurs, puis relancez."
        exit 1
    }
}
Write-Host "Workspace : $Workspace"

# --- 1. Installer le projet VBA (Trans3D.mvba dedie) -------------------------
# Le Default.mvba existe deja sur tout MicroStation et ne doit jamais etre
# ecrase : la macro est livree dans son propre projet Trans3D.mvba, copie
# dans le workspace et charge automatiquement via le .ucf (etape 2).
$Mvba = Join-Path $Source "Trans3D.mvba"
$VbaDir = Join-Path $Workspace "Standards\vba"
$MvbaOk = $false
if (Test-Path $Mvba) {
    New-Item -ItemType Directory -Force $VbaDir | Out-Null
    Copy-Item $Mvba $VbaDir -Force
    $MvbaOk = $true
    Write-Host "[OK] Trans3D.mvba copie vers $VbaDir" -ForegroundColor Green
} else {
    Write-Host "[!] Trans3D.mvba absent du dossier d'installation :" -ForegroundColor Yellow
    Write-Host "    importez les fichiers de src\ dans un projet VBA nomme Trans3D,"
    Write-Host "    enregistrez le .mvba a la racine du depot, puis relancez."
}

# --- 1b. Neutraliser les copies masquantes dans les projets MicroStation -----
# MicroStation cherche le .mvba d'abord dans le dossier vba du projet actif
# (WorkSpace\Projects\<projet>\vba) : une vieille copie y masquerait la version
# installee dans Standards\vba. On les renomme en .bak (aucune suppression).
if ($MvbaOk) {
    $ProjetsDir = Join-Path $Workspace "Projects"
    if (Test-Path $ProjetsDir) {
        $Masquantes = @(Get-ChildItem $ProjetsDir -Recurse -Filter "Trans3D.mvba" -File -ErrorAction SilentlyContinue)
        foreach ($M in $Masquantes) {
            $Bak = "$($M.FullName).bak"
            if (Test-Path $Bak) { Remove-Item $Bak -Force }
            Rename-Item $M.FullName $Bak
            Write-Host "[OK] Copie masquante neutralisee : $($M.FullName) -> .bak" -ForegroundColor Yellow
        }
        if ($Masquantes.Count -eq 0) {
            Write-Host "[OK] Aucune copie masquante dans $ProjetsDir" -ForegroundColor Green
        }
    }
}

# --- 2. Chargement automatique du projet (fichier .ucf utilisateur) ----------
if ($MvbaOk) {
    $UsersDir = Join-Path $Workspace "Users"
    $UcfFiles = @()
    if (Test-Path $UsersDir) {
        $UcfFiles = @(Get-ChildItem $UsersDir -Filter *.ucf -File)
    }
    if ($UcfFiles.Count -eq 0) {
        Write-Host "[!] Aucun fichier .ucf dans $UsersDir : configurez l'autoload" -ForegroundColor Yellow
        Write-Host "    dans MicroStation (Utilities > Macros > Project Manager >"
        Write-Host "    clic droit sur Trans3D > Autoload)."
    } else {
        foreach ($UcfFile in $UcfFiles) {
            $Contenu = Get-Content $UcfFile.FullName -Raw
            if ($Contenu -match "MS_VBAAUTOLOADPROJECTS.*Trans3D") {
                Write-Host "[OK] Autoload deja configure dans $($UcfFile.Name)" -ForegroundColor Green
            } else {
                $Lignes = "`r`n# --- Trans3D (ajoute par install.ps1) ---`r`n" + `
                          "MS_VBASEARCHDIRECTORIES < $VbaDir\`r`n" + `
                          "MS_VBAAUTOLOADPROJECTS > Trans3D.mvba`r`n"
                Add-Content -Path $UcfFile.FullName -Value $Lignes -Encoding ASCII
                Write-Host "[OK] Autoload configure dans $($UcfFile.Name)" -ForegroundColor Green
            }
        }
    }
}

# --- 3. Recapitulatif --------------------------------------------------------
Write-Host ""
Write-Host "=== Installation terminee ===" -ForegroundColor Cyan
Write-Host "1. (Re)demarrez MicroStation et ouvrez le fichier 3D de travail"
Write-Host "2. Key-in Convertir : vba run [Trans3D]Convertir"
Write-Host "3. Key-in Semer     : vba run [Trans3D]Semer"
Write-Host "4. Key-in Points    : vba run [Trans3D]Points"
Write-Host "5. Affectation touches : Utilitaires > Touches de fonction (ex. F6/F7/F8)"
Write-Host ""
Read-Host "Appuyez sur Entree pour fermer"
