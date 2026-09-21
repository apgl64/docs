<#
.SYNOPSIS
    Script de conversion de documents Word (.docx) vers Markdown (.md) pour le portail Docs (compatible MkDocs & GitHub).
.DESCRIPTION
    Ce script prend un fichier Word (.docx), crée un sous-dossier dédié dans `docs/<NomDoc>/`,
    extrait automatiquement les images dans `assets/`, ajoute l'étiquette 'Publié par' et le cartouche auteur avec avatar,
    et génère les fichiers `index.md` (MkDocs) et `README.md` (GitHub).
.EXAMPLE
    .\convert-doc.ps1 -DocxPath "import\Guide TYPO3.docx"
.EXAMPLE
    .\convert-doc.ps1 -DocxPath "import\Doc.docx" -DocName "Guide utilisateur TYPO3"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$DocxPath,

    [Parameter(Mandatory = $false)]
    [string]$DocName,

    [Parameter(Mandatory = $false)]
    [string]$AssetsDirName = "assets",

    [Parameter(Mandatory = $false)]
    [string]$AuthorName = "Service du Numérique",

    [Parameter(Mandatory = $false)]
    [string]$AuthorRole = "Support & Administration Applicative"
)

# 1. Vérification de la présence de Pandoc
if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
    Write-Error "Pandoc n'est pas installé ou pas présent dans le PATH. Veuillez l'installer."
    exit 1
}

# 2. Si aucun chemin n'est fourni, chercher dans le dossier import/
if ([string]::IsNullOrWhiteSpace($DocxPath)) {
    $ImportDir = Join-Path $PSScriptRoot "import"
    $DocxFiles = Get-ChildItem -Path $ImportDir -Filter "*.docx" | Where-Object { $_.Name -notlike "~$*" }
    
    if ($DocxFiles.Count -eq 0) {
        Write-Host "Aucun fichier .docx trouvé dans le dossier 'import/'." -ForegroundColor Yellow
        Write-Host "Veuillez déposer un fichier .docx dans import/ ou utiliser la commande :" -ForegroundColor Cyan
        Write-Host "  .\convert-doc.ps1 -DocxPath `"chemin/vers/document.docx`"" -ForegroundColor Cyan
        exit 0
    }
    elseif ($DocxFiles.Count -eq 1) {
        $DocxFile = $DocxFiles[0]
        Write-Host "Traitement automatique du fichier trouvé dans import/ : $($DocxFile.Name)" -ForegroundColor Green
    }
    else {
        Write-Host "Plusieurs fichiers .docx trouvés dans import/ :" -ForegroundColor Yellow
        for ($i = 0; $i -lt $DocxFiles.Count; $i++) {
            Write-Host "  [$($i+1)] $($DocxFiles[$i].Name)"
        }
        $Selection = Read-Host "Entrez le numéro du fichier à convertir (1-$($DocxFiles.Count))"
        $Index = [int]$Selection - 1
        $DocxFile = $DocxFiles[$Index]
    }
    $DocxPath = $DocxFile.FullName
} else {
    if (-not (Test-Path $DocxPath)) {
        Write-Error "Le fichier spécifié n'existe pas : $DocxPath"
        exit 1
    }
    $DocxFile = Get-Item $DocxPath
}

# 3. Déterminer le nom de la documentation et la structure de dossiers dans docs/
if ([string]::IsNullOrWhiteSpace($DocName)) {
    $DocName = $DocxFile.BaseName
}

$TargetFolderName = $DocName.Trim()
$DocsBaseDir = Join-Path $PSScriptRoot "docs"
$TargetFolder = Join-Path $DocsBaseDir $TargetFolderName
$AssetsFolder = Join-Path $TargetFolder $AssetsDirName
$OutputFileIndex = Join-Path $TargetFolder "index.md"
$OutputFileReadme = Join-Path $TargetFolder "README.md"

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host " Conversion de : $($DocxFile.Name)" -ForegroundColor Cyan
Write-Host " Destination   : docs/$TargetFolderName" -ForegroundColor Cyan
Write-Host "==========================================`n" -ForegroundColor Cyan

# 4. Création des répertoires cible
if (-not (Test-Path $TargetFolder)) {
    New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
    Write-Host "[+] Création du sous-dossier : docs/$TargetFolderName" -ForegroundColor Green
}

if (-not (Test-Path $AssetsFolder)) {
    New-Item -ItemType Directory -Path $AssetsFolder -Force | Out-Null
    Write-Host "[+] Création du dossier assets : $AssetsDirName" -ForegroundColor Green
}

# 5. Exécution de Pandoc
Write-Host "[*] Conversion avec Pandoc en cours..." -ForegroundColor Yellow

$ExtractMediaArg = "--extract-media=$TargetFolder"
$PandocArgs = @(
    "`"$($DocxFile.FullName)`"",
    "-f", "docx",
    "-t", "gfm",
    "--wrap=none",
    $ExtractMediaArg,
    "-o", "`"$OutputFileIndex`""
)

$Process = Start-Process -FilePath "pandoc" -ArgumentList ($PandocArgs -join " ") -NoNewWindow -Wait -PassThru

if ($Process.ExitCode -eq 0) {
    Write-Host "[OK] Conversion réussie !" -ForegroundColor Green

    # 6. Post-traitement du fichier Markdown (Card Auteur + Nettoyage images)
    if (Test-Path $OutputFileIndex) {
        $RawContent = Get-Content -Path $OutputFileIndex -Raw -Encoding UTF8
        
        # Remplacer les antislashs Windows dans les chemins d'images par des slashs
        $RawContent = $RawContent -replace '\\', '/'
        
        # Déplacer les médias vers assets/ si Pandoc les a extraits dans media/
        $MediaDir = Join-Path $TargetFolder "media"
        if (Test-Path $MediaDir) {
            Get-ChildItem -Path $MediaDir -Recurse -File | ForEach-Object {
                $DestPath = Join-Path $AssetsFolder $_.Name
                Move-Item -Path $_.FullName -Destination $DestPath -Force
            }
            Remove-Item -Path $MediaDir -Recurse -Force
            $RawContent = $RawContent -replace 'media/', "$AssetsDirName/"
        }

        # Date actuelle
        $CurrentDate = (Get-Date).ToString("dd MMMM yyyy", (New-Object System.Globalization.CultureInfo("fr-FR")))

        # Génération du cartouche Auteur
        $HeaderHTML = @"
<!-- Label et Cartouche Auteur -->
<div class="author-label">Publié par</div>
<div class="author-card">
  <img src="../assets/images/default-avatar.svg" alt="Avatar Auteur" class="avatar" />
  <div class="author-details">
    <span class="author-name">$AuthorName</span>
    <span class="author-role">$AuthorRole</span>
    <span class="author-date">Mise à jour : $CurrentDate</span>
  </div>
</div>

"@

        $FinalContent = $HeaderHTML + $RawContent
        Set-Content -Path $OutputFileIndex -Value $FinalContent -Encoding UTF8
        Copy-Item -Path $OutputFileIndex -Destination $OutputFileReadme -Force
    }

    Write-Host "[OK] Fichiers générés : docs/$TargetFolderName/index.md et README.md" -ForegroundColor Green
    Write-Host "[OK] Assets extraits dans : docs/$TargetFolderName/$AssetsDirName/" -ForegroundColor Green
} else {
    Write-Error "Échec de la conversion Pandoc (Code sortie: $($Process.ExitCode))"
}
