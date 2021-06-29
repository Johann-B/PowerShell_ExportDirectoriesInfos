############################################################################################################
#
# SCRIPT        : ExportDirectoriesInfos.ps1
#
# DESCRIPTION   : Ce script permet de faire une extraction d'un répertoire ciblé dans un fichier CSV.
#                 Il récupère pour chaque sous-répertoires le nom, la taille et les dates des derniers
#                 accès et écriture.
#
# AUTEUR        : Johann BARON
#
# DATE          : 28/11/2017
#
# MODIFICATIONS :
#
############################################################################################################


# -----------------------------------------
# Paramétrage de l'affichage de la console
# -----------------------------------------

[Console]::ForegroundColor = "Green"
[Console]::BackgroundColor = "Black"
[Console]::WindowHeight = 50
[Console]::BufferHeight = 50
[Console]::WindowWidth = 120
[Console]::BufferWidth = 120
[Console]::Title = "Export de répertoire"
Clear-Host


# ----------------------------------
# Ajout de l'Assembly Windows.Forms
# ----------------------------------

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# ---------------------------
# Fonction du menu principal
# ---------------------------

Function MainMenu {

    Clear-Host
    ""
    ""
    "                         *******************************************"
    "                         *                                         *"
    "                         *          Export de répertoire           *"
    "                         *                                         *"
    "                         *******************************************"
    ""
    ""
    ""
    "                             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    "                             x                                x"
    "                             x              Menu              x"
    "                             x             ------             x"
    "                             x                                x"
    "                             x   1) Sélection du répertoire   x"
    "                             x                                x"
    "                             x   Q) Quitter                   x"
    "                             x                                x"
    "                             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    ""
    ""
    ""
    [Console]::ForegroundColor = "White"

    $Choix = Read-Host "   Choix "

    [Console]::ForegroundColor = "Green"

    if($Choix -eq 1){

    $SelectFolder = New-Object System.Windows.Forms.FolderBrowserDialog
    $SelectFolder.ShowDialog() | Out-Null
    
    $InputPath = $SelectFolder.SelectedPath

        Write-Host "
     L'analyse du répertoire $InputPath est en cours.
     Merci de patienter..." -ForegroundColor Cyan
        Write-Host
        Write-Host
        Start-Sleep -Seconds 2
        AnalyseDirectory

    }elseif($Choix -like "q"){
        Exit

    }else{
        Write-Host "
         Entrée invalide !" -ForegroundColor Red
        Start-Sleep -Seconds 2
        MainMenu
    }

}


# ---------------------------------
# Fonction d'analyse du répertoire
# ---------------------------------

Function AnalyseDirectory{

    $CurrentDate = (Get-Date -UFormat "%Y%m%d_%H%M%S")

    $CurrentLocation = (Get-Location)

    $AnalyzedDirectory = Split-Path $InputPath -Leaf

    $ExportFileName = "$CurrentLocation\$AnalyzedDirectory`_$CurrentDate.csv"


    $ListDir = Get-ChildItem $InputPath
    $MeasureElement = Get-ChildItem $InputPath | Measure-Object
    $NbrElement = $MeasureElement.Count

    $ListDir | select Name,LastWriteTime,LastAccessTime,@{n="Taille";e={
        $DirSize = ((Get-ChildItem $_.fullname -Recurse | Measure-Object Length -Sum).Sum)

        if($DirSize -lt '1024'){$FormattedSize = (("{0:N0}" -f ($DirSize)) -replace "\s","") +" octets"}
        elseif($DirSize -ge '1024' -and $DirSize -lt '1048576'){
            $FormattedSize = (("{0:N2}" -f ($DirSize /1Kb)) -replace "\s","") +" Ko"
        }elseif($DirSize -ge '1048576' -and $DirSize -lt '1073741824'){
            $FormattedSize = (("{0:N2}" -f ($DirSize /1Mb)) -replace "\s","") +" Mo"
        }elseif($DirSize -ge '1073741824' -and $DirSize -lt '1099511627776'){
            $FormattedSize = (("{0:N2}" -f ($DirSize /1Gb)) -replace "\s","") +" Go"
        }else{$FormattedSize = (("{0:N2}" -f ($DirSize /1Tb)) -replace "\s","") +" To"}

        $FormattedSize

        $Compteur = $ListDir.IndexOf($_)+1
        Write-Host " Répertoire " -ForegroundColor White -NoNewline
        Write-Host "$Compteur " -ForegroundColor Yellow -NoNewline
        Write-Host "sur $NbrElement : " -ForegroundColor White -NoNewline
        Write-Host "$_" -ForegroundColor Green -NoNewline
        Write-Host "  -->  $FormattedSize" -ForegroundColor Cyan

    }},Fullname | Export-Csv -Delimiter ";" -NoTypeInformation -Encoding UTF8 $ExportFileName

    $NbLigne = (Get-Content $ExportFileName | measure -Line).Lines -1

    Write-Host
    Write-Host
    Write-Host "=============================================" -ForegroundColor Magenta
    Write-Host
    Write-Host "       Export terminé !" -ForegroundColor Yellow    Write-Host    Write-Host "  $NbLigne sur $NbrElement éléments traités" -ForegroundColor Gray
    Write-Host
    Write-Host "=============================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Host ""
    Write-Host " Le fichier est enregistré sous $ExportFileName"
    Write-Host
    Write-Host
    Write-Host "       ---> Appuyer sur une touche pour continuer <---" -ForegroundColor White    Write-Host    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    MainMenu

}

MainMenu