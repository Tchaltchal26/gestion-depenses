<#
.SYNOPSIS
  Construit le classeur SG (SG_Application.xlsm) a partir des sources VBA.

.DESCRIPTION
  1. Cree un classeur Excel vierge ;
  2. importe les modules src/vba/*.bas (convertis en Windows-1252 + CRLF,
     l'editeur VBA n'acceptant pas l'UTF-8) ;
  3. insere le code de ThisWorkbook ;
  4. enregistre au format .xlsm ;
  5. lance SG_Installer (tables, interface, formulaire de connexion,
     creation du compte administrateur).

  Prerequis : Excel de bureau pour Windows ; option activee une seule fois :
  Fichier > Options > Centre de gestion de la confidentialite > Parametres
  du Centre > Parametres des macros > "Acces approuve au modele d'objet du
  projet VBA". (Desactivez-la apres l'installation.)

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File installer\Installer-SG.ps1 -Destination C:\SG\SG_Application.xlsm -Mode INDEPENDANT

.EXAMPLE
  ... -Tests   # construit un classeur de TEST et lance modTests (efface les donnees du classeur construit)
#>
param(
    [string]$Destination = (Join-Path (Get-Location) 'SG_Application.xlsm'),
    [ValidateSet('INDEPENDANT', 'CENTRALISE')][string]$Mode = 'INDEPENDANT',
    [switch]$Demo,
    [switch]$Tests
)

$ErrorActionPreference = 'Stop'
$source = Resolve-Path (Join-Path $PSScriptRoot '..\src\vba')
if (Test-Path $Destination) { throw "Le fichier $Destination existe deja : choisir une autre destination (aucun ecrasement)." }

$temp = Join-Path ([IO.Path]::GetTempPath()) ("sg_install_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
$ansi = [Text.Encoding]::GetEncoding(1252)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Add()
    try { $null = $wb.VBProject.VBComponents.Count }
    catch { throw "Acces au projet VBA refuse : activer 'Acces approuve au modele d'objet du projet VBA' (voir l'aide du script)." }

    $exclus = @()
    if (-not $Tests) { $exclus += 'modTests.bas' }
    Get-ChildItem $source -Filter '*.bas' | Where-Object { $_.Name -notlike 'zz*' -and $exclus -notcontains $_.Name } | ForEach-Object {
        $texte = [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
        $cible = Join-Path $temp $_.Name
        [IO.File]::WriteAllText($cible, $texte, $ansi)
        $null = $wb.VBProject.VBComponents.Import($cible)
        Write-Host "Module importe : $($_.Name)"
    }
    $codeTW = [IO.File]::ReadAllText((Join-Path $source 'ThisWorkbook.code'), [Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
    $tw = $wb.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
    $tw.DeleteLines(1, $tw.CountOfLines)
    $tw.AddFromString(($codeTW -replace '(?m)^Option Explicit\r?\n', ''))

    $wb.SaveAs($Destination, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled

    if ($Tests) {
        # Fixtures de licence (cle de TEST jetable) : CLE=base64 -> "cle<31>valeur<30>..."
        $fixtures = @()
        Get-Content (Join-Path $PSScriptRoot '..\tools\tests\fixtures-licences.txt') -Encoding UTF8 | Where-Object { $_ -match '^[A-Z_]+=' } | ForEach-Object {
            $cle, $val = $_ -split '=', 2
            $fixtures += $cle + [char]31 + [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($val))
        }
        $rapport = $excel.Run("'$($wb.Name)'!SG_ExecuterTests", '', ($fixtures -join [char]30))
        $chemin = [IO.Path]::ChangeExtension($Destination, '.rapport-tests.txt')
        [IO.File]::WriteAllText($chemin, $rapport, [Text.Encoding]::UTF8)
        Write-Host $rapport
        Write-Host "Rapport : $chemin"
    }
    else {
        $excel.Run("'$($wb.Name)'!SG_Installer", $Mode)
        if ($Demo) { Write-Host ($excel.Run("'$($wb.Name)'!SG_ChargerDemo")) }
    }
    $wb.Save()
    Write-Host "Classeur cree : $Destination"
}
finally {
    Remove-Item -Recurse -Force $temp -ErrorAction SilentlyContinue
}
