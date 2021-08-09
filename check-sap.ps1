if ($args.Contains("debug") -or $args.Contains("-debug")) {
    $DebugPreference = "Continue"
}

function Get-ElevationStatus() {
    $CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $ElevationStatus = $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Return $ElevationStatus
}

function Get-SAPVersion() {
    try {
        $SAPObject = Get-WmiObject -Class Win32_Product | Where-Object Name -Like "SAP Crystal Reports*"
        Write-Debug "Get-SapVersion(): SAPObject = $SAPObject"
        $FullVersion = $SAPObject | Select-Object Version
        $ShortVersion = ($FullVersion.Version).ToString()
        Write-Host "Get-SAPVersion(): Uzywana wersja SAP Crystal Reports = $FullVersion"
        Write-Debug "Get-SapVersion(): Returned $ShortVersion in Substring (0, ShortVersion.Length-5)"
        Return $ShortVersion.Substring(0, $ShortVersion.Length - 5)
    }
    catch {
        Return "none"
    }
}

function Compare-SAPVersions([String]$Source) {
    Write-Debug "Compare-SAPVersions($Source):"
    $DownloadLink = (Invoke-WebRequest -Uri "https://magres.pl/materialy-do-pobrania").Links.Href
    Write-Debug "Compare-SAPVersions($Source): DownloadLink = $DownloadLink"
    $DownloadLink = $DownloadLink | Select-String -Pattern "https://magres.pl/wp-content/uploads/CRRuntime_$Arch" -SimpleMatch
    $DownloadLink = ($DownloadLink.ToString()).Trim()
    Write-Debug "Compare-SAPVersions($Source): DownloadLink (trimmed) = $DownloadLink"
    Write-Debug "Compare-SAPVersions(): $DownloadLink contains $Source ? True:False"
    if ($DownloadLink.Contains($Source)) {
        Write-Debug "Compare-SAPVersions(): Returned True"
        Return $True
    }
    else {
        Write-Debug "Compare-SAPVersions(): Returned False"
        Return $False
    }
}

function Switch-SAPVersions() {
    Write-Host "Switch-SAPVersions(): Wywolanie zapytania HTML do strony Magresu, pobranie najnowszej wersji"
    try {
        Write-Debug "Switch-SAPVersions(): Invoked Web Request to https://magres.pl/materialy-do-pobrania"
        $DownloadLink = (Invoke-WebRequest -Uri "https://magres.pl/materialy-do-pobrania").Links.Href
        $DownloadLink = $DownloadLink | Select-String -Pattern "https://magres.pl/wp-content/uploads/CRRuntime_$Arch" -SimpleMatch
        $DownloadLink = ($DownloadLink.ToString()).Trim()
        Write-Host -ForegroundColor Cyan "Znaleziony link do instalacji programu: $DownloadLink"
    }
    catch {
        Write-Host -ForegroundColor Red "Wystapil blad: $_"
        Write-Host -ForegroundColor Red "Instalacja przerwana."
        pause
        exit(2)
    }
    #Pobierz nowa wersje
    $DestinationPath = "$env:userprofile\Downloads\SAP.msi"
    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadFile($DownloadLink, $DestinationPath)
    #Odinstaluj poprzednia wersje
    $AppObject = Get-WmiObject -Class Win32_Product | Where-Object {
        $_.Name -Match "SAP Crystal Reports*"
    }
    if ($AppObject) {
        Write-Host "Switch-SAPVersions(): Odinstalowanie obecnej wersji SAP"
        $AppObject.Uninstall() | Out-Null
    }
    #Zainstaluj nowa wersje / kontroluj bledy
    try {
        Write-Host "Switch-SAPVersions(): Instalacja nowej wersji SAPa"
        Write-Debug "Switch-SAPVersions(): executed - msiexec /i $DestinationPath /qn+"
        msiexec /i $DestinationPath /qn+
        Start-Sleep -Seconds 15
        Write-Host -ForegroundColor Green "Czesc skryptowa zakonczona. Za niedlugo pojawi sie monit o pomyslnym zainstalowaniu SAPa."
    }
    catch {
        Write-Host -ForegroundColor Black -BackgroundColor Yellow "Wystapil blad: $_"
        Write-Host -ForegroundColor Black -BackgroundColor Yellow "Zainstaluj aplikacje recznie z lokalizacji $DestinationPath"
    }
}
if (Get-ElevationStatus) {
    if (Test-Path 'Env:ProgramFiles(x86)') {
        $Arch = "64bit"
    }
    else {
        $Arch = "32bit"
    }
    $Ver = Get-SAPVersion
    Write-Debug "Start(): Ver = $Ver"
    $Ver = $Ver.Replace(".", "_")
    Write-Debug "Start(): Ver = $Ver"
    $VerRC = Compare-SAPVersions($Ver)
    if ($VerRC -eq $False) {
        Switch-SAPVersions
        Start-Sleep(15)
    }
    else {
        Write-Host -ForegroundColor Green "Uzywana wersja SAP Crystal Reports jest zgodna z wersja dostepna na stronie. Uzywana wersja: $Ver"
    }
}
else {
    Write-Host -ForegroundColor Red "Uruchom skrypt jako administrator!"
    pause
    exit(2)
}
