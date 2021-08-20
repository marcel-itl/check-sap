$Version = $PSVersionTable.PSVersion.Major
if ($Version -le 4) {
    Write-Error -Message "Uzywana wersja Powershell jest zbyt przestarzala by uzyc tego skryptu. Zaktualizuj Powershella. Obecna wersja: " + $PSVersionTable.PSVersion
}

if ($args.Contains("debug") -or $args.Contains("-debug")) {
    $DebugPreference = "Continue"
    $ErrorActionPreference = "Continue"
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
    Write-Host "Switch-SAPVersions(): Czekaj..."
    Write-Debug "Switch-SAPVersions(): Sleeping (10s) to make sure SAP is downloaded before continuing"
    Start-Sleep -Seconds 10
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

# SIG # Begin signature block
# MIIFhQYJKoZIhvcNAQcCoIIFdjCCBXICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJqxqgoDuGEBQ0+i/UZvHXI05
# AGKgggMhMIIDHTCCAgWgAwIBAgIQF4gpq1LXWJNB4r1tAfxavTANBgkqhkiG9w0B
# AQsFADAZMRcwFQYDVQQDDA5NQVJDRUwtSVRMT0dJQzAeFw0yMTA4MTIxMjA2NDFa
# Fw0yMjA4MTIxMjI2NDFaMBkxFzAVBgNVBAMMDk1BUkNFTC1JVExPR0lDMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAxLZmksuXy4QcpO2HUxf3L1BR1j6X
# c/qhvI4w3IoGyz/Mg1CEGrX3GrOxwGmBzLBq9YEmXOERkUy56Wm3ju/TwjNhvzB3
# BwdSXG8JGTuOQkjkdSA4R1iIVWyi02QWth692xNndbG0HmCUWUKPkkevpI3bpIIl
# 17NnF0C9z38b7H2uDGu7sjfNAUbiHrINTmQz3b3l7yysnhNWI6fERwA+yy+94UCo
# G4FwWvz1/ALYbYh27sPOnNLPfu7aGCUI4bjGx119FtLL3BTWfJVj1q4YBfLTux09
# qLiImKUZbRZ0g1iJVLtn11PsYWhVwxSJSS0CC5Oce+RxhxzM56ODT+HyIQIDAQAB
# o2EwXzAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGQYDVR0R
# BBIwEIIOTUFSQ0VMLUlUTE9HSUMwHQYDVR0OBBYEFAdOheEYfuZzo6w/oZGmL1Ev
# 6I8OMA0GCSqGSIb3DQEBCwUAA4IBAQCsy1TaoSZhTHFYmr+TwhJFWzZZgIabjyqd
# A++bCDJF0DJawVFCnAP0i8Es1ERnPhrUi6jcTJVXoEzgvBPsQf8GkLopImvisHie
# GLREBlWpV92mEx8rfTE+vjj7CHBXVeBabemLSlYfBSNS9aRUdrL5l4QWDupCezVf
# 27Xu7+QS1lEO2ng5ZqSMtu4X/ZH3SwiAyKbGNJqh9OTRdjhxOUwvnOG6VX2sZiIF
# QzIF5CKWXHdfVIenq950j+0/YBmxY48u2AlunOU3punRZQBwWZozmkj3OVoCTptr
# msAzAQB9Gx959qSoTC8BWRaDvS4WCSISSzFWrW7DfOFqyuSffbLbMYIBzjCCAcoC
# AQEwLTAZMRcwFQYDVQQDDA5NQVJDRUwtSVRMT0dJQwIQF4gpq1LXWJNB4r1tAfxa
# vTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIB
# FTAjBgkqhkiG9w0BCQQxFgQU0S/k/sT3zZ4ZLTByB4ncO92wBkQwDQYJKoZIhvcN
# AQEBBQAEggEAepy2usi2a/QbufdlyUbV1SyKQG2NKkUBlCtp2m4yaSS/8yuzRu3z
# 52qceTGrJ/5TUmDTf3CjuBCajuUg+2+FfxBGu536n/sunqTxxxEKjx7Qn7nLEuVb
# PsheXgS/57Csc5NJYVOIUUEFyNF72dUs54UGClqyu9VhxjBTPBbAaegsc+m21zKw
# KZXI/WjK9l5pr0uwWQmHoBv+3kMHFLghF+44cjSlSk/aoCKElOz2qboeOsI2kcLA
# QPMAD+HokepKbxZKFBq94kD3sfI9Nc7BpNpGnKYZt9QTPFLFY9xqua5T98pzvBiR
# liHcbt247HSR38ZhWMZ8IVi9vxhrHH1hgg==
# SIG # End signature block
