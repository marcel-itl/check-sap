#requires -Version 5
#requires -RunAsAdministrator
#requires -PSEdition Desktop


if ($args.Contains("debug") -or $args.Contains("-debug")) {
    $DebugPreference = "Continue"
    $ErrorActionPreference = "Continue"
}
else {
    $DebugPreference = "SilentlyContinue"
    $ErrorActionPreference = "SilentlyContinue"
}

function Get-SAPVersion() {
    try {
        $SAPObject = Get-WmiObject -Class Win32_Product | Where-Object Name -Like "SAP Crystal Reports*"
        Write-Debug "Get-SapVersion(): SAPObject = $SAPObject"
        $FullVersion = $SAPObject | Select-Object Version
        $ShortVersion = ($FullVersion.Version).ToString()
        Write-Host "Get-SAPVersion(): SAP Crystal Reports version in use: $FullVersion"
        Write-Debug "Get-SapVersion(): Returned $ShortVersion in Substring (0, ShortVersion.Length-5)"
        Return $ShortVersion.Substring(0, $ShortVersion.Length - 5)
    }
    catch {
        #return something so program won't crash when nothing exists
        Return "none"
    }
}

function Compare-SAPVersions([String]$Source) {
    #workaround on the ssl error on some PCs
    Add-Type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                    return true;
                }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    $AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
    [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
    #disable the first run customize for IE so invoke-webrequest works correctly on fresh GUI machines
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 2
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
    Write-Host "Switch-SAPVersions(): Invoking the Web Request. Downloading the current SAP version."
    try {
        Write-Debug "Switch-SAPVersions(): Invoked Web Request to https://magres.pl/materialy-do-pobrania"
        #-SkipCertificateCheck could be redundant due to the previous workaround, check and remove if needed
        #removed due to errors on windows srv
        $DownloadLink = (Invoke-WebRequest -Uri "https://magres.pl/materialy-do-pobrania").Links.Href
        $DownloadLink = $DownloadLink | Select-String -Pattern "https://magres.pl/wp-content/uploads/CRRuntime_$Arch" -SimpleMatch
        $DownloadLink = ($DownloadLink.ToString()).Trim()
        Write-Host -ForegroundColor Cyan "Found download link: $DownloadLink"
    }
    catch {
        Write-Host -ForegroundColor Red "Error occurred: $_"
        Write-Host -ForegroundColor Red "Installation stopped."
        pause
        exit(2)
    }
    #download the new version
    Write-Host "Switch-SAPVersions(): Downloading..."
    $StartTime = Get-Date
    $DestinationPath = "$env:userprofile\Downloads\SAP.msi"
    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadFile($DownloadLink, $DestinationPath)
    $FinishTime = $((Get-Date).Subtract($StartTime).Seconds)
    Unblock-File -Path $DestinationPath
    Write-Host "Switch-SAPVersions(): Download finished in $FinishTime s"
    Write-Debug "Switch-SAPVersions(): Sleeping (5s) to make sure SAP is downloaded before continuing"
    Start-Sleep -Seconds 5
    #uninstall last version
    $AppObject = Get-WmiObject -Class Win32_Product | Where-Object {
        $_.Name -Match "SAP Crystal Reports*"
    }
    if ($AppObject) {
        Write-Host "Switch-SAPVersions(): Uninstall SAP from the local machine"
        $AppObject.Uninstall() | Out-Null
    }
    #install the new version, control the errors
    try {
        Write-Host "Switch-SAPVersions(): Installing SAP..."
        Write-Debug "Switch-SAPVersions(): executed - msiexec /i $DestinationPath /qn+"
        msiexec /i $DestinationPath /qn+
        Start-Sleep -Seconds 15
        while ($true) {
            Write-Debug "SAP still isn't installed. Waiting 5 seconds."
            Start-Sleep -Seconds 5
            if (Get-WmiObject -Class Win32_Product | Where-Object Name -Like "SAP Crystal Reports*") {
                break
            }
        }
        Write-Host "Switch-SAPVersions(): Installation finished. Exiting."
    }
    catch {
        Write-Host -ForegroundColor Black -BackgroundColor Yellow "Error occurred: $_"
        Write-Host -ForegroundColor Black -BackgroundColor Yellow "Try installing manually from path: $DestinationPath"
    }
}


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
    Write-Host -ForegroundColor Green "Used SAP version equals to the version available on the website. Current version: $Ver"
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
