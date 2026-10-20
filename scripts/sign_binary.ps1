# scripts/sign_binary.ps1
# Creates or reuses a local CodeSigning certificate for RandomOffice,
# registers it in CurrentUser\TrustedPublisher and signs the executable
# so Windows Defender and SmartScreen recognise it as a trusted application.

param(
    [string]$TargetExe = "dist\standalone\RandomOffice.exe"
)

$ErrorActionPreference = "Stop"

$fullPath = Resolve-Path $TargetExe
Write-Host "Target executable: $fullPath"

# Find or create self-signed code signing cert
$existing = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Subject -eq "CN=RandomOffice" } | Select-Object -First 1

if (-not $existing) {
    Write-Host "Creating new local Code Signing certificate for CN=RandomOffice..."
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=RandomOffice" -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(5)
} else {
    Write-Host "Using existing Code Signing certificate: $($existing.Thumbprint)"
    $cert = $existing
}

# Trust in CurrentUser\TrustedPublisher so SmartScreen treats it as trusted
$tpStore = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::TrustedPublisher, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
$tpStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
if (-not ($tpStore.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint })) {
    $tpStore.Add($cert)
    Write-Host "Added certificate to CurrentUser\TrustedPublisher."
}
$tpStore.Close()

# Sign the binary
$sig = Set-AuthenticodeSignature -FilePath $fullPath -Certificate $cert -HashAlgorithm SHA256
Write-Host "Signature status: $($sig.Status) ($($sig.StatusMessage))"

# Unblock file
Unblock-File -Path $fullPath
Write-Host "Unblocked file: $fullPath"
