<#
.SYNOPSIS
    This script creates a self-signed certificate for use with Microsoft Graph API.
.DESCRIPTION
    This script creates a self-signed certificate and exports it to a .pfx file and the public key to a .cer file.
.PARAMETER CertificateName
    The name of the certificate to create.
.PARAMETER CertificatePassword
    The password to protect the certificate.
.PARAMETER CertificatePath
        The path to the folder where the certificate will be saved.
.PARAMETER RemoveCert
        Switch parameter to remove the certificate from the certificate store after exporting.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CertificateName,
    
    [Parameter(Mandatory = $true)]
    [string]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [string]$CertificatePath = '.',

    [Parameter(Mandatory = $false)]
    [switch]$RemoveCert = $false
)

$cert = New-SelfSignedCertificate -DnsName $CertificateName -CertStoreLocation "cert:\CurrentUser\My" -KeyLength 4096
if ($null -eq $cert) {
    Write-Error "Failed to create certificate."
    exit 1
}
else {
    $cert | Export-PfxCertificate -FilePath "$CertificatePath\$CertificateName.pfx" -Password (ConvertTo-SecureString -String $CertificatePassword -Force -AsPlainText)
    $cert | Export-Certificate -FilePath "$CertificatePath\$CertificateName.cer"
    if ($RemoveCert) {
        Remove-Item -Path "cert:\CurrentUser\My\$($cert.Thumbprint)" -Force
    }
}
