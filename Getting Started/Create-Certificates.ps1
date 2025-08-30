<#
.SYNOPSIS
    This script creates a self-signed certificate for use with Microsoft Graph API.
.DESCRIPTION
    This script creates a self-signed certificate and exports it to a .pfx file and the public key to a .cer file.
    It can optionally keep the certificate in the certificate store for thumbprint-based authentication.
.PARAMETER CertificateName
    The name of the certificate to create.
.PARAMETER CertificatePassword
    The password to protect the certificate.
.PARAMETER CertificatePath
    The path to the folder where the certificate will be saved.
.PARAMETER RemoveCert
    Switch parameter to remove the certificate from the certificate store after exporting.
.PARAMETER CreatePasswordFile
    Switch parameter to create an encrypted password file for secure password storage.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CertificateName,
    
    [Parameter(Mandatory = $true)]
    [string]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [string]$CertificatePath = '.',

    [Parameter(Mandatory = $false)]
    [switch]$RemoveCert = $false,

    [Parameter(Mandatory = $false)]
    [switch]$CreatePasswordFile = $false
)

$cert = New-SelfSignedCertificate -DnsName $CertificateName -CertStoreLocation "cert:\CurrentUser\My" -KeyLength 4096
if ($null -eq $cert) {
    Write-Error "Failed to create certificate."
    exit 1
}
else {
    Write-Host "Certificate created successfully!" -ForegroundColor Green
    Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Yellow
    Write-Host "Certificate Subject: $($cert.Subject)" -ForegroundColor Yellow
    
    # Export PFX and CER files
    $cert | Export-PfxCertificate -FilePath "$CertificatePath\$CertificateName.pfx" -Password (ConvertTo-SecureString -String $CertificatePassword -Force -AsPlainText)
    $cert | Export-Certificate -FilePath "$CertificatePath\$CertificateName.cer"
    
    # Create encrypted password file if requested
    if ($CreatePasswordFile) {
        $securePassword = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
        $securePassword | ConvertFrom-SecureString | Out-File "$CertificatePath\$CertificateName.cred"
        Write-Host "Encrypted password file created: $CertificatePath\$CertificateName.cred" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Files created:" -ForegroundColor Cyan
    Write-Host "  - $CertificatePath\$CertificateName.pfx (Private key + certificate - keep secure!)" -ForegroundColor White
    Write-Host "  - $CertificatePath\$CertificateName.cer (Public certificate for Azure upload)" -ForegroundColor White
    if ($CreatePasswordFile) {
        Write-Host "  - $CertificatePath\$CertificateName.cred (Encrypted password file)" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "SECURITY RECOMMENDATIONS:" -ForegroundColor Red -BackgroundColor Yellow
    Write-Host ""
    if (-not $RemoveCert) {
        Write-Host "RECOMMENDED: Use certificate thumbprint authentication (most secure):" -ForegroundColor Green
        Write-Host "  -CertificateThumbprint '$($cert.Thumbprint)'" -ForegroundColor White
        Write-Host ""
        Write-Host "  This method does not require password storage and is the most secure option." -ForegroundColor Yellow
        Write-Host "  The certificate remains in your Windows Certificate Store." -ForegroundColor Yellow
        Write-Host ""
    }
    
    if ($CreatePasswordFile) {
        Write-Host "ALTERNATIVE: Use encrypted password file:" -ForegroundColor Yellow
        Write-Host "  -CertificatePath '$CertificatePath\$CertificateName.pfx'" -ForegroundColor White
        Write-Host "  -CertificatePasswordFile '$CertificatePath\$CertificateName.cred'" -ForegroundColor White
        Write-Host ""
    }
    
    Write-Host "AVOID: Plaintext password in scripts (security risk):" -ForegroundColor Red
    Write-Host "  -CertificatePassword 'plaintext_password'" -ForegroundColor White
    Write-Host ""
    
    if ($RemoveCert) {
        Remove-Item -Path "cert:\CurrentUser\My\$($cert.Thumbprint)" -Force
        Write-Host "Certificate removed from certificate store as requested." -ForegroundColor Yellow
        Write-Host "Note: Thumbprint authentication is no longer available." -ForegroundColor Yellow
    }
}
