<#
.SYNOPSIS
    This script creates an encrypted password file for use with GraphContactSync.
.DESCRIPTION
    This script prompts for a password and saves it in an encrypted format that can only be decrypted
    by the same user on the same machine. This provides a secure alternative to plaintext passwords.
.PARAMETER OutputPath
    The path where the encrypted password file will be saved.
.EXAMPLE
    .\Create-EncryptedPassword.ps1 -OutputPath "C:\Certs\certificate.cred"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

Write-Host "Create Encrypted Password File for GraphContactSync" -ForegroundColor Cyan
Write-Host "This utility creates an encrypted password file that can only be decrypted by the same user on the same machine." -ForegroundColor Yellow
Write-Host ""

# Prompt for password securely
$securePassword = Read-Host -Prompt "Enter the certificate password" -AsSecureString

# Encrypt and save the password
try {
    $securePassword | ConvertFrom-SecureString | Out-File $OutputPath
    Write-Host ""
    Write-Host "Encrypted password file created successfully: $OutputPath" -ForegroundColor Green
    Write-Host ""
    Write-Host "Usage in GraphContactSync.ps1:" -ForegroundColor Cyan
    Write-Host "  -CertificatePasswordFile '$OutputPath'" -ForegroundColor White
    Write-Host ""
    Write-Host "Note: This file can only be decrypted by the current user on this machine." -ForegroundColor Yellow
}
catch {
    Write-Error "Failed to create encrypted password file: $($_.Exception.Message)"
    exit 1
}