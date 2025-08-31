<#
.SYNOPSIS
    This script will sync contacts from the GAL to 1:many mailboxes.
.DESCRIPTION
    This script will sync contacts from one mailbox to another. It will copy contacts from the source mailbox to the target mailbox. If the contact already exists in the target mailbox, it will be updated.
.PARAMETER ExchangeOrg
    The Exchange Organization to connect to.
.PARAMETER ClientID
    The Client ID for the application.
.PARAMETER CertificatePath
    The path to the certificate PFX file. Required when using PFX file authentication.
.PARAMETER CertificateThumbprint
    The thumbprint of the certificate installed in the certificate store. Use this for secure authentication without passwords.
.PARAMETER CertificatePassword
    The password for the PFX certificate file. Use only when CertificateThumbprint is not specified.
.PARAMETER CertificatePasswordFile
    The path to an encrypted password file for the PFX certificate. Alternative to plaintext password.
.PARAMETER MailboxList
    The list of mailboxes to sync contacts to.
.PARAMETER ManagedContactFolderName
    The name of the folder to sync contacts to.
.PARAMETER LogPath
    The path to the log file.
.PARAMETER FileAsFormat
    The format for the FileAs field. Valid values are "FirstLast" (default) or "LastFirst".
.PARAMETER Categories
    Optional array of categories to assign to contacts. Useful when syncing to main Contacts folder.
#>

Param(    
    [Parameter(Mandatory = $true)]
    [string]$ExchangeOrg,
    
    [Parameter(Mandatory = $true)]
    [string]$ClientID,
    
    [Parameter(Mandatory = $false)]
    [System.IO.FileInfo]$CertificatePath,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificatePassword,
    
    [Parameter(Mandatory = $false)]
    [System.IO.FileInfo]$CertificatePasswordFile,

    [Parameter(Mandatory = $true)]
    [string]$MailboxList,
    
    [Parameter(Mandatory = $true)]
    [string]$ManagedContactFolderName,

    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("FirstLast", "LastFirst")]
    [string]$FileAsFormat = "FirstLast",

    [Parameter(Mandatory = $false)]
    [string[]]$Categories = @()
)

# Parameter validation
if (-not $CertificateThumbprint -and -not $CertificatePath) {
    throw "Either CertificateThumbprint or CertificatePath must be specified."
}

if ($CertificateThumbprint -and $CertificatePath) {
    throw "CertificateThumbprint and CertificatePath cannot both be specified. Choose one authentication method."
}

if ($CertificatePath -and -not $CertificatePassword -and -not $CertificatePasswordFile) {
    throw "When using CertificatePath, either CertificatePassword or CertificatePasswordFile must be specified."
}

Import-Module PoShLog

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

function Sync-ManagedContacts {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,

        [Parameter(Mandatory = $true)]
        [string]$ManagedContactFolderName,

        [Parameter(Mandatory = $true)]
        $ManagedContacts,

        [Parameter(Mandatory = $false)]
        [ValidateSet("FirstLast", "LastFirst")]
        [string]$FileAsFormat = "FirstLast",

        [Parameter(Mandatory = $false)]
        [string[]]$Categories = @()
    )

    # Get the given User's Managed Contact Folder
    Write-InfoLog "Locating Managed Contact folder for $Mailbox"
    $ManagedContactFolder = Get-MgUserContactFolder -UserId $Mailbox -Filter "DisplayName eq '$ManagedContactFolderName'"

    # If the Managed Contact Folder does not exist, create it
    if ($null -eq $ManagedContactFolder) {
        Write-InfoLog "Creating Managed Contact folder"
        $ManagedContactFolder = New-MgUserContactFolder -UserId $Mailbox -DisplayName $ManagedContactFolderName
    }

    $ExistingManagedContacts = Get-MgUserContactFolderContact -UserId $Mailbox -ContactFolderId $ManagedContactFolder.Id -All -ExpandProperty "extensions(`$filter=id eq 'ManagedContactCorrelation'`)"

    # So, now we have the list of contacts in the Managed Contact Folder, and we have the list of contacts that we want to sync. We need to compare the two lists and determine which contacts need to be added, updated, or deleted.

    $ContactsToAdd = @()
    $ContactsToDelete = @()

    $ContactsToAdd += $ManagedContacts | Where-Object { $ExistingManagedContacts.Extensions.AdditionalProperties.CorrelationId -notcontains $_.Id }
    $ContactsToDelete += $ExistingManagedContacts | Where-Object { $ManagedContacts.Id -notcontains $_.Extensions.AdditionalProperties.CorrelationId }
    $ContactsToChecksum = $ExistingManagedContacts | Where-Object { $ManagedContacts.Id -contains $_.Extensions.AdditionalProperties.CorrelationId }

    $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = New-Object -TypeName System.Text.UTF8Encoding
    foreach ($ExistingContact in $ContactsToChecksum) {
        # Get the Managed Contact that corresponds to the existing contact entry in the Mailbox's Managed Contact Folder
        $ManagedContact = $ManagedContacts | Where-Object { $_.Id -eq $ExistingContact.Extensions.AdditionalProperties.CorrelationId }
        #$ExistingContactChecksum = $ExistingContact.Extensions.AdditionalProperties.ContactChecksum

        $ManagedContactChecksumFields = ""
        $ExistingContactChecksumFields = ""
        if ($ManagedContact.EntryType -eq 'User') {
            $ManagedContactChecksumFields = ($ManagedContact | Select-Object -Property `
                @{Name = 'DisplayName'; Expression = { $_.DisplayName ?? "" } }, `
                @{Name = 'GivenName'; Expression = { $_.GivenName ?? "" } }, `
                @{Name = 'Surname'; Expression = { $_.Surname ?? "" } }, `
                @{Name = 'CompanyName'; Expression = { $_.CompanyName ?? "" } }, `
                @{Name = 'JobTitle'; Expression = { $_.JobTitle ?? "" } }, `
                @{Name = 'Department'; Expression = { $_.Department ?? "" } }, `
                @{Name = 'OfficeLocation'; Expression = { $_.OfficeLocation ?? "" } }, `
                @{Name = 'Mail'; Expression = { $_.Mail ?? "" } }, `
                @{Name = 'BusinessPhones'; Expression = { $_.BusinessPhones ?? "" } }, `
                @{Name = 'MobilePhone'; Expression = { $_.MobilePhone ?? "" } }, `
                @{Name = 'StreetAddress'; Expression = { $_.StreetAddress ?? "" } }, `
                @{Name = 'City'; Expression = { $_.City ?? "" } }, `
                @{Name = 'State'; Expression = { $_.State ?? "" } }, `
                @{Name = 'PostalCode'; Expression = { $_.PostalCode ?? "" } }, `
                @{Name = 'Country'; Expression = { $_.Country ?? "" } } `
                | ConvertTo-Json -Depth 10)
            $ExistingContactChecksumFields = ($ExistingContact | Select-Object -Property `
                @{Name = 'DisplayName'; Expression = { $_.DisplayName ?? "" } }, `
                @{Name = 'GivenName'; Expression = { $_.GivenName ?? "" } }, `
                @{Name = 'Surname'; Expression = { $_.Surname ?? "" } }, `
                @{Name = 'CompanyName'; Expression = { $_.CompanyName ?? "" } }, `
                @{Name = 'JobTitle'; Expression = { $_.JobTitle ?? "" } }, `
                @{Name = 'Department'; Expression = { $_.Department ?? "" } }, `
                @{Name = 'OfficeLocation'; Expression = { $_.OfficeLocation ?? "" } }, `
                @{Name = 'Mail'; Expression = { $_.EmailAddresses[0].Address ?? "" } }, `
                @{Name = 'BusinessPhones'; Expression = { $_.BusinessPhones ?? "" } }, `
                @{Name = 'MobilePhone'; Expression = { $_.MobilePhone ?? "" } }, `
                @{Name = 'StreetAddress'; Expression = { $_.BusinessAddress.Street ?? "" } }, `
                @{Name = 'City'; Expression = { $_.BusinessAddress.City ?? "" } }, `
                @{Name = 'State'; Expression = { $_.BusinessAddress.State ?? "" } }, `
                @{Name = 'PostalCode'; Expression = { $_.BusinessAddress.PostalCode ?? "" } }, `
                @{Name = 'Country'; Expression = { $_.BusinessAddress.CountryOrRegion ?? "" } } `
                | ConvertTo-Json -Depth 10)
        }
        elseif ($ManagedContact.EntryType -eq 'Contact') {
            $ManagedContactChecksumFields = ($ManagedContact | Select-Object -Property `
                @{Name = 'DisplayName'; Expression = { $_.DisplayName ?? "" } }, `
                @{Name = 'GivenName'; Expression = { $_.GivenName ?? "" } }, `
                @{Name = 'Surname'; Expression = { $_.Surname ?? "" } }, `
                @{Name = 'CompanyName'; Expression = { $_.CompanyName ?? "" } }, `
                @{Name = 'JobTitle'; Expression = { $_.JobTitle ?? "" } }, `
                @{Name = 'Mail'; Expression = { $_.Mail ?? "" } }, `
                @{Name = 'Mobile'; Expression = { $_.Phones[1].Number ?? "" } }, `
                @{Name = 'BusinessPhone'; Expression = { $_.Phones[2].Number ?? "" } }, `
                @{Name = 'StreetAddress'; Expression = { $_.Addresses[0].Street ?? "" } }, `
                @{Name = 'City'; Expression = { $_.Addresses[0].City ?? "" } }, `
                @{Name = 'State'; Expression = { $_.Addresses[0].State ?? "" } }, `
                @{Name = 'PostalCode'; Expression = { $_.Addresses[0].PostalCode ?? "" } }, `
                @{Name = 'Country'; Expression = { $_.Addresses[0].Country ?? "" } } `
                | ConvertTo-Json -Depth 10)
            $ExistingContactChecksumFields = ($ExistingContact | Select-Object -Property `
                @{Name = 'DisplayName'; Expression = { $_.DisplayName ?? "" } }, `
                @{Name = 'GivenName'; Expression = { $_.GivenName ?? "" } }, `
                @{Name = 'Surname'; Expression = { $_.Surname ?? "" } }, `
                @{Name = 'CompanyName'; Expression = { $_.CompanyName ?? "" } }, `
                @{Name = 'JobTitle'; Expression = { $_.JobTitle ?? "" } }, `
                @{Name = 'Mail'; Expression = { $_.EmailAddresses[0].Address ?? "" } }, `
                @{Name = 'Mobile'; Expression = { $_.MobilePhone ?? "" } }, `
                @{Name = 'BusinessPhone'; Expression = { $_.BusinessPhones[0] ?? "" } }, `
                @{Name = 'StreetAddress'; Expression = { $_.BusinessAddress.Street ?? "" } }, `
                @{Name = 'City'; Expression = { $_.BusinessAddress.City ?? "" } }, `
                @{Name = 'State'; Expression = { $_.BusinessAddress.State ?? "" } }, `
                @{Name = 'PostalCode'; Expression = { $_.BusinessAddress.PostalCode ?? "" } }, `
                @{Name = 'Country'; Expression = { $_.BusinessAddress.CountryOrRegion ?? "" } } `
                | ConvertTo-Json -Depth 10)
        }

        $ManagedContactChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($ManagedContactChecksumFields)))
        $ExistingContactChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($ExistingContactChecksumFields)))

        # Check for photo changes for User contacts
        $PhotoChanged = $false
        if ($ManagedContact.EntryType -eq 'User') {
            # Get current photo metadata
            $CurrentPhotoMeta = Get-MgUserPhoto -UserId $ManagedContact.UserPrincipalName -ProfilePhotoId 120x120 -ErrorAction SilentlyContinue
            $CurrentPhotoChecksum = ""
            if ($CurrentPhotoMeta) {
                # Use mediaETag if available (preferred method), fallback to ID+dimensions for compatibility
                $PhotoFingerprint = $CurrentPhotoMeta.AdditionalProperties["@odata.mediaEtag"]
                if (-not $PhotoFingerprint) {
                    $PhotoFingerprint = "$($CurrentPhotoMeta.Id)_$($CurrentPhotoMeta.Height)x$($CurrentPhotoMeta.Width)"
                }
                $CurrentPhotoChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($PhotoFingerprint)))
            }
            
            # Compare with stored photo checksum
            $StoredPhotoChecksum = $ExistingContact.Extensions.AdditionalProperties.PhotoChecksum ?? ""
            if ($CurrentPhotoChecksum -ne $StoredPhotoChecksum) {
                $PhotoChanged = $true
                Write-VerboseLog "Photo changed for contact: $($ManagedContact.DisplayName) - Old: [$StoredPhotoChecksum] New: [$CurrentPhotoChecksum]"
            }
        }

        Write-VerboseLog "Comparing Managed Contact Checksums: [$ExistingContactChecksum] vs [$ManagedContactChecksum]"

        if ($ExistingContactChecksum -ne $ManagedContactChecksum -or $PhotoChanged) {
            #if this is an edited contact or photo changed, effectively delete the old one and add the new one
            if ($ExistingContactChecksum -ne $ManagedContactChecksum) {
                Write-DebugLog "Detected changed contact: $($ManagedContact.DisplayName)"
            }
            if ($PhotoChanged) {
                Write-DebugLog "Detected photo change for contact: $($ManagedContact.DisplayName)"
                # Remove old photo file to force re-download
                $OldPhotoPath = "Photos\$($ManagedContact.UserPrincipalName).jpg"
                if (Test-Path $OldPhotoPath) {
                    Remove-Item $OldPhotoPath -Force -ErrorAction SilentlyContinue
                    Write-VerboseLog "Removed old photo file: $OldPhotoPath"
                }
            }
            $ContactsToDelete += $ExistingContact
            $ContactsToAdd += $ManagedContact
        }
    }

    foreach ($Contact in $ContactsToDelete) {
        Write-VerboseLog "Deleting contact: $($Contact.DisplayName)"
        Remove-MgUserContactFolderContact -UserId $Mailbox -ContactFolderId $ManagedContactFolder.Id -ContactId $Contact.Id
    }

    foreach ($Contact in $ContactsToAdd) {
        #Only Members have photos
        $ContactPhotoFile = $null
        $PhotoChecksum = ""
        if ($Contact.EntryType -eq 'User') {
            # Get photo metadata to create checksum for change detection
            $PhotoMeta = Get-MgUserPhoto -UserId $Contact.UserPrincipalName -ProfilePhotoId 120x120 -ErrorAction SilentlyContinue
            if ($PhotoMeta) {
                # Use mediaETag if available (preferred method), fallback to ID+dimensions for compatibility
                $PhotoFingerprint = $PhotoMeta.AdditionalProperties["@odata.mediaEtag"]
                if (-not $PhotoFingerprint) {
                    $PhotoFingerprint = "$($PhotoMeta.Id)_$($PhotoMeta.Height)x$($PhotoMeta.Width)"
                }
                $PhotoChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($PhotoFingerprint)))
            }
            
            # Download photo if it doesn't exist (including after photo change detection)
            if (!(Test-Path -PathType Leaf -Path "Photos\$($Contact.UserPrincipalName).jpg")) {
                Write-VerboseLog "Downloading photo for contact: $($Contact.DisplayName)"
                Get-MgUserPhotoContent -UserId $Contact.UserPrincipalName -ProfilePhotoId 120x120 -OutFile "Photos\$($Contact.UserPrincipalName).jpg" -ErrorAction SilentlyContinue
            }
            if ((Test-Path -PathType Leaf -Path "Photos\$($Contact.UserPrincipalName).jpg")) {
                $ContactPhotoFile = "Photos\$($Contact.UserPrincipalName).jpg"
            }
            #$ContactPhoto = Get-Content -Path "$($Contact.UserPrincipalName).jpg" -AsByteStream -ErrorAction SilentlyContinue
        }
      
        if ($Contact.EntryType -eq 'User') {
            $ManagedContactString = ($Contact | Select-Object -Property DisplayName, GivenName, Surname, CompanyName, JobTitle, Department, OfficeLocation, Mail, BusinessPhones, MobilePhone, StreetAddress, City, State, PostalCode, Country | ConvertTo-Json -Depth 10)
        }
        elseif ($Contact.EntryType -eq 'Contact') {
            $ManagedContactString = ($Contact | Select-Object -Property DisplayName, GivenName, Surname, CompanyName, JobTitle, Department, Mail, Phones, Addresses | ConvertTo-Json -Depth 10)
        }
        $ManagedContactChecksum = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($ManagedContactString)))

        Write-VerboseLog "Adding contact: $($Contact.DisplayName) with checksum: $ManagedContactChecksum"

        # Determine the FileAs value based on the format
        $fileAsValue = ""
        if ($FileAsFormat -eq "LastFirst" -and $Contact.Surname -and $Contact.GivenName) {
            $fileAsValue = "$($Contact.Surname), $($Contact.GivenName)"
        }
        elseif ($FileAsFormat -eq "FirstLast" -and $Contact.GivenName -and $Contact.Surname) {
            $fileAsValue = "$($Contact.GivenName) $($Contact.Surname)"
        }
        elseif ($Contact.DisplayName) {
            $fileAsValue = $Contact.DisplayName
        }

        $newContact = @{
            extensions     = @(
                @{
                    "@odata.type" = "microsoft.graph.openTypeExtension"
                    ExtensionName = "ManagedContactCorrelation"
                    CorrelationId = $Contact.Id.ToString()
                    PhotoChecksum = $PhotoChecksum
                    #ContactChecksum = $ManagedContactChecksum
                }
            )
            displayName    = $Contact.DisplayName
            givenName      = $Contact.GivenName
            surname        = $Contact.Surname
            companyName    = $Contact.CompanyName
            jobTitle       = $Contact.JobTitle
            department     = $Contact.Department
            officeLocation = $Contact.OfficeLocation
            fileAs         = $fileAsValue
            emailAddresses = @(
                @{
                    name    = $Contact.DisplayName
                    address = $Contact.Mail
                }
            )
        }

        # Add categories if specified
        if ($Categories.Count -gt 0) {
            $newContact.categories = $Categories
        }

        if ($Contact.EntryType -eq 'User') {
            $newContact.businessAddress = @{
                street          = $Contact.StreetAddress
                city            = $Contact.City
                state           = $Contact.State
                postalCode      = $Contact.PostalCode
                countryOrRegion = $Contact.Country
            }
            $newContact.businessPhones = $Contact.BusinessPhones
            $newContact.mobilePhone = $Contact.MobilePhone
        }
        elseif ($Contact.EntryType -eq 'Contact') {
            $newContact.businessAddress = @{
                street          = $Contact.Addresses[0].Street
                city            = $Contact.Addresses[0].City
                state           = $Contact.Addresses[0].State
                postalCode      = $Contact.Addresses[0].PostalCode
                countryOrRegion = $Contact.Addresses[0].Country
            }
            $newContact.businessPhones = @($Contact.Phones[2].Number)
            $newContact.mobilePhone = $Contact.Phones[1].Number
        }

        $newContactObject = New-MgUserContactFolderContact -UserId $Mailbox -ContactFolderId $ManagedContactFolder.Id -BodyParameter $newContact

        if ($null -ne $ContactPhotoFile) {
            Write-VerboseLog "Adding photo to contact: $($Contact.DisplayName)"
            Set-MgUserContactFolderContactPhotoContent -UserId $Mailbox -ContactFolderId $ManagedContactFolder.Id -ContactId $newContactObject.Id -InFile $ContactPhotoFile
        }
    }
}

New-Logger `
| Set-MinimumLevel -Value Verbose `
| Add-SinkConsole `
| Add-SinkFile -Path "$LogPath\GraphContactSync-.log" -RollingInterval Day -RestrictedToMinimumLevel Debug `
| Start-Logger

Write-InfoLog "Starting Graph Contact Sync"

$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

# Ensure Photos directory exists
if (!(Test-Path -Path "Photos")) {
    New-Item -ItemType Directory -Path "Photos" -Force | Out-Null
    Write-VerboseLog "Created Photos directory"
}

# Load certificate based on authentication method
# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($CertificateThumbprint) {
    # Method 1: Use certificate from Windows Certificate Store by thumbprint
    Write-InfoLog "Loading certificate from certificate store using thumbprint: $CertificateThumbprint"
    
    # Try CurrentUser store first, then LocalMachine
    $Certificate = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
    
    if (-not $Certificate) {
        $Certificate = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
    }
    
    if (-not $Certificate) {
        throw "Certificate with thumbprint '$CertificateThumbprint' not found in CurrentUser\My or LocalMachine\My certificate stores."
    }
    
    Write-InfoLog "Successfully loaded certificate: $($Certificate.Subject)"
}
else {
    # Method 2: Use PFX file with password (legacy method)
    Write-InfoLog "Loading certificate from PFX file: $CertificatePath"
    
    if ($CertificatePasswordFile) {
        # Use encrypted password file
        Write-InfoLog "Reading encrypted password from file: $CertificatePasswordFile"
        try {
            $SecureCertificatePassword = Get-Content -Path $CertificatePasswordFile | ConvertTo-SecureString
        }
        catch {
            throw "Failed to read or decrypt password file '$CertificatePasswordFile'. Ensure the file was created on this machine by the same user."
        }
    }
    else {
        # Use plaintext password (not recommended)
        Write-InfoLog "WARNING: Using plaintext password for certificate. Consider using CertificateThumbprint or CertificatePasswordFile for better security."
        $SecureCertificatePassword = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
    }
    
    # Create certificate object with secure password
    $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath, $SecureCertificatePassword)
}

# Connect to the Microsoft Graph API
Connect-MgGraph -Certificate $Certificate -ClientId $ClientID -TenantId $ExchangeOrg -Verbose

# Get all Users in the directory, selecting only key fields
$UserList = Get-MgUser -Filter '(AccountEnabled eq true)' -All -Property `
    <#                                                     #>    Id, UserType, UserPrincipalName, ShowInAddressList, EmployeeId, DisplayName, GivenName, Surname, CompanyName, JobTitle, Department, OfficeLocation, Mail, BusinessPhones, MobilePhone, StreetAddress, City, State, PostalCode, Country
| Select-Object @{Name = 'EntryType'; Expression = { 'User' } }, Id, UserType, UserPrincipalName, ShowInAddressList, EmployeeId, DisplayName, GivenName, Surname, CompanyName, JobTitle, Department, OfficeLocation, Mail, BusinessPhones, MobilePhone, StreetAddress, City, State, PostalCode, Country

# Filter out users that are not members, have no job title, or are not in the address list
#$FilteredOutUsers = $UserList | Where-Object UserType -NE 'Member' | Where-Object JobTitle -EQ $null | Where-Object ShowInAddressList -In ($false, $null)
#$UserList = $UserList | Where-Object UserType -EQ 'Member' | Where-Object ShowInAddressList -NE $false | Where-Object JobTitle -NE $null
$UserList = $UserList | Where-Object UserType -EQ 'Member' | Where-Object ShowInAddressList -NE $false | Where-Object JobTitle -NE $null | Where-Object EmployeeId -NE $null


$OrgContactList = Get-MgContact -All -Property `
    <#                                                        #>    Id, DisplayName, GivenName, Surname, CompanyName, JobTitle , Mail, Phones, Addresses
| Select-Object @{Name = 'EntryType'; Expression = { 'Contact' } }, Id, DisplayName, GivenName, Surname, CompanyName, JobTitle , Mail, Phones, Addresses

$CombinedContactList = $OrgContactList + $UserList 

if ($MailboxList -eq "DIRECTORY" ) {
    $MailboxTargets = ($UserList | Select-Object UserPrincipalName).UserPrincipalName
}
else {
    $MailboxTargets = $MailboxList -Split ","
}

foreach ($MailboxTarget in $MailboxTargets) {
    try {
        Write-DebugLog "Syncing Managed Contacts for Mailbox: $MailboxTarget"
        Sync-ManagedContacts -Mailbox $MailboxTarget -ManagedContactFolderName $ManagedContactFolderName -ManagedContacts $CombinedContactList -FileAsFormat $FileAsFormat -Categories $Categories
    }
    catch {
        Write-ErrorLog "Error syncing Managed Contacts for Mailbox: $MailboxTarget Exception: $_.Exception"
    }
}

Disconnect-MgGraph

Close-Logger
