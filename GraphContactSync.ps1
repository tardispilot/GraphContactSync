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
    The path to the certificate file.
.PARAMETER CertificateAuthPath
    The path to the certificate password file.
.PARAMETER MailboxList
    The list of mailboxes to sync contacts to.
.PARAMETER ManagedContactFolderName
    The name of the folder to sync contacts to.
.PARAMETER LogPath
    The path to the log file.
#>

Param(    
    [Parameter(Mandatory = $true)]
    [string]$ExchangeOrg,
    
    [Parameter(Mandatory = $true)]
    [string]$ClientID,
    
    [Parameter(Mandatory = $true)]
    [System.IO.FileInfo]$CertificatePath,
    
    [Parameter(Mandatory = $true)]
    [string]$CertificatePassword,
    
    [Parameter(Mandatory = $true)]
    [System.IO.FileInfo]$CertificatePasswordFile,

    [Parameter(Mandatory = $true)]
    [string]$MailboxList,
    
    [Parameter(Mandatory = $true)]
    [string]$ManagedContactFolderName,

    [Parameter(Mandatory = $true)]
    [string]$LogPath
)

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
        $ManagedContacts
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

        Write-VerboseLog "Comparing Managed Contact Checksums: [$ExistingContactChecksum] vs [$ManagedContactChecksum]"

        if ($ExistingContactChecksum -ne $ManagedContactChecksum) {
            #if this is an edited contact, effectively delete the old one and add the new one
            Write-DebugLog "Detected changed contact: $($ManagedContact.DisplayName)"
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
        if ($Contact.EntryType -eq 'User') {
            # Get the photo for the contact
            if (!(Test-Path -PathType Leaf -Path "Photos\$($Contact.UserPrincipalName).jpg")) {
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

        $newContact = @{
            extensions     = @(
                @{
                    "@odata.type" = "microsoft.graph.openTypeExtension"
                    ExtensionName = "ManagedContactCorrelation"
                    CorrelationId = $Contact.Id.ToString()
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
            emailAddresses = @(
                @{
                    name    = $Contact.DisplayName
                    address = $Contact.Mail
                }
            )
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

# Import the Application Certificate Password
#[Security.SecureString]$CertificatePassword = Get-Content -Path $CertificatePasswordFile | ConvertTo-SecureString -Key (Get-Content -Path $CertificatePasswordKeyFile)

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Create a new X509 Certificate object with the PFX file and password
$Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath, $CertificatePassword)

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
        Sync-ManagedContacts -Mailbox $MailboxTarget -ManagedContactFolderName $ManagedContactFolderName -ManagedContacts $CombinedContactList
    }
    catch {
        Write-ErrorLog "Error syncing Managed Contacts for Mailbox: $MailboxTarget Exception: $_.Exception"
    }
}

Disconnect-MgGraph

Close-Logger
