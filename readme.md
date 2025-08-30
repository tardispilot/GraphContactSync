# Graph Contact Sync

Synchronizes Global Address List and Organizational Contacts from a M365 environment to selected/all mailboxes in the directory. Uses the MS Graph API PowerShell module to perform all operations.

Heavily inspired by the excellent [EWS-Office365-Contact-Sync](https://github.com/grahamr975/EWS-Office365-Contact-Sync) code by grahamr975.

Uses [PoShLog](https://github.com/PoShLog/PoShLog) for structured console/file logs.

## History

My company needed a centralized/automated method to ditribute corporate contact lists to all employees' phones/Exchange mailboxes. There is a shocking lack of good affordable tooling for what seems like a common use case. So I wrote this.

Please see the **Acknowledgements** section for attribution of the idea that started this all off.

## Features

- Includes Org-level contacts from M365 > Users > Contacts. Useful for non-person contacts such as office/branch information.
- Compares old and new field values so it only replaces a contact entry if change detected.
- Handles employee photos. Not very well, but hey it works-ish.
- **FileAs field formatting**: Configure how contacts are filed ("First Last" or "Last, First" format).
- **Categories support**: Assign categories to contacts, useful when syncing to main Contacts folder.

## Security

### Certificate Authentication Methods

This script supports three authentication methods, listed in order of security (most secure first):

1. **Certificate Thumbprint (Recommended)**: Uses certificates installed in the Windows Certificate Store. No password storage required.
2. **Encrypted Password File**: Stores the PFX password in an encrypted file that can only be decrypted by the same user on the same machine.
3. **Plaintext Password**: Stores the password in plaintext in scripts or command line. **Not recommended for production use.**

### Security Best Practices

- **Use Certificate Thumbprint authentication** whenever possible for the highest security
- **Never commit plaintext passwords** to source control
- **Store encrypted password files securely** and restrict access
- **Regularly rotate certificates** and update thumbprints
- **Use least privilege principles** when assigning Azure application permissions

### Creating Encrypted Password Files

If you have existing PFX files and need to create encrypted password files:

```powershell
.\Getting Started\Create-EncryptedPassword.ps1 -OutputPath "C:\Certs\certificate.cred"
```

## Parameters

### Required Parameters
- `ExchangeOrg`: The Exchange Organization to connect to
- `ClientID`: The Client ID for the application
- `MailboxList`: The list of mailboxes to sync contacts to (or "DIRECTORY" for all)
- `ManagedContactFolderName`: The name of the folder to sync contacts to
- `LogPath`: The path to the log file

### Certificate Authentication Parameters (Choose One Method)

#### Method 1: Certificate Thumbprint (Recommended - Most Secure)
- `CertificateThumbprint`: The thumbprint of the certificate installed in the Windows Certificate Store

#### Method 2: PFX File with Encrypted Password File  
- `CertificatePath`: The path to the certificate PFX file
- `CertificatePasswordFile`: The path to an encrypted password file

#### Method 3: PFX File with Plaintext Password (Not Recommended)
- `CertificatePath`: The path to the certificate PFX file  
- `CertificatePassword`: The certificate password (security risk - use other methods instead)

### Optional Parameters
- `FileAsFormat`: How to format the FileAs field. Valid values:
  - `"FirstLast"` (default): "John Smith"
  - `"LastFirst"`: "Smith, John"
- `Categories`: Array of categories to assign to contacts. Example: `@("Business Contacts", "Company Directory")`

## Notes and disclaimers

- Test this thoroughly before deploying in any automated/unconstrained way. This application has permissions to **DELETE** contacts from mailboxes, so be wary.
- This works for my own company, but seems only fitting to release a variant of the project that inspired its creation for others to use/modify for their own needs.
- I am fairly new to writing PowerShell, and it probably shows. There are a few known issues and improvements that I will work on as time allows. These will be documented as GitHub Issues.
- **PRs welcome!** There is no formal Code of Conduct for this (yet), other than "be nice". Unfortunately this has to be said these days, but in general if you wouldn't want to see/experience certain behavior, then don't do it yourself.

## Getting Started

1. Install the Exchange Online PowerShell module
   ```
   Install-Module ExchangeOnlineManagement
   ```
2. Install PoShLog (for console logging)
   ```
   Install-Module PoShLog
   ```
3. Create Certificate files
   - Using the script in the `Getting Started` folder
   ```
   .\Create-Certificates.ps1 -CertificateName contactsync.mydomain.com -CertificatePassword 'myPassword!' [-CertificatePath <path>] [-CreatePasswordFile]
   # Use -CreatePasswordFile to create an encrypted password file for secure storage
   # Do NOT use -RemoveCert if you want to use thumbprint authentication (recommended)
   ```
   - This will result in files being created and display the certificate thumbprint:
   ```
   contactsync.mydomain.com.pfx <-- This file contains the public and PRIVATE KEY. Take care!
   contactsync.mydomain.com.cer <-- This file contains the public key for uploading to Azure.
   contactsync.mydomain.com.cred <-- Encrypted password file (if -CreatePasswordFile used)
   Certificate Thumbprint: 1234567890ABCDEF... <-- Use this for secure authentication
   ```
4. Create an Azure app & certificate file using [the tutorial here](https://github.com/MicrosoftDocs/office-docs-powershell/blob/main/exchange/docs-conceptual/app-only-auth-powershell-v2.md), taking note of the differences below.
   - The app will require **Global Reader** permission (Referenced in tutorial).
   - Take a record of the Azure app's **Application (client) ID** as you'll need this later.
   - Enable Public Client Flows in the Azure App (**Authenication** -> **Allow public client flows**)
   - Specify a redirect URI (**Authenication** -> **Platform Configurations** -> **Add a platform** -> **Mobile and desktop applications** -> Enable 'https://login.microsoftonline.com/common/oauth2/nativeclient' as a redirect URI.)
   - When updating the app's Manifest, insert the code below for **requiredResourceAccess** instead of following what the tutorial suggests.
     ```
     	"requiredResourceAccess": [
     	{
     	"resourceAppId": "00000002-0000-0ff1-ce00-000000000000",
     		"resourceAccess": [
     			{
     				"id": "dc50a0fb-09a3-484d-be87-e023b12c6440",
     			"type": "Role"
     			},
     			{
     				"id": "dc890d15-9560-4a4c-9b7f-a736ec74ec40",
     				"type": "Role"
     		}
     		]
     	},
     {
     	"resourceAppId": "00000003-0000-0000-c000-000000000000",
     		"resourceAccess": [
     			{
     				"id": "6918b873-d17a-4dc1-b314-35f528134491",
     			"type": "Role"
     			},
     			{
     				"id": "df021288-bdef-4463-88db-98f22de89214",
     				"type": "Role"
     		}
     		]
     	}
     ]
     ```
     The application's certificate has already been generated in a previous step, so skip that section in the tutorial, uploading the .cer file .
5. Confirm permissions are correct from the **API permissions** page.
   ![Correct API permissions example](images/api_permissions.png)
6.
7. You'll also need your Office 365 organization URL (Ends in .onmicrosoft.com). To find this, navigate to the **Office 365 Admin Center** -> **Setup** -> **Domains**
8. To test the script, choose one of the secure authentication methods below:

   **Method 1 - Certificate Thumbprint (Recommended):**
   ```powershell
   #_RunSingle.ps1
   .\GraphContactSync.ps1 `
   	-ExchangeOrg "mytenantname.onmicrosoft.com" `
   	-ClientID "506dcb63-64c6-4b8b-9a5a-f5cdabb123e9" `
   	-CertificateThumbprint "1234567890ABCDEF..." `
   	-MailboxList "justme@mycompany.com" `
   	-ManagedContactFolderName "My Company - Managed" `
   	-LogPath "$PSScriptRoot\Logs" `
   	-FileAsFormat "LastFirst" `
   	-Categories @("Business Contacts", "Company Directory")
   ```

   **Method 2 - Encrypted Password File:**
   ```powershell
   #_RunSingle.ps1
   .\GraphContactSync.ps1 `
   	-ExchangeOrg "mytenantname.onmicrosoft.com" `
   	-ClientID "506dcb63-64c6-4b8b-9a5a-f5cdabb123e9" `
   	-CertificatePath "C:\CERTS\contactsync.mydomain.com.pfx" `
   	-CertificatePasswordFile "C:\CERTS\contactsync.mydomain.com.cred" `
   	-MailboxList "justme@mycompany.com" `
   	-ManagedContactFolderName "My Company - Managed" `
   	-LogPath "$PSScriptRoot\Logs" `
   	-FileAsFormat "LastFirst" `
   	-Categories @("Business Contacts", "Company Directory")
   ```

   **Method 3 - Plaintext Password (Not Recommended):**
   ```powershell
   #_RunSingle.ps1
   .\GraphContactSync.ps1 `
   	-ExchangeOrg "mytenantname.onmicrosoft.com" `
   	-ClientID "506dcb63-64c6-4b8b-9a5a-f5cdabb123e9" `
   	-CertificatePath "C:\CERTS\contactsync.mydomain.com.pfx" `
   	-CertificatePassword "ThereHasToBeABetterWay?!" `
   	-MailboxList "justme@mycompany.com" `
   	-ManagedContactFolderName "My Company - Managed" `
   	-LogPath "$PSScriptRoot\Logs" `
   	-FileAsFormat "LastFirst" `
   	-Categories @("Business Contacts", "Company Directory")
   ```
9. Once you're ready, specify "DIRECTORY" for the MailboxList parameter
   ```
    -MailboxList "DIRECTORY" `
   ```
10. Once you are comfortable that the scipt is working, adding a Task in Task Scheduler on an always-on system is the simplest way to set this and "forget it", until the certificate needs renewing.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgements

- Thanks to **Ryan Graham** for the [EWS-Office365-Contact-Sync](https://github.com/grahamr975/EWS-Office365-Contact-Sync) code, but moreso the concept of this process.
