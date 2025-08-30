# GraphContactSync - Permission Verification Results

## Summary
The Azure app registration permissions have been verified and updated to remove obsolete Exchange Web Services (EWS) dependencies. The application now requires only Microsoft Graph permissions.

## Changes Made

### Removed Permissions (Obsolete)
**Office 365 Exchange Online** (`00000002-0000-0ff1-ce00-000000000000`):
- ❌ `dc50a0fb-09a3-484d-be87-e023b12c6440` - Exchange.ManageAsApp
- ❌ `dc890d15-9560-4a4c-9b7f-a736ec74ec40` - full_access_as_app (EWS permission)

**Reason for removal:** This version of GraphContactSync uses the Microsoft Graph PowerShell SDK exclusively and does not make any Exchange Web Services (EWS) calls.

### Required Permissions (Minimal Set)
**Microsoft Graph** (`00000003-0000-0000-c000-000000000000`):
- ✅ `6918b873-d17a-4dc1-b314-35f528134491` - Contacts.ReadWrite
- ✅ `df021288-bdef-4463-88db-98f22de89214` - User.Read.All

## Permission Justification

### Contacts.ReadWrite
**Required for:**
- `Get-MgUserContactFolder` - Read contact folders
- `New-MgUserContactFolder` - Create contact folders  
- `Get-MgUserContactFolderContact` - Read contacts from folders
- `New-MgUserContactFolderContact` - Create contacts in folders
- `Remove-MgUserContactFolderContact` - Delete contacts from folders
- `Set-MgUserContactFolderContactPhotoContent` - Set contact photos

### User.Read.All  
**Required for:**
- `Get-MgUser` - Read user profiles from directory
- `Get-MgContact` - Read organizational contacts (covered by User.Read.All)
- `Get-MgUserPhoto` - Read user photo metadata
- `Get-MgUserPhotoContent` - Download user photos

## Updated Manifest

```json
"requiredResourceAccess": [
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

## Impact
- **Security:** Reduced permission scope by removing Exchange-specific permissions
- **Maintenance:** Simplified permission model with only Graph permissions  
- **Compatibility:** No functional impact - all operations continue to work through Graph API
- **Future-proof:** Aligned with Microsoft's Graph-first strategy

## Migration Notes
Existing Azure app registrations can safely remove the Exchange Online permissions. The application will continue to function with only the Microsoft Graph permissions listed above.