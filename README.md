# ERA-Report Exchange Online

This script automatically connects to Exchange Online, counts all mailboxes in Exchange Online and sends a report by e-mail via the Microsoft Graph API for reporting purposes. The actual purpose is to count the mailboxes for the billing of Fortinet ERA licensing. 

## Requirements

- Azure AD admin access
- Microsoft 365 account with sufficient rights
- A (self-signed) certificate for connecting to Exchange Online.

## Usage
### Befor starting

1. App registration in the Azure portal
 - Go to https://portal.azure.com
 - "Azure Active Directory" → "App registrations" → New registration

2. Upload (self-signed) certificate and create secret
  - Run the Script under certs -> New-SelfSignedCertificate.ps1 to create self-signed certificate.
    - The script will create a self-signed certificate in your machine local certificate store an an .crt file in the same directory as the script.
  - Upload certificate
  - Create new secret (client secret) -> Copy the value & save

3. Add API authorizations
  - Links "API permissions" -> Add permission:
    - Microsoft Graph → Application:
      - Mail.Send
      - After adding: Grant admin approval
   - Add permission:
     - Exchange.ManageAsApp
     - After adding: Grant admin approval

4. Modify config.json
  - Open the JSON file config -> config.json and modify following:
    - AppID: (MS365 App ID)
    - ClientSecret: (The secret you created in step 2. Attention: Not the ID from the secret, the value!)
    - CertificateThumbprint: (Thumbprint of the certificate you uploaded)
    - Organization: (Mosty like ORGANIZATION.onmicrosoft.com)
    - FromEmail: (Which sender mail should be used? The e-mail address must exist in MS365. A shared mailbox is sufficient)
    - ToEmail: (to whom should the mail be sent?)
    - CCEmail: (If no CC is required, the value must still not be empty. The same address as “ToEmail” should be entered)

  
