# Set-AutoResponderNotification

## Getting Started

### Requirements

* PowerShell 3.0 or later
* Microsoft.Exchange.WebServices (EWS Managed API)
* Microsoft.Identity.Client (MSAL)

### Usage

```powershell
$cred=Get-Credential
.\Set-AutoResponderNotification.ps1 -Identity michel@contoso.com -OldMail michel@fabrikam.com -Server outlook.office365.com -TemplateFile .\Template.xml -Credential $cred
```
Configure autoresponder rule on mailbox of michel@contoso.com, using previous e-mail address michel@fabrikam as inbox rule SendTo predicate. The autoresponder will be set using parameters stored in the template file, and Basic Authentication is used to authenticate.

```powershell
$Secret= Read-Host 'Secret' -AsSecureString
.\Set-AutoResponderNotification.ps1 -Identity michel@contoso.com -OldMail michel@fabrikam.com -Server outlook.office365.com -TemplateFile .\Template.xml -TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -Secret $Secret
```
Configure autoresponder rule on mailbox of michel@contoso.com, using previous e-mail address michel@fabrikam as inbox rule SendTo predicate. The tenant indicated by specified identities is used to authenticate against, using specified application identity and provided secret. The autoresponder will be set using parameters stored in the template file.

```powershell
$PfxPwd= Read-Host 'PFX password' -AsSecureString
.\Set-AutoResponderNotification.ps1 -Identity michel@contoso.com -Server outlook.office365.com -TemplateFile .\Template.xml -TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -CertificateFile .\AutoResponder.pfx -CertificatePassword $PfxPwd -Clear
```
Clear autoresponder rule on mailbox of michel@contoso.com. The tenant indicated by specified identities is used to authenticate against, using specified application identity and provided pfx certificate file and pfx decryption password. The subject to look for will be taken from the specified template file.

```powershell
.\Set-AutoResponderNotification.ps1 -Identity michel@contoso.com -OldMail michel@fabrikam.com -Server outlook.office365.com -Impersonation -TemplateFile .\Template.xml -TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -Overwrite -CertificateFile .\AutoResponder.pfx -CertificatePassword (ConvertTo-SecureString 'P@ssw0rd' -Force -AsPlainText)
```
Configure autoresponder rule on mailbox of michel@contoso.com using old e-mail address michel@fabrikam as predicate. The tenant indicated by specified identities is used to authenticate against, using specified application identity. The certificate used to authenticate is picked from the personal certificate store by looking for specified thumbprint. The autoresponder will be set using parameters stored in the template file. Any existing inbox rules with the same name will be overwritten.

### About


## License

This project is licensed under the MIT License - see the LICENSE.md for details.

 