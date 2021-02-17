<#
    .SYNOPSIS
    Set-AutoResponderNotification

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
    ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
    WITH THE USER.

    Version 1.01, February 17th, 2021

    .DESCRIPTION
    This script will set an AutoResponder inbox rule on an Exchange mailbox. This can be used to inform relations and partners sending messages
    to old e-mail addresses after mergers, acquisitions or rebranding.

    The autoresponder is configured on the new/migrated mailbox, and will look for messages specifically sent to the previous e-mail address.
    The autoresponder message can consist of subject and body and an embedded image if required to customize it further.

    Usage of the Verbose, Confirm and WhatIf parameters is supported.

    .LINK
    http://eightwone.com

    .NOTES
    Requires:
    - PowerShell 3.0 or up
    - Package Exchange.WebServices.Managed.Api (or Microsoft.Exchange.WebServices.dll in current folder)
    - Package Microsoft.Identity.Client (or Microsoft.Identity.Client.dll in current folder)
    
    For information on installing packages, see https://eightwone.com/2020/10/05/ews-webservices-managed-api)

    Special thanks to Siegfried Jagott from Intellity for the idea and testing.
    

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial public release
    1.01    Removed CSVFile parameter
            Added Begin/Process/End block for pipeline processing
            Fixed bug in module loading

    .PARAMETER Identity
    Specifies one or more e-mail addresses of mailboxes to process. Identity can also be passed through the pipeline (see examples).

    .PARAMETER OldMail
    Specifies one or more old e-mail addresses to use when configuring the autoresponder message. When specifying multiple entries,
    the number of OldMail entries need to match the number of Identity entries. Like Identity, OldMail is also passable through the pipeline.

    .PARAMETER Server
    Exchange Web Services endpoint to use. When ommited, script will attempt to use Autodiscover. By specifying Server, you
    can bypass endpoint discovery by specifying outlook.office365.com for example when using Exchange Online only, speeding
    up the process.
    
    .PARAMETER Template
    Specifies the template file to use when configuring or clearing the autoresponder inbox rule. The file should be an XML
    file in the following format:

    <?xml version="1.0" encoding="ISO-8859-1"?>  
    <config>
      <rule>Contoso Autoresponder</rule>
      <subject>Please update your email recipient to Contoso</subject>
      <body>Dear Sender,
      Thank you for your message. Fabrikam is now a whole Contoso subsidiary, and the fabrikam.com e-mail address will change to contoso.com. Your e-mail was forwarded to my new e-mail address. 
      Please update contact information, distribution lists, etc. to update [OldMail] e-mail references with my new [Identity] e-mail address. 
  
      [logo] 
      The Contoso Corporation is a multinational business with its headquarters in Paris.</body>
      <logo>Contoso.png</logo>
    </config>

    Template file elements:
    - <rule> contains the name of the inbox rule to configure or clear.
    - <subject> is the subject used to respond to messages.
    - <body> is the body of the response. Note that you can use the following placeholders, which will be subsituted with actual value:
      - [logo] is the place where you want to put a logo. The logo file is specified in the logo element of the template file.
      - [Identity] is the e-mail address of the mailbox where the rule is set, i.e. the new e-mail address
      - [OldMail] is the old e-mail address. This will also be used to only respond to messages sent to this address.
    - <logo> is the filename of the logo file to embed in the response (Optional).

    .PARAMETER TenantId
    Specifies the identity of the Tenant.

    .PARAMETER ClientId
    Specifies the identity of the application configured in Azure Active Directory.

    .PARAMETER Credentials
    Specify credentials to use with Basic Authentication. Credentials can be set using $Credentials= Get-Credential
    This parameter is mutually exclusive with CertificateFile, CertificateThumbprint and Secret. 

    .PARAMETER CertificateThumbprint
    Specify the thumbprint of the certificate to use with OAuth authentication. The certificate needs
    to reside in the personal store. When using OAuth, providing TenantId and ClientId is mandatory.
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificateFile
    Specify the .pfx file containing the certificate to use with OAuth authentication. When a password is required,
    you will be prompted or you can provide it using CertificatePassword.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificatePassword
    Sets the password to use with the specified .pfx file. The provided password needs to be a secure string, 
    eg. -CertificatePassword (ConvertToSecureString -String 'P@ssword' -Force -AsPlainText)

    .PARAMETER Secret
    Specifies the client secret to use with OAuth authentication. The secret needs to be provided as a secure string.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and CertificateThumbprint. 

    .PARAMETER Impersonation
    When specified, uses impersonation when accessing the mailbox, otherwise account specified with Credentials is
    used. When using OAuth authentication with a registered app, you don't need to specify Impersonation.
    For details on how to configure impersonation access for Exchange 2010 using RBAC, see this article:
    https://eightwone.com/2014/08/13/application-impersonation-to-be-or-pretend-to-be/

    .PARAMETER Clear
    Specifies if any you want to remove the inbox rules with name specified in the template. Use this when you want
    to remove autoresponder rules from mailboxes. When using Clear, you don't need to specify OldMail.

    .PARAMETER Overwrite
    Specifies if any existing inbox rules with name specified in the template should be overwritten. When omitted, the script will skip processing
    a mailbox if an existing rule is found. Use this when you want to configure the autoresponder only on mailboxes which do not have the rule.

    .PARAMETER TrustAll
    Specifies if all certificates should be accepted, including self-signed certificates.

    .EXAMPLE
    $cred=Get-Credential
    .\Set-AutoResponderNotification.ps1 -Identity michel@contoso.com -OldMail michel@fabrikam.com -Server outlook.office365.com -TemplateFile .\Template.xml -Credential $cred

    Configure autoresponder rule on mailbox of michel@contoso.com, using previous e-mail address michel@fabrikam as inbox rule SendTo predicate. 
    The autoresponder will be set using parameters stored in the template file, and credentials are used to perform basic authentication against endpoint.

    .EXAMPLE
    $Secret= Read-Host 'Secret' -AsSecureString
    .\Set-AutoResponderNotification.ps1 -Identity michel@contoso.com -OldMail michel@fabrikam.com -Server outlook.office365.com -TemplateFile .\Template.xml -TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -Secret $Secret

    Configure autoresponder rule on mailbox of michel@contoso.com, using previous e-mail address michel@fabrikam as inbox rule SendTo predicate. 
    The tenant indicated by specified identities is used to authenticate against, using specified application identity and provided secret. 
    The autoresponder will be set using parameters stored in the template file.

    .EXAMPLE
    $PfxPwd= Read-Host 'PFX password' -AsSecureString
    .\Set-AutoResponderNotification.ps1 -Identity michel@contoso.com -Server outlook.office365.com -TemplateFile .\Template.xml -TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -CertificateFile .\AutoResponder.pfx -CertificatePassword $PfxPwd -Clear

    Clear autoresponder rule on mailbox of michel@contoso.com. The tenant indicated by specified identities is used to authenticate against, 
    using specified application identity and provided pfx certificate file and pfx decryption password. The subject to look for will be taken 
    from the specified template file.

    .EXAMPLE
    Import-CSV -Path Users.csv | .\Set-AutoResponderNotification.ps1 -Server outlook.office365.com -Impersonation -TemplateFile .\Template.xml -TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -Overwrite -CertificateFile .\AutoResponder.pfx -CertificatePassword (ConvertTo-SecureString 'P@ssw0rd' -Force -AsPlainText)

    Configure autoresponder rule using Identity/OldMail properties from the CSV file. The tenant specified is authenticated against, 
    using specified application identity, as well as the certificate from the personal certificate store with the specified thumbprint. 
    The autoresponder will be set using parameters stored in the template file. Any existing inbox rules with the same name will be overwritten.

#>

[cmdletbinding( SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemBasic')]
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'ClearBasic')]
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'ClearOAuthCertThumb')]
    [string[]]$Identity,
    [parameter( Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemBasic')]
    [parameter( Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [string[]]$OldMail,
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemBasic')]
    [parameter( Mandatory= $false, ParameterSetName= 'ClearBasic')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthCertThumb')]
    [string]$Server,
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemBasic')]
    [parameter( Mandatory= $false, ParameterSetName= 'ClearBasic')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthCertThumb')]
    [switch]$Impersonation,
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemBasic')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearBasic')]
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthSecret')]
    [System.Security.SecureString]$Secret,
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertThumb')]
    [String]$CertificateThumbprint,
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertFile')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [String]$CertificateFile,
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthCertFile')]
    [System.Security.SecureString]$CertificatePassword,
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertThumb')]
    [string]$TenantId,
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertThumb')]
    [string]$ClientId,
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemBasic')]
    [parameter( Mandatory= $true, ParameterSetName= 'ClearBasic')]
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Mandatory= $true, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertThumb')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string]$TemplateFile,
    [parameter( Mandatory= $true, ParameterSetName= 'ClearBasic')]
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Mandatory= $true, ParameterSetName= 'ClearOAuthCertThumb')]
    [switch]$Clear,
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemBasic')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [switch]$Overwrite,
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemBasic')]
    [parameter( Mandatory= $false, ParameterSetName= 'ClearBasic')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthSecret')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthCertFile')]
    [parameter( Mandatory= $false, ParameterSetName= 'SingleItemOAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'ClearOAuthCertThumb')]
    [switch]$TrustAll
)
#Requires -Version 3.0

begin {

    # Errors
    $ERR_DLLNOTFOUND                         = 1000
    $ERR_DLLLOADING                          = 1001
    $ERR_AUTODISCOVERFAILED                  = 1003
    $ERR_CANTACCESSMAILBOXSTORE              = 1004
    $ERR_TEMPLATECONFIG                      = 1008
    
    Function Import-ModuleDLL {
        param(
            [string]$Name,
            [string]$FileName,
            [string]$Package,
            [string]$ValidateObjName
        )

        $AbsoluteFileName= Join-Path -Path $PSScriptRoot -ChildPath $FileName
        If ( Test-Path $AbsoluteFileName) {
            # OK
        }
        Else {
           If( $Package) {
               If( Get-Command -Name Get-Package -ErrorAction SilentlyContinue) {
                    If( Get-Package -Name -ErrorAction SilentlyContinue) {
                        $AbsoluteFileName= (Get-ChildItem -ErrorAction SilentlyContinue -Path (Split-Path -Parent (get-Package -Name $Package | -Object -Property Version -Descending | Select-Object -First 1).Source) -Filter $FileName -Recurse).FullName
                    }
                }
            }
        }

        If( $absoluteFileName) {
            $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
            If( $ModLoaded) {
                Write-Verbose ('Module {0} v{1} already loaded' -f $ModLoaded.Name, $ModLoaded.Version)
            }
            Else {
                Write-Verbose ('Loading module {0}' -f $absoluteFileName)
                try {
                    Import-Module -Name $absoluteFileName -Global -Force
                }
                catch {
                    Write-Error ('Problem loading module {0}: {1}' -f $Name, $error[0])
                    Exit $ERR_DLLLOADING
                }
                $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
                If( $ModLoaded) {
                    Write-Verbose ('Module {0} v{1} loaded' -f $ModLoaded.Name, $ModLoaded.Version)
                }
                Try {
                    If( $validateObjName) {
                        $null= New-Object -TypeName $validateObjName
                    }
                }
                Catch {
                    Write-Error ('Problem initializing test-object from module {0}: {1}' -f $Name, $error[0])
                    Exit $ERR_DLLLOADING
                }
            }
       }
       Else {
           Write-Verbose ('Required module {0} could not be located' -f $FileName)
           Exit $ERR_DLLNOTFOUND
       }
    }

    Function Set-SSLVerification {
        param(
            [switch]$Enable,
            [switch]$Disable
        )

        Add-Type -TypeDefinition  @"
            using System.Net.Security;
            using System.Security.Cryptography.X509Certificates;
            public static class TrustEverything
            {
                private static bool ValidationCallback(object sender, X509Certificate certificate, X509Chain chain,
                    SslPolicyErrors sslPolicyErrors) { return true; }
                public static void SetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback = ValidationCallback; }
                public static void UnsetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback = null; }
        }
"@
        If($Enable) {
            Write-Verbose ('Enabling SSL certificate verification')
            [TrustEverything]::UnsetCallback()
        }
        Else {
            Write-Verbose ('Disabling SSL certificate verification')
            [TrustEverything]::SetCallback()
        }
    }

    Import-ModuleDLL -Name 'Microsoft.Exchange.WebServices' -FileName 'Microsoft.Exchange.WebServices.dll' -Package 'Exchange.WebServices.Managed.Api' -validateObjName 'Microsoft.Exchange.WebServices.Data.ExchangeVersion'
    Import-ModuleDLL -Name 'Microsoft.Identity.Client' -FileName 'Microsoft.Identity.Client.dll' -Package 'Microsoft.Identity.Client' -validateObjName 'Microsoft.Identity.Client.ConfidentialClientApplicationBuilder'

    Try  {
        $Config= ([xml](Get-Content -Path $TemplateFile)).Config
    }
    Catch {
        Write-Error ('Provided template file malformed or improper xml format.')
        Exit $ERR_TEMPLATECONFIG
    }

    If(-not $Clear) {

        If( $Config.logo) {
            If( Test-Path -Path $Config.logo) {
                Write-Verbose ('External logo file {0} found.' -f $Config.logo)
            }
            Else {
                Write-Error ( 'Specified logo file {0} not found.' -f $Config.logo)
                Exit $ERR_TEMPLATECONFIG
            }
        }
        Else {
            # No logo
        }
    }
    If( $Config.Rule) {
        Write-Verbose ('Using Rule name: {0}' -f $Config.Rule)
    }
    Else {
        Write-Error ( 'Required Rule element not found in template file.')
        Exit $ERR_TEMPLATECONFIG
    }

    $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    $EwsService= [Microsoft.Exchange.WebServices.Data.ExchangeService]::new( $ExchangeVersion)

    If( $Credentials) {
        try {
            Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
            $EwsService.Credentials= [System.Net.NetworkCredential]::new( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
        }
        catch {
            Write-Error ('Invalid credentials provided: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }
    Else {
        # Use OAuth (and impersonation/X-AnchorMailbox always set)
        $Impersonation= $true

        If( $CertificateThumbprint -or $CertificateFile) {
            If( $CertificateFile) {
                
                # Use certificate from file using absolute path to authenticate
                $CertificateFile= (Resolve-Path -Path $CertificateFile).Path
                
                If( $CertificatePassword) {
                    $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $CertificatePassword)))
                }
                Else {
                    $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile)
                }
                If(!( $X509Certificate2)) {
                    Throw 'Problem importing PFX'
                }
            }
            Else {
                # Use provided certificateThumbprint to retrieve certificate from My store, and authenticate with that
                $CertStore= [System.Security.Cryptography.X509Certificates.X509Store]::new( [Security.Cryptography.X509Certificates.StoreName]::My, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
                $CertStore.Open( [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly )
                $X509Certificate2= $CertStore.Certificates.Find( [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $CertificateThumbprint, $False) | Select-Object -First 1
                If(!( $X509Certificate2)) {
                    Throw 'Problem locating certificate in My store'
                }
            }
            Write-Verbose ('Will use certificate {0}, issued by {1} and expiring {2}' -f $X509Certificate2.Thumbprint, $X509Certificate2.Issuer, $X509Certificate2.NotAfter)
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithCertificate( $X509Certificate2).withTenantId( $TenantId).Build()
               
        }
        Else {
            # Use provided secret to authenticate
            $PlainSecret= [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Secret))
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithClientSecret( $PlainSecret).withTenantId( $TenantId).Build()
        }
        $Scopes = New-Object System.Collections.Generic.List[string]
        $Scopes.Add( 'https://outlook.office365.com/.default')
        Try {
            $Response=$App.AcquireTokenForClient( $Scopes).executeAsync()
            $Token= $Response.Result
            $EwsService.Credentials= [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$Token.AccessToken
        }
        Catch {
            Write-Error ('Problem acquiring token: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }

    If( $TrustAll) {
        Set-SSLVerification -Disable
    }

}

Process {

    $Entries= @{}
    $i=0
    ForEach($Entry in $Identity) {
        Try {
            $Entries[ $Entry]= $OldMail[ $i]
        }
        Catch {
            $Entries[ $Entry]= $Entry
        }
        $i++
    }

    ForEach( $Item in $Entries.getEnumerator()) {

        $ID= $Item.Name
        $OldID= $Item.Value

        Write-Host ('Processing mailbox {0}, old e-mail identity {1}' -f $ID, $OldID)

        If( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $ID)
            $EwsService.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new( [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $ID)
            $EwsService.HttpHeaders.Clear()
            $EwsService.HttpHeaders.Add( 'X-AnchorMailbox', $ID)
        }
            
        If ($Server) {
            $EwsUrl= 'https://{0}/EWS/Exchange.asmx' -f $Server
            Write-Verbose ('Using Exchange Web Services URL {0}' -f $EwsUrl)
            $EwsService.Url= $EwsUrl
        }
        Else {
            Write-Verbose ('Looking up EWS URL using Autodiscover for {0}' -f $EmailAddress)
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference= 'Stop'
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error ('Autodiscover failed: {0}' -f $error[0])
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference= 'Continue'
            Write-Verbose 'Using EWS endpoint {0}' -f $EwsService.Url
        } 

        #This is where magic starts..
        try {
            $null= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::MsgFolderRoot)
        }
        catch {
            Write-Error ('Cannot access mailbox information store for {0}: {1}' -f $ID, $_.Exception.Message)
            Exit $ERR_CANTACCESSMAILBOXSTORE
        }

        # See if any matching rules already configured
        $InboxRules= $EwsService.getInboxRules( $ID)
        $ExistingRuleIds= [System.Collections.ArrayList]@()
        ForEach( $InboxRule in $InboxRules) {
            If( $InboxRule.displayName -eq $Config.Rule) {
                $ExistingRuleIds.Add( $InboxRule.Id) | Out-Null
            }
        }

        # In Clear mode, remove existing matching rules only
        If( $Clear -or $Overwrite) {
            If( $ExistingRuleIds.Count -gt 0) {
                $deleRule= New-Object Microsoft.Exchange.WebServices.Data.DeleteRuleOperation[] $ExistingRuleIds.Count
                $i=0
                ForEach( $RuleId in $ExistingRuleIds) { 
                    $deleRule[ $i++]= [Microsoft.Exchange.WebServices.Data.DeleteRuleOperation]::new( $RuleId) 
                }
                If ( $Force -or $PSCmdlet.ShouldProcess( ('Remove existing inbox rule(s) "{1}" from mailbox {0}' -f $ID, $Config.Rule))) {
                    $EwsService.updateInboxRules( $deleRule, $true)
                }
            }
            Else {
                Write-Host ('No existing inbox rule "{1}" found on mailbox {0}' -f $ID, $Config.Rule)
            }
        }

        If( -not $Clear) {

            # NotClear, process mailbox when no existing rules found
            If( (-not $Overwrite) -and ($RuleIdsToDelete.Count -gt 0)) {

                Write-Host ('Skipping mailbox {0} with existing rule "{1}"' -f $ID, $Config.Rule)
            }
            Else {

                # Construct template mail to use and store in inbox
                $TemplateEMail= [Microsoft.Exchange.WebServices.Data.EmailMessage]::new( $EwsService)
                $TemplateEmail.ItemClass= 'IPM.Note.Rules.ReplyTemplate.Microsoft'
                $TemplateEmail.IsAssociated= $true
                $TemplateEmail.Subject= $Config.Subject

                # Replace any mention of old address in template
                $Body= $Config.Body -ireplace '\[Identity\]', $ID

                #CSV Mode, replace mention of new address
                $Body= $Body -ireplace '\[OldMail\]', $OldID

                # Replace LF with <BR>+LF for HTML messages
                $Body= $Body -replace "`n","<br />`n"

                # When specified, attach & embed logo
                If( $Config.logo) {
                    $logoFileName= Split-Path -Path $Config.logo -Leaf
                    $logoFullName= (Resolve-Path -Path $Config.logo).Path 
                    $Body= $Body -replace '\[logo\]', ('<img id="1" src="cid:{0}">' -f $logoFileName)
                    $TemplateEmail.Attachments.AddFileAttachment( $logoFullName) | Out-Null
                    $TemplateEmail.Attachments[0].isInline= $true
                    $TemplateEmail.Attachments[0].contentId= $logoFileName
                }

                $TemplateEmail.Body= [Microsoft.Exchange.WebServices.Data.MessageBody]::new( $Body)

                $pidTagReplyTemplateId= [Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition]::new( 0x65C2, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
                $TemplateEmail.setExtendedProperty( $pidTagReplyTemplateId, [System.Guid]::NewGuid().ToByteArray())
                $TemplateEmail.save( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox) 
                $InboxRule= [Microsoft.Exchange.WebServices.Data.Rule]::new()
                $InboxRule.DisplayName= $Config.Rule
                $InboxRule.Conditions.SentToAddresses.Add( $OldMail) | Out-Null
                $InboxRule.Actions.ServerReplyWithMessage= $TemplateEmail.Id
                $InboxRule.Actions.StopProcessingRules= $true
                $InboxRule.Exceptions.ContainsSubjectStrings.Add( $Config.Subject)
                $creaRule= New-Object Microsoft.Exchange.WebServices.Data.CreateRuleOperation[] 1
                $creaRule[0]= $InboxRule

                If ( $Force -or $PSCmdlet.ShouldProcess( ('Configure inbox rule "{1}" on mailbox {0}' -f $ID, $Config.Rule))) {
                    $EwsService.updateInboxRules( $creaRule, $true)
                }
            }
        }
    }
}

End {
    If( $TrustAll) {
        Set-SSLVerification -Enable
    }
}
