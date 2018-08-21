<#

.SYNOPSIS

Created by: https://ingogegenwarth.wordpress.com/
Version:    2.0.0.0 42 ("What do you get if you multiply six by nine?")
Changed:    15.06.2018

.LINK
http://www.get-blog.com/?p=189
http://learn-powershell.net/2012/05/13/using-background-runspaces-instead-of-psjobs-for-better-performance/
http://gsexdev.blogspot.de/
https://ingogegenwarth.wordpress.com/2018/06/15/exo-mig-invalidperm/
https://ingogegenwarth.wordpress.com/2015/04/01/the-good-the-bad-and-sidhistory/
https://ingogegenwarth.wordpress.com/2015/04/16/get-mailbox-folder-permissions-using-ews-multithreading/

.DESCRIPTION

The purpose of the script is to get for each folder granted permissions for given mailboxes

.PARAMETER EmailAddress

The e-mail address of the mailbox, which will be checked. The script accepts piped objects from Get-Mailbox or Get-Recipient

.PARAMETER Credentials

Credentials you want to use. If omitted current user context will be used.

.PARAMETER Impersonate

Use this switch, when you want to impersonate.

.PARAMETER CalendarOnly

By default the script will enumerate all folders. If you want to limit to folders with type "IPF.Appointment" use this switch.

.PARAMETER RootFolder

From where you want to start the search for folders. Default is "MsgFolderRoot". Other posible value is "Root".

.PARAMETER Server

By default the script tries to retrieve the EWS endpoint via Autodiscover. If you want to run the script against a specific server, just provide the name in this parameter. Not the URL!

.PARAMETER Threads

How many threads will be created. Be careful as this is really CPU intensive! By default 10 threads will be created. I limit the number of threads to 20.

.PARAMETER MultiThread

If you want to run the script multi-threaded use this switch. By default the script don't use threads.

.PARAMETER MaxResultTime

The timeout for a job, when using multi-threads. Default is 240 seconds.

.PARAMETER MrMapi

The path to MrMapi.exe. Only full path is accepted!

.PARAMETER UseMrMapi

Switch to tell the script to use MrMapi or not.

.PARAMETER TrustAnySSL

Switch to trust any certificate.

.PARAMETER SearchUnknownOnly

Returns only folders with invalid permissions.

.EXAMPLE

#run the script against a single mailbox
.\Get-MailboxFolderPermissionEWS.PS1 -EmailAddress trick@adatum.com

#pipe all the mailboxes on database USEX01-DB01 into the script
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1

#pipe all the mailboxes on database USEX01-DB01 into the script and use 10 threads
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Threads 10 -MultiThread

#pipe all the mailboxes on database USEX01-DB01 into the script and use 10 threads
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Threads 10 -MultiThread

#same as the previous, but this time using different credentials and impersonate
$SA= Get-Credential adatum\serviceaccount
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Threads 10 -MultiThread -Server ex01.adatum.com -UseDefaultCred 0 -Impersonate -Credentials $SA

#to export the result just use the Export-CSV CmdLet
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 | Export-Csv -NoTypeInformation -Path .\result.csv

#scan multiple mailboxes for invalid permission entries using ApplicationImpersonation, give credentials and multithreading
Get-MailboxDatabase dee16dag01* | Get-Mailbox -ResultSize Unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Impersonate -Credentials $cred -MultiThread -SearchUnknownOnly

.NOTES
#>

[CmdletBinding()]
Param (
    [parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
    [Alias('PrimarySmtpAddress')]
    $EmailAddress,

    [parameter( Mandatory=$false, Position=1)]
    [System.Management.Automation.PsCredential]$Credentials,

    [parameter( Mandatory=$false, Position=2)]
    [System.Management.Automation.SwitchParameter]$Impersonate,

    [parameter( Mandatory=$false, Position=3)]
    [System.Management.Automation.SwitchParameter]$CalendarOnly,

    [parameter( Mandatory=$false, Position=4)]
    [ValidateSet("MsgFolderRoot", "Root")]
    [System.String]$RootFolder="MsgFolderRoot",

    [parameter( Mandatory=$false, Position=5)]
    [System.String]$Server,

    [parameter( Mandatory=$false, Position=6)]
    [ValidateRange(0,20)]
    [System.Int16]$Threads= '15',

    [parameter( Mandatory=$false, Position=7)]
    [System.Management.Automation.SwitchParameter]$MultiThread,

    [parameter( Mandatory=$false, Position=8)]
    [System.Int16]$MaxResultTime='240',

    [parameter( Mandatory=$false, Position=9)]
    [ValidateScript({If (!($_.Contains('.\')) -and (Test-Path $_ -PathType leaf)) {$True} Else {Write-Warning "MrMapi could not be found or relative path used!Please use absolute path!"}})]
    [System.String]$MrMapi,

    [parameter( Mandatory=$false, Position=10)]
    [System.Management.Automation.SwitchParameter]$UseMrMapi,

    [parameter( Mandatory=$false, Position=11)]
    [System.Management.Automation.SwitchParameter]$TrustAnySSL,

    [parameter( Mandatory=$false, Position=12)]
    [System.Management.Automation.SwitchParameter]$SearchUnknownOnly,

    [parameter( Mandatory=$false, Position=13)]
    [System.Management.Automation.SwitchParameter]$ClearUnknown

)

Begin {
    #sanity check for MrMapi
    If ($UseMrMapi -and !($MrMapi)) {
        Write-Warning "Please provide path for MrMapi!"
        Break
    }
    #initiate runspace and make sure we are using single-threaded apartment STA
    $Jobs = @()
    $Sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
    $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads,$Sessionstate, $Host)
    $RunspacePool.ApartmentState = "STA"
    $RunspacePool.Open()
    [System.Int32]$j='1'
    $Timer = [System.Diagnostics.Stopwatch]::StartNew()
}

Process {

    #start function Get-MailboxFolderPermission
    function Get-MailboxFolderPermission {
        Param(
            [parameter( ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
            [Alias('PrimarySmtpAddress')]
            [System.String]$EmailAddress,

            [parameter( Mandatory=$false, Position=1)]
            [System.Management.Automation.PsCredential]$Credentials,

            [parameter( Mandatory=$false, Position=2)]
            [System.Boolean]$Impersonate=$false,
    
            [parameter( Mandatory=$false, Position=3)]
            [System.Boolean]$CalendarOnly=$false,

            [parameter( Mandatory=$false, Position=4)]
            [ValidateSet("MsgFolderRoot", "Root")]
            [System.String]$RootFolder="MsgFolderRoot",

            [parameter( Mandatory=$false, Position=5)]
            [System.String]$Server,

            [parameter( Mandatory=$false, Position=6)]
            [ValidateScript({If (!($_.Contains('.\')) -and (Test-Path $_ -PathType leaf)) {$True} Else {Write-Warning "MrMapi could not be found or relative path used!Please use absolute path!"}})]
            [System.String]$MrMapi,

            [parameter( Mandatory=$false, Position=7)]
            [System.Boolean]$UseMrMapi,

            [parameter( Mandatory=$false, Position=8)]
            [System.Boolean]$TrustAnySSL,

            [parameter( Mandatory=$false, Position=9)]
            [System.Int32]$ProgressID,

            [parameter( Mandatory=$false, Position=10)]
            [System.Boolean]$SearchUnknownOnly,

            [parameter( Mandatory=$false, Position=11)]
            [System.Boolean]$ClearUnknown
        )

        function Cleanup-FolderPermission
        {
            [CmdletBinding()]
            Param(
                [parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$false, Position=0)]
                [Alias('PrimarySmtpAddress')]
                $EmailAddress,
                
                [parameter(Mandatory=$true, Position=1)]
                [System.String]$FolderID,
                
                [parameter(Mandatory=$true, Position=2)]
                [System.String]$ChangeKey,
        
                [parameter( Mandatory=$false, Position=3)]
                [System.Management.Automation.PsCredential]$Credentials,
        
                [parameter( Mandatory=$false, Position=4)]
                [System.Boolean]$Impersonate,
                
                [parameter( Mandatory=$true, Position=5)]
                $URL

            )

            Begin
            {
                [System.Boolean]$returnValue = $true
                # SOAP templates
                $GetFolder = "
                    <?xml version=`"1.0`" encoding=`"utf-8`"?>
                      <soap:Envelope xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:m=`"http://schemas.microsoft.com/exchange/services/2006/messages`" xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`">
                        <soap:Header>
                        <t:RequestServerVersion Version=`"Exchange2010_SP2`" />"
                If ($Impersonate)
                {
                    $GetFolder += "
                        <t:ExchangeImpersonation>
                            <t:ConnectingSID>
                            <t:SmtpAddress>$($EmailAddress)</t:SmtpAddress>
                            </t:ConnectingSID>
                        </t:ExchangeImpersonation>"
                }

                $GetFolder += "
                    </soap:Header>
                      <soap:Body>
                      <m:GetFolder>
                          <m:FolderShape>
                          <t:BaseShape>IdOnly</t:BaseShape>
                          <t:AdditionalProperties>
                              <t:FieldURI FieldURI=`"folder:PermissionSet`" />
                          </t:AdditionalProperties>
                          </m:FolderShape>
                          <m:FolderIds>
                          <t:FolderId Id=`"$($FolderID)`" ChangeKey=`"$($ChangeKey)`" />
                          </m:FolderIds>
                      </m:GetFolder>
                      </soap:Body>
                    </soap:Envelope>"

                $Update = "
                    <?xml version=`"1.0`" encoding=`"utf-8`"?>
                      <soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`">
                        <soap:Header>
                        <t:RequestServerVersion Version=`"Exchange2010_SP2`" />"

                If ($Impersonate)
                {
                    $Update += "
                        <t:ExchangeImpersonation>
                            <t:ConnectingSID>
                            <t:SmtpAddress>$($EmailAddress)</t:SmtpAddress>
                            </t:ConnectingSID>
                        </t:ExchangeImpersonation>"
                }

                $Update += "
                    </soap:Header>
                      <soap:Body>
                        <UpdateFolder xmlns=`"http://schemas.microsoft.com/exchange/services/2006/messages`" xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`">
                          <FolderChanges>
                            <t:FolderChange>
                              <t:FolderId Id=`"$($FolderID)`" ChangeKey=`"$ChangeKey`"/>
                              <t:Updates>
                                <t:SetFolderField>
                                <t:FieldURI FieldURI=`"folder:PermissionSet`"/>
                                  <t:FolderType>
                                    <t:Bla>
                                    </t:Bla>
                                  </t:FolderType>
                                </t:SetFolderField>
                              </t:Updates>
                            </t:FolderChange>
                          </FolderChanges>
                        </UpdateFolder>
                      </soap:Body>
                    </soap:Envelope>"

            }
            Process
            {
                try{
                    $GetParams = @{
                        Uri = $Url
                        Method = 'Post'
                        Body = $($GetFolder.TrimStart())
                        ContentType = 'text/xml'
                        Verbose = $false
                    }

                    $Headers = @{}

                    if ($Credentials)
                    {
                        $CredString = "$($Credentials.UserName.ToString()):$($Credentials.GetNetworkCredential().password.ToString())"
                        $encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($CredString))
                        $Headers.Add('Authorization',"Basic $($encodedCreds)")
                    }

                    #query folder
                    $Headers.Add('Accept','application/xml')
                    $Headers.Add('Accept-Encoding','gzip,deflate')
                    $GetParams.Add('Headers',$Headers)
                    [XML]$GetResp = (Invoke-WebRequest @GetParams).Content

                    if (-not [System.String]::IsNullOrEmpty($GetResp))
                    {
                        #query response for PermissionSet
                        $root = $GetResp.get_DocumentElement()
                        $NamespaceURI = "http://schemas.microsoft.com/exchange/services/2006/types"
                        $ns = New-Object System.Xml.XmlNamespaceManager($GetResp.NameTable)
                        $ns.AddNameSpace("ns",$NamespaceURI)
                        $XPath  = "//PermissionSet"
                        $XPath = $XPath -replace "/(?!/)", "/ns:"
                        $PermNode = $root.SelectSingleNode($XPath,$ns)

                        if ($PermNode)
                        {
                            #replace FolderType with the current type
                            $type = $GetResp.Envelope.Body.GetFolderResponse.ResponseMessages.GetFolderResponseMessage.Folders.ChildNodes.LocalName
                            $Update = $Update.Replace('FolderType',$type)
                            #convert to XML
                            [xml]$Update = $Update.TrimStart()
                            #import PermissionSet XML node
                            $Update.Envelope.Body.UpdateFolder.FolderChanges.FolderChange.Updates.SetFolderField.$type.AppendChild($Update.ImportNode($root.SelectSingleNode($XPath,$ns),$true)) | Out-Null
                            #remove helper XML node
                            $updateRoot = $Update.get_DocumentElement()
                            $NamespaceURI = "http://schemas.microsoft.com/exchange/services/2006/types"
                            $ns = New-Object System.Xml.XmlNamespaceManager($update.NameTable)
                            $ns.AddNameSpace("ns",$NamespaceURI)
                            $BlaNode = $updateRoot.SelectSingleNode("//ns:Bla",$ns)
                            $UnknownEntries = $updateRoot.SelectSingleNode("//ns:UnknownEntries",$ns)
                            $BlaNode.get_ParentNode().RemoveChild($BlaNode) | Out-Null
                            $UnknownEntries.get_ParentNode().RemoveChild($UnknownEntries) | Out-Null

                            $UpdateParams = @{
                                Uri = $Url
                                Method = 'Post'
                                Body = $Update
                                ContentType = 'text/xml'
                                Verbose = $false
                            }

                            $UpdateParams.Add('Headers',$Headers)
                            $UpdatePost = Invoke-WebRequest @UpdateParams
                        }
                        
                    }
                }
                catch{
                    Write-Verbose "Error occured while clearing UnknownEntries"
                    $returnValue = $false
                }
            }
            End
            {
                $returnValue
            }
        }

        try {
            $MailboxName = $EmailAddress
            ## Load Managed API dll
            ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
            $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
            If (Test-Path $EWSDLL) {
                Import-Module $EWSDLL
            }
            Else {
                    "$(get-date -format yyyyMMddHHmmss):"
                    "This script requires the EWS Managed API 1.2 or later."
                    "Please download and install the current version of the EWS Managed API from"
                    "http://go.microsoft.com/fwlink/?LinkId=255472"
                    ""
                    "Exiting Script."
                    exit
            }

            ## Set Exchange Version
            $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2

            ## Create Exchange Service Object
            $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
            #$service.PreAuthenticate = $true
            ## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials
            If ($Credentials) {
                #Credentials Option 1 using UPN for the windows Account
                #$psCred = Get-Credential
                $psCred = $Credentials
                $creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
                $service.Credentials = $creds
                #$service.TraceEnabled = $true
            }
            Else {
                #Credentials Option 2
                $service.UseDefaultCredentials = $true
            }

            If ($TrustAnySSL) {
                ## Choose to ignore any SSL Warning issues caused by Self Signed Certificates
                ## Code From http://poshcode.org/624
                ## Create a compilation environment
                $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
                $Compiler=$Provider.CreateCompiler()
                $Params=New-Object System.CodeDom.Compiler.CompilerParameters
                $Params.GenerateExecutable=$False
                $Params.GenerateInMemory=$True
                $Params.IncludeDebugInformation=$False
                $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

                $TASource=@'
                namespace Local.ToolkitExtensions.Net.CertificatePolicy{
                    public class TrustAll : System.Net.ICertificatePolicy {
                        public TrustAll() {}
                        public bool CheckValidationResult(System.Net.ServicePoint sp,
                        System.Security.Cryptography.X509Certificates.X509Certificate cert,
                        System.Net.WebRequest req, int problem) {
                            return true;
                        }
                    }
                }
'@
                $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
                $TAAssembly=$TAResults.CompiledAssembly

                ## We now create an instance of the TrustAll and attach it to the ServicePointManager
                $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
                [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

                ## end code from http://poshcode.org/624
            }

            ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use
            If ($Server) {
                #CAS URL Option 2 Hardcoded
                $uri=[system.URI] "https://$server/ews/exchange.asmx"
                $service.Url = $uri
            }
            Else {
                #CAS URL Option 1 Autodiscover
                $service.AutodiscoverUrl($MailboxName,{$true})
                #"Using CAS Server : " + $Service.url
            }

            ## Optional section for Exchange Impersonation
            If ($Impersonate) {
                $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
                If ($Service.HttpHeaders.keys.Contains("X-AnchorMailbox")) {
                    $Service.HttpHeaders.Remove("X-AnchorMailbox") | Out-Null
                }
                $Service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName)
            }

            $rootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$RootFolder,$MailboxName)
            #Define Extended properties
            $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
            $folderidcnt = $rootFolderId
            #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling
            $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
            #Deep Transval will ensure all folders in the search path are returned
            $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
            $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
            #Add Properties to the Property Set
            $psPropertySet.Add($PR_Folder_Path);
            $fvFolderView.PropertySet = $psPropertySet;

            If ($CalendarOnly) {
                # Search only for Calendar
                $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Appointment")
            }
            Else {
                #The Search filter will exclude any Search Folders
                $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")
            }

            $fiResult = $null
            #loop through folders, when more than 1000
            do {
                $fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)
                # Add Properties for the Folder Property Set
                $PR_NT_SECURITY_DESCRIPTOR = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E27, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
                $folderPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                If ($UseMrMapi){
                    $folderPropset.Add($PR_NT_SECURITY_DESCRIPTOR)
                }
                Else {
                    $folderPropset.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::Permissions)
                }
                $folderPropset.Add($PR_Folder_Path)

                # Anonymous
                $STANDARDUSER_ANONYMOUS = [Microsoft.Exchange.WebServices.Data.StandardUser]::Anonymous
                # Default
                $STANDARDUSER_DEFAULT = [Microsoft.Exchange.WebServices.Data.StandardUser]::Default

                $objcol = @()

                function ConvertToString($ipInputString){
                    $Val1Text = ""
                    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){
                            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))
                            $clInt++
                    }
                    return $Val1Text
                }

                #helper function. Thanks to Danijel Klaric
                function BinToHex {
                    param(
                        [Parameter(
                            Position=0,
                            Mandatory=$true,
                            ValueFromPipeline=$true)
                    ]
                    [Byte[]]$Bin)
                    # assume pipeline input if we don't have an array (surely there must be a better way)
                    If ($bin.Length -eq 1) {
                        $bin = @($input)
                    }
                    $return = -join ($Bin | foreach {"{0:X2}" -f $_ })
                    Write-Output $return
                }

                function Get-SID {
                    param (
                        [parameter( Mandatory=$False, Position=0)]
                        [System.String]$Domain=$env:USERDOMAIN,
                        [parameter( Mandatory=$true, Position=1)]
                        [System.String]$User
                    )
                    try {
                        $objUser = New-Object System.Security.Principal.NTAccount("$Domain","$User")
                        [System.Security.Principal.SecurityIdentifier]$sid = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
                        $sid
                    }
                    catch {
                        #create object
                        $returnValue = New-Object -TypeName PSObject
                        #get all properties from last error
                        $ErrorProperties =$Error[0] | Get-Member -MemberType Property
                        #add existing properties to object
                        foreach ($Property in $ErrorProperties){
                            if ($Property.Name -eq 'InvocationInfo'){
                                $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                            }
                            else {
                                $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                            }
                        }
                        #return object
                        $returnValue
                    }
                }

                function Get-UserForSID {
                    param (
                        [parameter( Mandatory=$true, Position=0)]
                        [System.String]$SID
                    )
                    try {
                        $objSID = New-Object System.Security.Principal.SecurityIdentifier("$SID")
                        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
                        $objUser.Value
                    }
                    catch {
                        #create object
                        $returnValue = New-Object -TypeName PSObject
                        #get all properties from last error
                        $ErrorProperties =$Error[0] | Get-Member -MemberType Property
                        #add existing properties to object
                        foreach ($Property in $ErrorProperties){
                            if ($Property.Name -eq 'InvocationInfo'){
                                $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                            }
                            else {
                                $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                            }
                        }
                        #return object
                        $returnValue
                    }
                }

                [System.Int32]$i='1'
                ForEach ($Folder in $fiResult) {
                    Write-Progress `
                            -id $ProgressID `
                            -ParentId 1 `
                            -Activity "Processing mailbox - $($MailboxName) with $($fiResult.Folders.count) folders" `
                            -PercentComplete ( $i / $fiResult.Folders.count * 100) `
                            -Status "Remaining: $($fiResult.Folders.count - $i) processing folder: $($Folder.DisplayName)"
                    If (($Folder.Displayname -ne 'System') -and ($Folder.Displayname -ne 'Audits')) {
                        #Write-Output "Working on $($Folder.Displayname)"
                        #Load Properties
                        $Folder.Load($folderPropset)
                        #$Folder.ExtendedProperties[0].Value

                        $foldpathval = $null
                        #Try to get the FolderPath Value and then covert it to a usable String
                        If ($Folder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval)) {
                            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)
                            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }
                            $hexString = $hexArr -join ''
                            #$hexString
                            #$hexString = $hexString.Replace("FEFF", "5C00")
                            $hexString = $hexString.Replace("EFBFBE", "5C")
                            $fpath = ConvertToString($hexString)
                        }
                        #set alternateID
                        $alternateID = new-object Microsoft.Exchange.WebServices.Data.AlternateId("EwsId",$Folder.ID,$MailboxName)
                        If ($UseMrMapi){
                            #get PR_NT_SECURITY_DESCRIPTOR
                            #save HEX Value to file and parse using MRMAPI
                            $tmpfile = [System.IO.Path]::GetTempFileName()
                            BinToHex -Bin ($Folder.ExtendedProperties[0].Value) | Out-File -FilePath $tmpfile -Encoding ascii
                            #& $MrMapi -Parsertype 19 -Input $tmpfile

                            #option 1: run MrMapi.exe directly
                            #$RawFolderSD = & $MrMapi -Parsertype 19 -Input $tmpfile
                            #$FolderSD = $RawFolderSD | %{$_ | ?{($_ -match 'Account*') -or ($_ -match 'SID:*')}} | %{$_ -replace ' ','' -replace 'Account:','' -replace 'SID:'}
                            #$FolderSD = $FolderSD -join ';' -replace ';S-1',"=S-1" -split ';' | select -Unique

                            #option 2: using [System.Diagnostics.Process] and wait for the process
                            $ps = new-object System.Diagnostics.Process
                            $ps.StartInfo.Filename = $MrMapi
                            $ps.StartInfo.Arguments = "-Parsertype 19 -Input $tmpfile"
                            $ps.StartInfo.RedirectStandardOutput = $True
                            $ps.StartInfo.UseShellExecute = $false
                            $ps.start()| Out-Null
                            [string]$RawFolderSD = $ps.StandardOutput.ReadToEnd()
                            $ps.WaitForExit()
                            $FolderSD = $RawFolderSD.Split("`r")| ?{($_ -match 'Account*') -or ($_ -match 'SID:*')} | %{$_ -replace ' ','' -replace 'Account:','' -replace 'SID:' -replace "`n",''}
                            $FolderSD = $FolderSD -join ';' -replace ';S-1',"=S-1" -split ';' | select -Unique
                            # remove temp file
                            Remove-Item $tmpfile
                            If ($FolderSD) {
                                ForEach ($entry in $FolderSD) {
                                    $User = $Entry.Split('=')[0].Split('\')[1]
                                    $LocalSID = (Get-SID -User $User).Value
                                    $data = new-object PSObject
                                    $data | add-member -type NoteProperty -Name Mailbox -Value $MailboxName
                                    $data | add-member -type NoteProperty -Name User -Value $($Entry.Split('=')[0])
                                    $data | add-member -type NoteProperty -Name SIDinSD -Value $($Entry.Split('=')[1])
                                    $data | add-member -type NoteProperty -Name LocalSID -Value $LocalSID
                                    $data | add-member -type NoteProperty -Name FolderName -Value $Folder.DisplayName
                                    $data | add-member -type NoteProperty -Name FolderType -Value $Folder.FolderClass
                                    $data | add-member -type NoteProperty -Name FolderPath -Value $fpath
                                    $data | add-member -type NoteProperty -Name EwsID -Value $Folder.ID
                                    $data | add-member -type NoteProperty -Name StoreID -Value ($service.ConvertId($alternateID,'StoreID')).UniqueId
                                    $data | add-member -type NoteProperty -Name HexEntryID -Value ($service.ConvertId($alternateID,'HexEntryId')).UniqueId
                                    $objcol += $data
                                }
                            }
                        }
                        ElseIf ($SearchUnknownOnly -or $ClearUnknown){
                            ForEach ($Unknown in $Folder.Permissions.UnknownEntries) {
                                $data = new-object PSObject
                                $data | add-member -type NoteProperty -Name Mailbox -Value $MailboxName
                                $data | add-member -type NoteProperty -Name User -Value $Unknown
                                $data | add-member -type NoteProperty -Name FolderName -Value $Folder.DisplayName
                                $data | add-member -type NoteProperty -Name FolderType -Value $Folder.FolderClass
                                $data | add-member -type NoteProperty -Name FolderPath -Value $fpath
                                $data | add-member -type NoteProperty -Name EwsID -Value $Folder.ID.UniqueId
                                $data | add-member -type NoteProperty -Name EwsChangeKey -Value $Folder.ID.ChangeKey
                                $data | add-member -type NoteProperty -Name StoreID -Value ($service.ConvertId($alternateID,'StoreID')).UniqueId
                                $data | add-member -type NoteProperty -Name HexEntryID -Value ($service.ConvertId($alternateID,'HexEntryId')).UniqueId
                                $objcol += $data
                            }

                            If ($ClearUnknown -and (($Folder.Permissions.UnknownEntries).Count -gt 0)) {
                                $FolderPermissionParams = @{
                                    EmailAddress = $MailboxName
                                    FolderID = $Folder.ID.UniqueId
                                    ChangeKey = $Folder.ID.ChangeKey
                                    Impersonate = $Impersonate
                                    URL = $service.Url
                                    Verbose = $VerbosePreference
                                }
                                If ($Credentials) {
                                    $FolderPermissionParams.Add('Credential',$Credentials)
                                }
                                #$ResultValue = Cleanup-FolderPermission -EmailAddress $MailboxName -FolderID $Folder.ID.UniqueId -ChangeKey $Folder.ID.ChangeKey -Credentials $creds -Impersonate $Impersonate -URL $service.Url
                                $ResultValue = Cleanup-FolderPermission @FolderPermissionParams
                                If ($ResultValue) {
                                    Write-Verbose "Succesfully cleaned folder $($Folder.DisplayName)"
                                }
                                Else {
                                    Write-Verbose "Couldn't clean folder $($Folder.DisplayName)with UnknownEntry $($Folder.Permissions.UnknownEntries[0])"
                                }
                            }
                        }
                        Else {
                            ForEach ($Perm in $Folder.Permissions) {
                                $data = new-object PSObject
                                $data | add-member -type NoteProperty -Name Mailbox -Value $MailboxName
                                If ($Perm.UserId.StandardUser -eq $STANDARDUSER_ANONYMOUS){
                                    $data | add-member -type NoteProperty -Name User -Value "Anonymous"
                                }
                                ElseIf($Perm.UserId.StandardUser -eq $STANDARDUSER_DEFAULT){
                                    $data | add-member -type NoteProperty -Name User -Value "Default"
                                }
                                Else {
                                    $data | add-member -type NoteProperty -Name User -Value $Perm.UserId.PrimarySmtpAddress
                                }
                                $data | add-member -type NoteProperty -Name Permissions -Value $Perm.DisplayPermissionLevel
                                $data | add-member -type NoteProperty -Name SID -Value $Perm.UserId.SID
                                $data | add-member -type NoteProperty -Name FolderName -Value $Folder.DisplayName
                                $data | add-member -type NoteProperty -Name FolderType -Value $Folder.FolderClass
                                $data | add-member -type NoteProperty -Name CanCreateItems -Value $Perm.CanCreateItems
                                $data | add-member -type NoteProperty -Name CanCreateSubFolders -Value $Perm.CanCreateSubFolders
                                $data | add-member -type NoteProperty -Name IsFolderOwner -Value $Perm.IsFolderOwner
                                $data | add-member -type NoteProperty -Name IsFolderVisible -Value $Perm.IsFolderVisible
                                $data | add-member -type NoteProperty -Name IsFolderContact -Value $Perm.IsFolderContact
                                $data | add-member -type NoteProperty -Name EditItems -Value $Perm.EditItems
                                $data | add-member -type NoteProperty -Name DeleteItems -Value $Perm.DeleteItems
                                $data | add-member -type NoteProperty -Name ReadItems -Value $Perm.ReadItems
                                $data | add-member -type NoteProperty -Name FolderPath -Value $fpath
                                $data | add-member -type NoteProperty -Name EwsID -Value $Folder.ID
                                $data | add-member -type NoteProperty -Name StoreID -Value ($service.ConvertId($alternateID,'StoreID')).UniqueId
                                $data | add-member -type NoteProperty -Name HexEntryID -Value ($service.ConvertId($alternateID,'HexEntryId')).UniqueId
                                $objcol += $data
                            }
                            }
                        $i++ | Out-Null
                    }
                #end loop of folders
                }
                $objcol
                $fvFolderView.Offset += $fiResult.Folders.Count
            }while($fiResult.MoreAvailable -eq $true)

        }
        catch{
            #create object
            $returnValue = New-Object -TypeName PSObject
            #get all properties from last error
            $ErrorProperties =$Error[0] | Get-Member -MemberType Property
            #add existing properties to object
            foreach ($Property in $ErrorProperties){
                if ($Property.Name -eq 'InvocationInfo'){
                    $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                }
                else {
                    $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                }
            }
            #return object
            $returnValue
        }
        Write-Progress -id $ProgressID -ParentId 1 -Activity "Processing mailbox - $($MailboxName) with $($fiResult.Folders.count) folders" -Status "Ready" -Completed
        #end function Get-MailboxFolderPermission
    }

    #if multi-threaded create jobs
    If ($MultiThread) {
        #create scriptblock from function
        $ScriptBlock = [scriptblock]::Create((Get-ChildItem Function:\Get-MailboxFolderPermission).Definition)
        ForEach($Address in $EmailAddress) {
            try{
                $j++ | Out-Null
                #Write-Host "Adding job for "$Address
                $MailboxName = $Address
                #$PowershellThread = [powershell]::Create().AddScript($ScriptBlock).AddParameter('EmailAddress',$MailboxName).AddParameter('Credentials',$Credentials).AddParameter('Impersonate',$Impersonate).AddParameter('UseDefaultCred',$UseDefaultCred).AddParameter('CalendarOnly',$CalendarOnly).AddParameter('RootFolder',$RootFolder).AddParameter('Server',$Server)
                $PowershellThread = [powershell]::Create().AddScript($ScriptBlock).AddParameter('EmailAddress',$MailboxName)
                $PowershellThread.AddParameter('Credentials',$Credentials) | Out-Null
                $PowershellThread.AddParameter('Impersonate',$Impersonate) | Out-Null
                $PowershellThread.AddParameter('CalendarOnly',$CalendarOnly) | Out-Null
                $PowershellThread.AddParameter('RootFolder',$RootFolder) | Out-Null
                $PowershellThread.AddParameter('Server',$Server) | Out-Null
                If ($MrMapi) {
                    $PowershellThread.AddParameter('MrMapi',$MrMapi) | Out-Null
                }
                If ($SearchUnknownOnly) {
                    $PowershellThread.AddParameter('SearchUnknownOnly',$SearchUnknownOnly) | Out-Null
                }
                If ($ClearUnknown) {
                    $PowershellThread.AddParameter('ClearUnknown',$ClearUnknown) | Out-Null
                }
                $PowershellThread.AddParameter('UseMrMapi',$UseMrMapi) | Out-Null
                $PowershellThread.AddParameter('TrustAnySSL',$TrustAnySSL) | Out-Null
                $PowershellThread.AddParameter('ProgressID',$j) | Out-Null
                $PowershellThread.RunspacePool = $RunspacePool
                $Handle = $PowershellThread.BeginInvoke()
                $Job = "" | Select-Object Handle, Thread, object
                $Job.Handle = $Handle
                $Job.Thread = $PowershellThread
                $Job.Object = $Address
                $Jobs += $Job
            }
            catch {
                #create object
                $returnValue = New-Object -TypeName PSObject
                #get all properties from last error
                $ErrorProperties =$Error[0] | Get-Member -MemberType Property
                #add existing properties to object
                foreach ($Property in $ErrorProperties){
                    if ($Property.Name -eq 'InvocationInfo'){
                        $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
                    }
                    else {
                        $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
                    }
                }
                #return object
                $returnValue
            }
        }
    }
    #if not mutli-threaded start sequential processing
    Else{
        Get-MailboxFolderPermission @psBoundParameters
    }
}

End{

    #monitor and retrieve the created jobs
    If ($MultiThread) {
        $SleepTimer = 200
        $ResultTimer = Get-Date
        While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0) {
            $Remaining = "$($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).object)"
            If ($Remaining.Length -gt 60){
                $Remaining = $Remaining.Substring(0,60) + "..."
            }
            Write-Progress `
                -id 1 `
                -Activity "Waiting for Jobs - $($Threads - $($RunspacePool.GetAvailableRunspaces())) of $Threads threads running" `
                -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
                -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining - $Remaining"

            ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True})) {
                $Job.Thread.EndInvoke($Job.Handle)
                $Job.Thread.Dispose()
                $Job.Thread = $Null
                $Job.Handle = $Null
                $ResultTimer = Get-Date
            }

            If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime) {
                Write-Warning "Child script appears to be frozen for $($Job.Object), try increasing MaxResultTime"
                #Exit
            }

            Start-Sleep -Milliseconds $SleepTimer
            # kill all incomplete threads when hit "CTRL+q"
            If ($Host.UI.RawUI.KeyAvailable) {
                $KeyInput = $Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho")
                If (($KeyInput.ControlKeyState -cmatch '(Right|Left)CtrlPressed') -and ($KeyInput.VirtualKeyCode -eq '81')) {
                    Write-Host -fore red "Kill all incomplete threads....."
                        ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})) {
                            Write-Host -fore yellow "Stopping job $($Job.Object) ...."
                            $Job.Thread.Stop()
                            $Job.Thread.Dispose()
                        }
                    Write-Host -fore red "Exit script now!"
                    Exit
                }
            }
        }
        # clean-up
        $RunspacePool.Close() | Out-Null
        $RunspacePool.Dispose() | Out-Null
        [System.GC]::Collect()
    }
    Write-Verbose "Total runtime:$($Timer.Elapsed.ToString())"
}