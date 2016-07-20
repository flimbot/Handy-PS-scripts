$global:OTWMSPageStatus = @{
 1 = "Page is released.";
 2 = "Page is waiting for release.";
 3 = "Page is waiting for correction.";
 4 = "Page is saved as draft.";
 5 = "Page is not available in this language variant.";
 6 = "Page has never been released in the selected language variant, in which it was created for the first time.";
 10 = "Page has been deleted and is still in the Recycle bin." ;
 50 = "Page will be archived.";
 99 = "Page will be removed completely."
}

#[Enum]::GetNames([PageFlags])
#[PageFlags]::NotBreadcrumb -as [int]
try { Add-Type -TypeDefinition @"
   [System.Flags]
   public enum OTWMSPageFlags
   {
 NotBreadcrumb = 4,
 WaitingForRelease = 64,
 RequiresTranslation = 1024,
 UnlinkedPage = 8192,
 RequiredCorrection = 131072,
 Draft = 262144,
 Released = 524288,
 BreadcrumbStartPoint = 2097152,
 ExternalUrlOrLink = 8388608,
 OwnPageWaitingForRelease = 134217728,
 Locked = 268435456
   }
"@ } catch{}

#Add additional information to the pretty useless error codes
function Resolve-OTWMS-Error {
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)][string] $ErrorText
    )

     $RDErrors = @{
     1 = "The number of modules in the license key does not correspond to the checksum.";
     2 = "The license key is only valid for the Beta test.";
     3 = "The license key is not correct. An error which could not be classified has occurred during the check.";
     4 = "License is no longer valid.";
     5 = "The server IP address is different than specified in the license.";
     6 = "License is not yet valid.";
     7 = "License is a cluster license. This error message is no longer supported beginning with CMS 6.0.";
     8 = "The IP address check in the license key is not correct.";
     9 = "Invalid version of the license key.";
     10 = "There are duplicate modules in the license.";
     11 = "A module in the license is flawed.";
     12 = "There are illegal characters in the license.";
     13 ="The checksum is not correct.";
     14 ="The serial number in the license is not correct.";
     15 ="The serial number of the license key is different from the serial number of the previous license key.";
     16 ="The IP address of the Loopback adapter is not supported in this license.";
     17 ="The license key contains no valid serial number.";
     101 ="The user is already logged on.";
     110 ="(NoRight) The user does not have the required privileges to execute the RQL. The cause of this situation may be that the logon GUID or session key has expired or the session has timed out.";
     201 ="Access to database = `"ioAdministration`" has failed.";
     301 ="Defective asynchronous component.";
     401 ="A project is locked for the executing user or the user level.";
     510 ="The application server is not available.";
     511 ="The application server or the Management Server databases are updated.";
     707 ="The Login GUID does not correspond to the logged on user.";
     800 ="The maximum number of logins for a user has been reached.";
     2910= "References still point to elements of this page.";
     2911 ="At least one element is still assigned as target container to a link.";
     3000 ="Package already exists.";
     3032 ="You have tried to delete a content class on the basis of which pages have already been created.";
     3049 ="Too many users to a license. Login to CMS failed. Please login again later.";
     4005 = "This user name already exists.";
     5001 = "The folder path could not be found or the folder does no longer exist.";
     6001 = "A file is already being used in the CMS.";
     15805 = "You have no right to delete this page.";
     16997 = "You cannot delete the content class. There are pages which were created on the basis of this content class in other projects."
    }

    [void]($ErrorText -match "RDError(\d+)")
    $RDErrors[[int]$matches[1]]
}
Export-ModuleMember -Function Resolve-OTWMS-Error

function Execute-OTWMS-RQL {
    param(
        [Parameter(Mandatory=$True)][String] $RQLString,
        [Parameter(Mandatory=$False)][Switch] $SupressIODATA
    )

    if(!$global:OTWMSVersion) {
        throw "Please run Identify-OTWMS-Version before executing RQL"
    }

    $head = "<IODATA "+("loginguid='$global:OTWMSLoginGuid'"|?{$global:OTWMSLoginGuid})+" "+("sessionguid='$global:OTWMSSessionKey'"|?{$global:OTWMSSessionKey})+">"
    $tail = "</IODATA>"

    if($global:OTWMSVersion -lt 11) {
        #Version 10.1.x
        $Url = "http"+("s"|?{$global:OTWMSServerHTTPS})+"://$global:OTWMSServer/CMS/Navigation/Services/RqlService.asmx"
    }
    else{
        $Url = "http"+("s"|?{$global:OTWMSServerHTTPS})+"://$global:OTWMSServer/CMS/WebService/RqlWebService.svc"
    }

    if(!$global:OTWMSWS -or $global:OTWMSWS.Url -ne $Url) {
        $global:OTWMSWS = New-WebServiceProxy -uri $Url -UseDefaultCredential
    }

    if(!$SupressIODATA) {
        $RQLString = $head+$RQLString+$tail
    }

    if($global:OTWMSVersion -lt 11) {
        #Version 10.1.x

        [xml]$global:OTWMSWS.ExecuteString($RQLString)
    }
    else{
        #Version 11.1-11.2
        $OTWMSError = ""
        $OTWMSResult = ""
        try{
            $OTWMSResponse = $global:OTWMSWS.Execute($RQLString,[ref]$OTWMSError,[ref]$OTWMSResult)
        }
        catch {
            $OTWMSException = $_
        }
        finally {
            if(!$OTWMSResponse -and !$OTWMSResult) {
                throw $OTWMSError
            }
            elseif($OTWMSException) {
                throw $OTWMSException
            }
        }
        [xml]$OTWMSResponse
    }
}
Export-ModuleMember -Function Execute-OTWMS-RQL

function Identify-OTWMS-Version {
    param(
        [Parameter(Mandatory=$True)][String] $ServerHostname
    )
    
    $global:OTWMSServer = $NULL
    $global:OTWMSServerHTTPS = $NULL
    $global:OTWMSVersion = $NULL
    $global:OTWMSWS = $NULL

    try {
        $content = Invoke-WebRequest "https://$ServerHostname/CMS/ioRD.asp?Action=Logout&type=1"
        $global:OTWMSServerSecure = $true
    }
    catch {
        try {
            $content = Invoke-WebRequest "http://$ServerHostname/CMS/ioRD.asp?Action=Logout&type=1"
            $global:OTWMSServerSecure = $false
        }
        catch {
            throw "Cannot connect to Management Server webpage"
        }
    }

    if($content.ParsedHtml.getElementById("VersionLabel").innerText -match "^Management Server (1[01]\.[12])") {
        $global:OTWMSServer = $ServerHostname
        $global:OTWMSVersion = [double]$matches[1]
        $global:OTWMSVersion
    }
    else {
        throw "Management Server version unsupported"
    }
}
Export-ModuleMember -Function Identify-OTWMS-Version

function Logout-OTWMS {
    Execute-OTWMS-RQL -RQLString "<ADMINISTRATION><LOGOUT guid='$global:OTWMSLoginGuid'/></ADMINISTRATION>"

    $global:OTWMSLoginGuid  = $NULL
    $global:OTWMSSessionKey = $NULL
}
Export-ModuleMember -Function Logout-OTWMS

#Sets $global:LoginGuid and returns LoginGuid
function Login-OTWMS {
    param(
        [Parameter(Mandatory=$True)][string] $Username,
        [Parameter(Mandatory=$True)][string] $Password,
        [Parameter(Mandatory=$True)][string] $ServerHostname
    )

    if($global:OTWMSLoginGuid) {
        Logout-OTWMS -LoginGuid $global:OTWMSLoginGuid
    }

    if($ServerHostname -and $global:OTWMSServer -ne $ServerHostname) {
        Identify-OTWMS-Version -ServerHostname $ServerHostname
    }
    
    if($Digest -eq $True) {
        $LoginString = "<ADMINISTRATION action=`"login`" digest=`"1`">Public Digest</ADMINISTRATION>"
    }
    else {
        $LoginString = "<ADMINISTRATION action=`"login`" name=`"$username`" password=`"$password`"/>"
    }
    
    try {
        $LoginResult = Execute-OTWMS-RQL -RQLString $LoginString

        $LoginResult.IODATA.LOGIN.guid | ?{$_ -match "^([0-9A-F]+)" } | %{$matches[1]}
        $global:OTWMSLoginGuid = $LoginResult.IODATA.LOGIN.guid | ?{$_ -match "^([0-9A-F]+)" } | %{$matches[1]}
        $global:OTWMSSessionKey = $LoginResult.IODATA.LOGIN.guid | ?{$_ -match "^([0-9A-F]+)" } | %{$matches[1]}
        $global:OTWMSLoginGuid
    }
    catch [Exception] {
        $ExceptionMessage = $_.Exception.Message
        if($ErrorText -match "RDError(\d+)") {
            $ExceptionMessage = $ExceptionMessage + " " + (Resolve-OTWMS-Error -ErrorText $_.Exception.Message)
        }
        throw $ExceptionMessage
    }
}
Export-ModuleMember -Function Login-OTWMS

function Get-OTWMS-Projects {
    $ListOfProjects = Execute-OTWMS-RQL -RQLString "<IODATA loginguid='$global:OTWMSLoginGuid'><ADMINISTRATION><PROJECTS action='list'/></ADMINISTRATION></IODATA>"
    $ListOfProjects.IODATA.PROJECTS | %{$_.PROJECT} | Select guid,name,testproject,locked
}
Export-ModuleMember -Function Get-OTWMS-Projects

function Open-OTWMS-Session {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match "^[0-9A-F]{32}$"})][string] $ProjectGuid
    )    
    $ValidateResult = Execute-OTWMS-RQL -RQLString "<ADMINISTRATION action='validate' guid='$global:OTWMSLoginGuid'><PROJECT guid='$ProjectGuid' /></ADMINISTRATION>"
    $ValidateResult.IODATA.SERVER.key
} 
Export-ModuleMember -Function Open-OTWMS-Session

function Open-OTWMS-Project {
    param(
        [Parameter(Mandatory=$True)][string] $ProjectName
    )
    Get-OTWMS-Projects | ?{$_.name -eq $ProjectName} | %{Open-OTWMS-Session -ProjectGuid $_.guid}
}
Export-ModuleMember -Function Open-OTWMS-Project
