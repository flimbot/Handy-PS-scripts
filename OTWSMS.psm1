$global:OTWSMSPageStatus = @{
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
   public enum OTWSMSPageFlags
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
function Resolve-OTWSMS-Error {
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
Export-ModuleMember -Function Resolve-OTWSMS-Error

function Execute-OTWSMS-RQL {
    param(
        [Parameter(Mandatory=$True)][String] $RQLString,
        [Parameter(Mandatory=$False)][Switch] $SupressIODATA
    )

    if(!$global:OTWSMSVersion) {
        throw "Please run Identify-OTWSMS-Version before executing RQL"
    }

    $head = "<IODATA "+("loginguid='$global:OTWSMSLoginGuid'"|?{$global:OTWSMSLoginGuid})+" "+("sessionguid='$global:OTWSMSSessionKey'"|?{$global:OTWSMSSessionKey})+">"
    $tail = "</IODATA>"

    if($global:OTWSMSVersion -lt 11) {
        #Version 10.1.x
        $Url = "http"+("s"|?{$global:OTWSMSServerHTTPS})+"://$global:OTWSMSServer/CMS/Navigation/Services/RqlService.asmx"
    }
    else{
        $Url = "http"+("s"|?{$global:OTWSMSServerHTTPS})+"://$global:OTWSMSServer/CMS/WebService/RqlWebService.svc"
    }

    if(!$global:OTWSMSWS -or $global:OTWSMSWS.Url -ne $Url) {
        $global:OTWSMSWS = New-WebServiceProxy -uri $Url -UseDefaultCredential
    }

    if(!$SupressIODATA) {
        $RQLString = $head+$RQLString+$tail
    }

    if($global:OTWSMSVersion -lt 11) {
        #Version 10.1.x

        [xml]$global:OTWSMSWS.ExecuteString($RQLString)
    }
    else{
        #Version 11.1-11.2
        $OTWSMSError = ""
        $OTWSMSResult = ""
        try{
            $OTWSMSResponse = $global:OTWSMSWS.Execute($RQLString,[ref]$OTWSMSError,[ref]$OTWSMSResult)
        }
        catch {
            $OTWSMSException = $_
        }
        finally {
            if(!$OTWSMSResponse -and !$OTWSMSResult) {
                throw $OTWSMSError
            }
            elseif($OTWSMSException) {
                throw $OTWSMSException
            }
        }
        [xml]$OTWSMSResponse
    }
}
Export-ModuleMember -Function Execute-OTWSMS-RQL

function Identify-OTWSMS-Version {
    param(
        [Parameter(Mandatory=$True)][String] $ServerHostname
    )
    
    $global:OTWSMSServer = $NULL
    $global:OTWSMSServerHTTPS = $NULL
    $global:OTWSMSVersion = $NULL
    $global:OTWSMSWS = $NULL

    try {
        $content = Invoke-WebRequest "https://$ServerHostname/CMS/ioRD.asp?Action=Logout&type=1"
        $global:OTWSMSServerSecure = $true
    }
    catch {
        try {
            $content = Invoke-WebRequest "http://$ServerHostname/CMS/ioRD.asp?Action=Logout&type=1"
            $global:OTWSMSServerSecure = $false
        }
        catch {
            throw "Cannot connect to Management Server webpage"
        }
    }

    if($content.ParsedHtml.getElementById("VersionLabel").innerText -match "^Management Server (1[01]\.[12])") {
        $global:OTWSMSServer = $ServerHostname
        $global:OTWSMSVersion = [double]$matches[1]
        Write-Verbose $global:OTWSMSVersion
    }
    else {
        throw "Management Server version unsupported"
    }
}
Export-ModuleMember -Function Identify-OTWSMS-Version

function Logout-OTWSMS {
    if($global:OTWSMSLoginGuid) {
        $resp = Execute-OTWSMS-RQL -RQLString "<ADMINISTRATION><LOGOUT guid='$global:OTWSMSLoginGuid'/></ADMINISTRATION>"
        if($resp) {
            Write-Verbose 'Logged out'
        }
    }
    else {
        Write-Verbose 'No login guid in use'
    }

    $global:OTWSMSLoginGuid  = $NULL
    $global:OTWSMSSessionKey = $NULL
    $global:OTWSMSProjectGuid = $NULL
}
Export-ModuleMember -Function Logout-OTWSMS

#Sets $global:LoginGuid and returns LoginGuid
function Login-OTWSMS {
    param(
        [Parameter(Mandatory=$False)][string] $Username,
        [Parameter(Mandatory=$False)][string] $Password,
        [Parameter(Mandatory=$True)][string] $ServerHostname
    )
    # Couldn't get digest authentication working, so just setting this logic for now.
    $Digest = $false


    if($global:OTWSMSLoginGuid) {
        Logout-OTWSMS -LoginGuid $global:OTWSMSLoginGuid
    }

    if($ServerHostname -and $global:OTWSMSServer -ne $ServerHostname) {
        Identify-OTWSMS-Version -ServerHostname $ServerHostname
    }
    
    # If neither credential entered, launch to an interactive prompt
    if(!$Username -or !$Password) {
        if(!$Username) {
            $Username = $env:USERNAME
        }
        
        $cred = Get-Credential -Message "Credentials for OTWSM on $ServerHostname" -UserName $Username
        $Password = $cred.GetNetworkCredential().password
    }
    
    if($Digest -eq $True) {
        $LoginString = "<ADMINISTRATION action=`"login`" digest=`"1`">Public Digest</ADMINISTRATION>"
    }
    else {
        $LoginString = "<ADMINISTRATION action=`"login`" name=`"$Username`" password=`"$Password`"/>"
    }
    
    try {
        $LoginResult = Execute-OTWSMS-RQL -RQLString $LoginString

        #$LoginResult.IODATA.LOGIN.guid | ?{$_ -match "^([0-9A-F]+)" } | %{$matches[1]}
        $global:OTWSMSLoginGuid = $LoginResult.IODATA.LOGIN.guid | ?{$_ -match "^([0-9A-F]+)" } | %{$matches[1]}
        $global:OTWSMSSessionKey = $LoginResult.IODATA.LOGIN.guid | ?{$_ -match "^([0-9A-F]+)" } | %{$matches[1]}
        
        Write-Verbose "Logged in: $global:OTWSMSLoginGuid"
    }
    catch [Exception] {
        $ExceptionMessage = $_.Exception.Message
        if($ErrorText -match "RDError(\d+)") {
            $ExceptionMessage = $ExceptionMessage + " " + (Resolve-OTWSMS-Error -ErrorText $_.Exception.Message)
        }
        throw $ExceptionMessage
    }
}
Export-ModuleMember -Function Login-OTWSMS

function Get-OTWSMS-Projects {
    $ListOfProjects = Execute-OTWSMS-RQL -RQLString "<ADMINISTRATION><PROJECTS action='list'/></ADMINISTRATION>"
    $ListOfProjects.IODATA.PROJECTS | %{$_.PROJECT} | Select guid,name,testproject,locked
    Write-Verbose 'Projects returned'
}
Export-ModuleMember -Function Get-OTWSMS-Projects

function Open-OTWSMS-Session {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match "^[0-9A-F]{32}$"})][string] $ProjectGuid
    )    
    $ValidateResult = Execute-OTWSMS-RQL -RQLString "<ADMINISTRATION action='validate' guid='$global:OTWSMSLoginGuid'><PROJECT guid='$ProjectGuid' /></ADMINISTRATION>"
    $global:OTWSMSProjectGuid = $ProjectGuid
    Write-Verbose "Opened session: $($ValidateResult.IODATA.SERVER.key)"
} 
Export-ModuleMember -Function Open-OTWSMS-Session

function Open-OTWSMS-Project {
    param(
        [Parameter(Mandatory=$True)][string] $ProjectName
    )
    Get-OTWSMS-Projects | ?{$_.name -eq $ProjectName} | %{ Open-OTWSMS-Session -ProjectGuid $_.guid }
}
Export-ModuleMember -Function Open-OTWSMS-Project

Function Get-OTWSMS-PagePreview {
    param(
		[Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$ProjectGuid,
		[Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid
    )

    Execute-OTWSMS-RQL -RQLString "<PREVIEW projectguid=`"$($global:OTWSMSProjectGuid)`" loginguid=`"$($global:OTWSMSLoginGuid)`" url=`"/CMS/ioRD.asp`" querystring=`"Action=Preview&amp;Pageguid=$PageGuid`" />"
}
Export-ModuleMember -Function Get-OTWSMS-PagePreview

function Load-OTWSMS-AsyncJobCategories() {
    $actionflag = @{
    	0 = "No action possible";
    	1 = "Start";
    	2 = "Delete";
    	4 = "Activation/Deactivation possible";
    	8 = "Stop";
    	16 = "Details retrievable";
    	32 = "Server adjustable"
    }

    $type = @{
    	0 = "Publication";
    	1 = "Clean up live server";
    	2 = "Escalation procedure";
    	3 = "XML export";
    	4 = "XML import";
    	5 = "Import 3->4";
    	6 = "Copy project";
    	7 = "Inherit publication package";
    	8 = "Check URLs";
    	9 = "RedDot database backup (not implemented)";
    	10 = "Content class replacement (not visible)";
    	11 = "Upload media element (not visible)";
    	12 = "Copy tree segment (not visible)";
    	13 = "Page forwarding";
    	14 = "Scheduled job";
    	15 = "Publishing queue";
    	16 = "Delete pages via FTP";
    	17 = "FTP transfer";
    	18 = "Export instances";
    	19 = "Start user-defined job";
    	20 = "XCMS project notifications";
    	21 = "Check spelling";
    	22 = "Validate page";
    	23 = "Find and replace";
    	24 = "Project report";
    	25 = "Check references to other projects";
    	26 = "Delete pages via FTP"
    }
    
    $loadcategories = Execute-OTWSMS-RQL -RQLString "<ADMINISTRATION><ASYNCQUEUE action=`"loadcategories`"/></ADMINISTRATION>"
    $loadcategories.IODATA.ADMINISTRATION.ASYNCQUEUE.JOBTYPE | %{ New-Object PSObject -property @{ Value = $_.Value; Type = $type[$_.type]; Action = $actionflag[$_.actionflag]; } }
}
Export-ModuleMember -Function Load-OTWSMS-AsyncJobCategories

Function Get-OTWSMS-AsyncJobs {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$LoginGuid
    )

    $Tasks = Execute-OTWSMS-RQL -RQLString "<ADMINISTRATION><ASYNCQUEUE action=`"list`" project=`"`"/></ADMINISTRATION>"
    $Tasks.IODATA.ChildNodes | %{$_.ASYNCQUEUE} | ?{$_.now -eq 1} | %{
        $Task = Execute-OTWSMS-RQL -RQLString "<ADMINISTRATION><ASYNCQUEUE action=`"load`" guid=`"" + $_.guid + "`"/></ADMINISTRATION>"
        $Task.IODATA.ADMINISTRATION.ASYNCQUEUE | %{ New-Object PSObject -property @{
                Guid = $_.guid;
                Name = $_.name; 
                Server = $_.servername; 
                Started = [DateTime]::FromOADate([double]$_.sendstartat);
                LastExecuted = if($_.lastexecute -ne 0){[DateTime]::FromOADate([double]$_.lastexecute)};
                NextExecuted = [DateTime]::FromOADate([double]$_.nextexecute);
                NextAction = $_.nextaction;
                User = $_.username
            }
        }
    }
}
Export-ModuleMember -Function Load-OTWSMS-AsyncJobs

#TODO Content Class filter
#Date filter
function Get-OTWSMS-Pages {
    #Based on RQL of PAGE/xsearch
    param(
        [Parameter(Mandatory=$False)][ValidateRange(1,60)][int] $PageSize = 60,
        [Parameter(Mandatory=$False)][ValidateRange(1,500)][int] $MaxHits = 100,
        [Parameter(Mandatory=$False)][int] $PageId,
		[Parameter(Mandatory=$False)][ValidateSet("linked","unlinked","recyclebin","active","all")][String] $SpecialPageType,        
        [Parameter(Mandatory=$False)][ValidateSet("checkedout","waitingforrelease","waitingforcorrection","pagesinworkflow","resubmitted","released")][String] $PageState,

        [Parameter(Mandatory=$False)][DateTime] $MinDateCreated,
        [Parameter(Mandatory=$False)][DateTime] $MaxDateCreated,

        [Parameter(Mandatory=$False)][DateTime] $MinDateModified,
        [Parameter(Mandatory=$False)][DateTime] $MaxDateModified,


        [Parameter(Mandatory=$False)][ValidateScript({$_ -match $GuidRegex})][string] $ContentClassGuid
    )
    
    $req = [xml]"<PAGE action=`"xsearch`" orderby=`"headline`" orderdirection=`"ASC`" pagesize=`"$PageSize`" maxhits=`"$MaxHits`" page=`"1`"><SEARCHITEMS></SEARCHITEMS></PAGE>"
    
    if($PageSize) { $req.PAGE.pagesize = $PageSize.ToString() }
    if($MaxHits)  { $req.PAGE.maxhits = $MaxHits.ToString() }

    if($PageId) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"pageid`" value=`"" + $PageId + "`" operator=`"eq`" displayvalue=`"`"></SEARCHITEM>"
    }

    if($SpecialPageType) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"specialpages`" value=`"$SpecialPageType`" operator=`"eq`"/>"
    }
    
    if($PageState) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"pagestate`" value=`"$PageState`" operator=`"eq`" users=`"all`"/>"
    }
    
    if($MinDateCreated) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"createdate`" value=`"$($MinDateCreated.ToOADate())`" operator=`"gt`"/>"
    }

    if($MaxDateCreated) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"createdate`" value=`"$($MaxDateCreated.ToOADate())`" operator=`"lt`"/>"
    }

    if($MinDateModified) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"changedate`" value=`"$($MinDateModified.ToOADate())`" operator=`"gt`"/>"
    }

    if($MaxDateModified) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"changedate`" value=`"$($MaxDateModified.ToOADate())`" operator=`"lt`"/>"
    }

    if($ContentClassGuid) {
        $req.SelectSingleNode("/PAGE/SEARCHITEMS").InnerXml += "<SEARCHITEM key=`"contentclassguid`" value=`"$ContentClassGuid`" operator=`"eq`"/>"
    }

    
    #implicit first page
    #write-host $req.OuterXml
    $ListOfPages = Execute-OTWSMS-RQL -RQLString $req.OuterXml
        
       
    if($MaxHits -ne -1 -and [int]($ListOfPages.IODATA.PAGES.hits) -gt [int]($ListOfPages.IODATA.PAGES.maxhits)) {
        Write-Host "Warning: Too many pages returned. Be more specific of increase maxhits"
        return
    }
    
    #$ret = @()
    $ListOfPages.IODATA.PAGES.PAGE | ?{$_.guid}
    
    $loops = [math]::ceiling([int]($ListOfPages.IODATA.PAGES.hits) / [int]($ListOfPages.IODATA.PAGES.pagesize))

    for($i=2;$i -lt $loops;$i++) {
        $req.PAGE.SetAttribute("page",$i.ToString())
        $ListOfPages = Execute-OTWSMS-RQL -RQLString $req.OuterXml
        $ListOfPages.IODATA.PAGES.PAGE | ?{$_.guid}
    }

	#$ret
}
Export-ModuleMember -Function Get-OTWSMS-Pages

function Get-OTWSMS-Page {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid
        #[Parameter(Mandatory=$False)][ValidateScript({$_ -match $GuidRegex})][int]$PageID
    )
    
	$Page = Execute-OTWSMS-RQL -RQLString "<PAGE action=`"load`" guid=`"$PageGuid`" option=`"extendedinfo`" contentbased=`"1`"/>"
    $Page.IODATA.PAGE
}
Export-ModuleMember -Function Get-OTWSMS-Page

function Remove-OTWSMS-PageConnection {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid,
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$ListGuid
    )
    
	$Page = Execute-OTWSMS-RQL -RQLString "<LINK action=`"unlink`" guid=`"$ListGuid`"><PAGE guid=`"$PageGuid`"/></LINK>"
    $Page.IODATA
}
Export-ModuleMember -Function Disconnect-OTWSMS-Page

function Add-OTWSMS-PageConnection {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid,
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$ListGuid
    )
    
	$Page = Execute-OTWSMS-RQL -RQLString "<LINKSFROM action=`"save`" pageid=`"`" guid=`"$ListGuid`" reddotcacheguid=`"`"><LINK guid=`"$ListGuid`"/></LINKSFROM>"
    $Page.IODATA
}
Export-ModuleMember -Function Disconnect-OTWSMS-Page

#Where this page is connected (i.e. parent pages)
function Get-OTWSMS-ConnectionToPage {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid
    )
    
	$Page = Execute-OTWSMS-RQL -RQLString "<PAGE guid='$PageGuid'><LINKSFROM action='load' /></PAGE>"
    $Page.IODATA.LINKSFROM.LINK | Select PageGuid,PageId,PageHeadline,eltname,elttype,guid,isreference,islink
}
Export-ModuleMember -Function Get-OTWSMS-ConnectionToPage

#Connections from this page (i.e. child pages)
function Get-OTWSMS-ConnectionFromPage {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid
    )
    
	$Page = Execute-OTWSMS-RQL -RQLString "<PAGE guid='$PageGuid'><LINKS action='load' /></PAGE>"
    $Page.IODATA.PAGE.LINKS.LINK | %{
        $link = $_
        Execute-OTWSMS-RQL -RQLString "<LINK guid='$($link.guid)'><PAGES action='list'/></LINK>" | %{ $_.IODATA.PAGES.PAGE } | Select headline,@{Name="PageGuid";Expression={$_.guid}}, @{Name="PageId";Expression={$_.id}}, @{Name="LinkGuid";Expression={$link.guid}}, @{Name="LinkName";Expression={$link.eltname}}, @{Name="ParentPageGuid";Expression={$PageGuid}}
    }
}
Export-ModuleMember -Function Get-OTWSMS-ConnectionFromPage

function Remove-OTWSMS-Page {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid,
        [switch]$Permanently
    )
    
	$resp = Execute-OTWSMS-RQL -RQLString "<PAGE action='delete' guid='$PageGuid'/>"
    if( $resp.IODATA -notmatch '\s*ok\s*' ) {
        if( $resp.IODATA -match '^\s*$' ) {
            throw 'no page to remove'
        }
        else {
            throw (Resolve-OTWSMS-Error -ErrorText $resp.IODATA)
        }
    }

    if( $Permanently ) {
        Remove-OTWSMS-PagePermanently -PageGuid $PageGuid
    }
}
Export-ModuleMember -Function Remove-OTWSMS-Page

function Remove-OTWSMS-PagePermanently {
    param(
        [Parameter(Mandatory=$True)][ValidateScript({$_ -match $GuidRegex})][string]$PageGuid
    )
    
	$resp = Execute-OTWSMS-RQL -RQLString "<PAGE action='deletefinally' guid='$PageGuid' />"
    if( $resp.IODATA -ne 'ok' ) {
        throw $resp.IODATA
    }
}
Export-ModuleMember -Function Remove-OTWSMS-PagePermanently

