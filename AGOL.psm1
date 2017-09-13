<#
	Module for accessing ArcGIS Online content
#>

function Get-AGOL-Token-Obj {
    param(
        [Parameter(Mandatory=$True)][string]$domain,
        [Parameter(Mandatory=$True)][pscredential]$credentials,
        [Parameter(Mandatory=$False)][int]$minutestoexpire = 60,
        [switch]$useSystemProxy
    )
    $responseformat = "json";

    $reqparams = @{ 
        username=$credentials.UserName; 
        password=$credentials.GetNetworkCredential().Password; 
        referer=("https://" + $domain);
        'client'     = 'referer';
        expiration=$minutestoexpire;
        f=$responseformat;
    }

    $url = ("https://" + $domain + "/sharing/rest/generateToken")
    if($useSystemProxy) {
        $resp = Invoke-RestMethod -Method Post -Uri $url -Body $reqparams -Proxy ([System.Net.WebRequest]::GetSystemWebProxy().GetProxy($url).AbsoluteUri) -ProxyUseDefaultCredentials -ErrorAction Stop
    }
    else {
        $resp = Invoke-RestMethod -Method Post -Uri $url -Body $reqparams -ErrorAction Stop
    }

    if($resp.token) {
        New-Object PSObject -Property @{
            token=$resp.token;
            domain=$domain;
            sessionexpiry=(Get-Date -Date "01/01/1970").AddMilliseconds($resp.expires).ToLocalTime(); #epoch time to local
            format=$responseformat;
            credentials=$creds;
            useproxy=$useSystemProxy;
            originalexpiryminutes=$minutestoexpire;
        }
    } 
    else {
        throw 'Token not returned'
    }
}

function Invoke-AGOL-Request {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$True)][String]$restPath,
        [Parameter(Mandatory=$False)][Hashtable]$parameters,
        [Parameter(Mandatory=$True)][ValidateSet('Get','Post')]$method
    )  

    #Generate new ticket if near expiration, reauthenticate
    if((Get-Date) -gt $tokenObj.sessionexpiry.AddSeconds(-30)) {
        if($tokenObj.useproxy) {
            $tokenObj = Get-AGOL-Token-Obj -domain $tokenObj.domain -credentials $tokenObj.credentials -minutestoexpire $TokenObj.originalexpiryminutes -useSystemProxy
        }
        else {
            $tokenObj = Get-AGOL-Token-Obj -domain $tokenObj.domain -credentials $tokenObj.credentials -minutestoexpire $TokenObj.originalexpiryminutes
        }
    }

    $url = "https://" + $tokenObj.domain + $restPath #https
    if(!$parameters) {
        $parameters = @{}
    }
    $parameters["token"] = $tokenObj.token;
    $parameters["f"] = $tokenObj.format;

    if($tokenObj.useproxy) {
        $resp = Invoke-RestMethod -Method $method -Uri $url -Proxy ([System.Net.WebRequest]::GetSystemWebProxy().GetProxy($url).AbsoluteUri) -ProxyUseDefaultCredentials -Body $parameters
    }
    else {
        $resp = Invoke-RestMethod -Method $method -Uri $url -Body $parameters
    }

    $resp
}

function Get-AGOL-Users {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj
    ) 

    $allusers = @()

    $params = @{
        start=1;
        num=10;
        sortField="fullname";
        sortOrder="asc";
    }
    do {
        $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/portals/self/users" -parameters $params -method Get

        $json.users | %{$allusers += $_}

        if($json.nextStart -gt 0) {
            $params["start"] = $json.nextStart
        }
    } while ($json.nextStart -gt 0)

    $allusers
}

function Get-AGOL-User-Subfolders {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$False)][String]$username
    ) 
    
    if(!$username) {
        Get-AGOL-Current-User-Info -tokenObj $tokenObj | %{ $username = $_.user.username }
    }

    $folders = @{}
    $params = @{
        num=0;
        start=0
    }
    $resp = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/users/$($username)" -parameters $params -method Get
    $resp.folders | select id,title
}

function Get-AGOL-Item {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$False)][String]$username,
        [Parameter(Mandatory=$False)][String][ValidatePattern("[a-f0-9]{32}")]$folderid,
        [Parameter(Mandatory=$False)][String]$foldername
    )

    if(!$username) {
        Get-AGOL-Current-User-Info -tokenObj $tokenObj | %{ $username = $_.user.username }
    }

    if($foldername -and !$folderid) {
        Get-AGOL-User-Subfolders -tokenObj $tokenObj -username $username | ?{ $_.title -eq $foldername } | %{ $folderid = $_.id }
    }

    if($folderid) {
        $folderid = "/"+$folderid
    }

    $items = @()

    $params = @{
        start=1;
        num=10
    }

    do {
        $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/users/$($username)$($folderid)" -parameters $params -method Get

        $json.items | %{ $items += $_ }

        if($json.nextStart -gt 0) {
            $params["start"] = $json.nextStart
        }
    } while ($json.nextStart -gt 0)
    
    $items
}

function Get-AGOL-Group {
 param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$False)][String]$filter = "*"
    )
    
    $currentUser = Get-AGOL-Current-User-Info -tokenObj $tokenObj

    $groups = @()

    $params = @{
        start=1;
        num=10;
        sortField='title';
        sortOrder='asc';
        q="($filter%20orgid%3A$($currentUser.user.orgId))"
    }

    do {
        $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/community/groups" -parameters $params -method Get
        $json.results | %{ $groups += $_ }

        if($json.nextStart -gt 0) {
            $params["start"] = $json.nextStart
        }
    } while ($json.nextStart -gt 0)

    $groups
}

function Get-AGOL-Current-User-Info {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj
    )
    $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/portals/self" -method 'Get'
    $json
}

function Get-AGOL-Item-Data {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$True)][String][ValidatePattern("[a-f0-9]{32}")]$itemId
    )
    $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/items/$($itemId)/data" -method Get
    $json
}

function Get-AGOL-Item-Details {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$True)][String][ValidatePattern("[a-f0-9]{32}")]$itemId
    )
    $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/items/$($itemId)" -method Get
    $json
}


<#-- Export Methods --#>
Export-ModuleMember -Function Get-AGOL-Token-Obj
Export-ModuleMember -Function Get-AGOL-Users
Export-ModuleMember -Function Get-AGOL-User-Subfolders
Export-ModuleMember -Function Get-AGOL-Item
Export-ModuleMember -Function Get-AGOL-Group
Export-ModuleMember -Function Get-AGOL-Current-User-Info
Export-ModuleMember -Function Get-AGOL-Item-Data
Export-ModuleMember -Function Get-AGOL-Item-Details
