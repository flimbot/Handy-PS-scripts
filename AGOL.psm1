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
        expiration=$minutestoexpire;
        f=$responseformat;
    }

    $reqparamstring = ""
    $reqparams.Keys | Select -First 1 | %{ $reqparamstring += $_ + "=" + $reqparams[$_] }
    $reqparams.Keys | Select -Skip  1 | %{ $reqparamstring += "&" + $_ + "=" + $reqparams[$_] }

    if($useSystemProxy) {
        $proxysettings = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        $resp = Invoke-RestMethod -Method Post -Uri ("https://" + $domain + "/sharing/rest/generateToken?" + $reqparamstring) -Proxy ("http://"+$proxysettings.ProxyServer) -ProxyUseDefaultCredentials
    }
    else {
        $resp = Invoke-RestMethod -Method Post -Uri ("https://" + $domain + "/sharing/rest/generateToken?" + $reqparamstring)
    }

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

function Invoke-AGOL-Request {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$True)][String]$restPath,
        [Parameter(Mandatory=$False)][Hashtable]$urlParameters,
        [Parameter(Mandatory=$False)][Hashtable]$jsonParameters,
        [Parameter(Mandatory=$True)][ValidateSet('Get','Post')]$method
    )  

    #Generate new ticket if expired
    if((Get-Date) -gt $tokenObj.sessionexpiry.AddMinutes(-2)) {
        if($tokenObj.useproxy) {
            $tokenObj = Get-AGOL-Token-Obj -domain $tokenObj.domain -credentials $tokenObj.credentials -minutestoexpire $TokenObj.originalexpiryminutes -useSystemProxy
        }
        else {
            $tokenObj = Get-AGOL-Token-Obj -domain $tokenObj.domain -credentials $tokenObj.credentials -minutestoexpire $TokenObj.originalexpiryminutes
        }
    }

    $url = "https://" + $tokenObj.domain + $restPath #https
    if($urlParameters) {
        $urlParameters["token"] = $tokenObj.token;
        $urlParameters["f"] = $tokenObj.format; #format
        #$urlParameters["referer"] = $tokenObj.domain;

        $url += "?"
        $urlParameters.Keys | Select -First 1 | %{ $url += $_ + "=" + $urlParameters[$_] }
        $urlParameters.Keys | Select -Skip  1 | %{ $url += "&" + $_ + "=" + $urlParameters[$_] }
    }
    else {
        $url += "?token=$($tokenObj.token)&f=$($tokenObj.format)";
    }

    $jsonbody = $NULL
    if($jsonParameters) {
        $jsonbody = ConvertTo-Json $jsonParameters
    }

    if($tokenObj.useproxy) {
        $proxysettings = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

        $resp = Invoke-RestMethod -Method $method -Uri $url -Proxy ("http://"+$proxysettings.ProxyServer) -ProxyUseDefaultCredentials -Body $jsonbody
    }
    else {
        $resp = Invoke-RestMethod -Method $method -Uri $url -Body $jsonbody
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
        $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/portals/self/users" -urlParameters $params -method Get

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
    $resp = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/users/$($username)" -urlParameters $params -method 'Get'
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
        $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/users/$($username)$($folderid)" -urlParameters $params -method 'Get'

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
        $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/community/groups" -method 'Get' -urlParameters $params
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
    $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/items/$($itemId)/data" -method 'Get'
    $json
}

function Get-AGOL-Item-Details {
    param(
        [Parameter(Mandatory=$True)][PSObject]$tokenObj,
        [Parameter(Mandatory=$True)][String][ValidatePattern("[a-f0-9]{32}")]$itemId
    )
    $json = Invoke-AGOL-Request -tokenObj $tokenObj -restPath "/sharing/rest/content/items/$($itemId)" -method 'Get'
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
