$creds = Get-Credential -Message 'Authenticate to Objective Connect' -UserName (Get-ADUser $env:USERNAME -Properties Mail).Mail
$browse = New-Object System.Windows.Forms.FolderBrowserDialog
if($browse.ShowDialog() -eq 'OK') {
    $outputdir = $browse.SelectedPath + '\'

    $useremail = $creds.UserName
    $pass = $creds.GetNetworkCredential().Password

    $ocUri = 'https://secure.objectiveconnect.com/rest'
    $illegalFilesystemCharRegex = '[<>:"/\|?*]'

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if(!(Test-Path $outputdir)) {
        [void](mkdir $outputdir)
    }

    if(!$global:objconSession) {
        $global:objconSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    }

    function Invoke-ObjectiveConnect {
        param(
            [Parameter(Mandatory=$true)][string]$Uri,
            [Parameter(Mandatory=$false)][Hashtable]$QueryStringParams,
            [Parameter(Mandatory=$true)][Microsoft.PowerShell.Commands.WebRequestMethod]$Method,
            [Parameter(ParameterSetName='Auth', Mandatory=$true)][string]$Username,
            [Parameter(ParameterSetName='Auth', Mandatory=$true)][string]$Password,
            [Parameter(ParameterSetName='Cont', Mandatory=$true)][switch]$ContinueSession = $false,
            [Parameter(Mandatory=$false)][string]$OutputPath = $null
        )

        $localObjConSession = $global:objconSession

        if($QueryStringParams -and ($Uri -notmatch '\?')) {
            $params = @()
            $QueryStringParams.Keys | %{
                $key = $_
                if($QueryStringParams[$key].Count -and $QueryStringParams[$key].Count -gt 0) {
                    $QueryStringParams[$key] | %{ $params += ($key + '=' + $_) }
                }
                else {
                    $params += ($key + '=' + $QueryStringParams[$key])
                }
            }
            $Uri = $Uri + '?' + ($params -join '&')
        }

        $proxy = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($uri)
        $header = @{'Accept'='application/hal+json'}

        if($ContinueSession) {
            $resp = Invoke-RestMethod -WebSession $localObjConSession -Uri $Uri -Method $Method -Headers $header -Proxy $proxy -ProxyUseDefaultCredentials -UseDefaultCredentials -OutFile $OutputPath
        }
        else {
            #Set basic authentication
            $pair = "${Username}:${Password}"
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
            $base64 = [System.Convert]::ToBase64String($bytes)
            $header['Authorization'] = "Basic $base64"

            $resp = Invoke-RestMethod -SessionVariable localObjConSession -Uri $Uri -Method $Method -Headers $header -Proxy $proxy -ProxyUseDefaultCredentials -UseDefaultCredentials -OutFile $OutputPath
        }
        Sleep -Milliseconds 250
 
        $global:objconSession = $localObjConSession
        $resp
    }

    function Get-ObjectiveConnectFolderAssets {
        param(
            [Parameter(Mandatory=$true)][string]$WorkspaceUuid,
            [Parameter(Mandatory=$false)][string]$FolderUuid,
            [switch]$IncludeFolders=$false,
            [Parameter(Mandatory=$false)][string]$PathString = ''
        )

        $queryParams = @{
            'noContentReplacesSeeOther'='true';
            #'length'     ='100';
            'offset'     ='0';
            'parentUuid' =$WorkspaceUuid;
            'sort'       ='modifiedDate:desc'
        }

        if($FolderUuid) {
            $queryParams['parentUuid'] = $FolderUuid
        }

        $ocAssets = Invoke-ObjectiveConnect -Uri "$ocUri/shares/$WorkspaceUuid/assets" -QueryStringParams $queryParams -Method Get -ContinueSession   

        # Return
        $ocAssets | ?{$_.type -ne 'FOLDER'} | %{
            $filepath = ($PathString+($_.name -replace $illegalFilesystemCharRegex,'_')+'.'+$_.extension) 
            Invoke-ObjectiveConnect -Uri "$ocUri/shares/$WorkspaceUuid/assets/$($_.uuid)/contents/latest" -QueryStringParams @{ 'request.preventCache'=(Get-Date).Ticks } -Method Get -ContinueSession -OutputPath $filepath
            Write-Host $filepath
            $_        
        }

        # Recurse
        $ocAssets | ?{$_.type -eq 'FOLDER'} | %{ 
            if($IncludeFolders) {
                $_
            }

            $newfolder = ($PathString + ($_.name -replace $illegalFilesystemCharRegex,'_') + '\')
            if(!(Test-Path $newfolder)) {
                [void](mkdir $newfolder)
            }

            Get-ObjectiveConnectFolderAssets -WorkspaceUuid $WorkspaceUuid -FolderUuid $_.uuid -PathString $newfolder
         }
    }



    $ocUser = Invoke-ObjectiveConnect -Uri "$ocUri/users" -QueryStringParams @{ 'emailAddress' = $useremail } -Method Get -Username $useremail -Password $pass

    $queryParams = @{
        'noContentReplacesSeeOther'='true';
        #'length'         ='100';
        'offset'         ='0';
        'participantUuid'=$ocUser.uuid;
        'userUuid'       =$ocUser.uuid
    }

    $ocWorkspaces = Invoke-ObjectiveConnect -Uri "$ocUri/shares" -QueryStringParams $queryParams -Method Get -ContinueSession
    $ocContent = @{}
    $ocWorkspaces | %{
        $ocContent[$_.uuid] = $_

        $newfolder = ($outputdir+($_.name -replace $illegalFilesystemCharRegex,'_')+'\')
        if(!(Test-Path $newfolder)) {
            [void](mkdir $newfolder)
        }

        Get-ObjectiveConnectFolderAssets -WorkspaceUuid $_.uuid -IncludeFolders -PathString $newfolder | %{ $ocContent[$_.uuid] = $_ }
    }
}