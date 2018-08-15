$modulepath = "C:\Users\40821\Powershell modules\"
if(!$env:PSModulePath.Contains($modulepath)) {
    $env:PSModulePath += ";$modulepath"
}

Remove-Module OTWSMS #Unload
Import-Module OTWSMS #Reload


$hostname = Read-Host "Enter the hostname of the OpenText Web Site Management Server"

try {
    Write-Host "Enter credentials to connect"
    Login-OTWSMS -ServerHostname $hostname

    $projects = @(Get-OTWSMS-Projects | ?{$_.locked -eq 0 -and $_.testproject -eq 0} | select guid,name)

    $projectguid = ''
    if($projects.Count -gt 1) {
        Write-Host "Projects:"
        for($i = 0 ; $i -lt $projects.Count ; $i++) {
            Write-Host "`t$i) $($projects[$i].name)"
        }
        $selection = $null
        $selectionint = $null
        do {
            $selection = Read-Host "Choose project by number 0-$($projects.Count-1)"
        } while([int]::TryParse($selection,[ref]$selectionint) -eq $false -and $selectionint -lt 0 -and $selectionint -ge $projects.Count)
        $projectguid = $projects[$selectionint].guid
    }
    else {
        $projectguid = $projects[0]
    }

    Write-Host "Project chosen:" 
    Write-host $projectguid

    Open-OTWSMS-Session -ProjectGuid $projectguid

    #Actions here
    Get-OTWSMS-Pages -ContentClassGuid 'E0F6A9D85EA746E8BEC3C4FBFABDC76D' -MaxHits 20 -PageSize 4 | %{ 
        $additionalinfo = Get-OTWSMS-Page -PageGuid $_.guid
        $_ | Select guid,id,headline,@{name='created';expression={[DateTime]::FromOADate($_.CREATION.date)}},@{name='createdraw';expression={[DateTime]::FromOADate($_.CREATION.date)}},@{name='externaluserid';expression={$additionalinfo.SelectSingleNode("//ELEMENT[@eltname='stf_ExternalUserID']").innerText}}
    } | Export-csv 'C:\Users\40821\Desktop\Desktop\communityevents.csv' -NoTypeInformation

} 
finally {
    Write-Host "Logging out"
    Logout-OTWSMS
}
