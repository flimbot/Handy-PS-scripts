if(-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Throw 'Must be run as Administrator'
}

$env:winget = (Get-ChildItem "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" | Sort-Object -Property LastWriteTime | Select-Object -Last 1).FullName

function Get-WingetApp {
    param(
        [Parameter()][String]$Id,
        [Parameter(ValueFromPipeline)]$InputObject
    )
    begin  {}
    process{

        # Handle difference between pipeline and non-pipeline requests        
        if(-not $Id) {
            if($InputObject.id) {
                $Id = $InputObject.Id
            } 
            else {
                $Id = $InputObject
            }
        }

        # Handle difference between specific id called
        if($Id) {
            $idarg = " --id $Id"
        }
        else {
            $idarg = ''
        }

        # Retrieve results and remove some of the non-ascii characters
        $output = (Invoke-Expression "& `"$env:winget`" list$idarg") -replace 'â€',''
        
        # Identify the end of the header information
        $headerrow = 0
        for($i = 0 ; $i -lt $output.count ; $i++) {
            if($output[$i] -match "^\-{20}") { 
                $headerrow = $i
            }
        }

        if($headerrow -eq 0) {
            $err = ($output | ?{$_ -notmatch "^\s"}) -join ' '
            Throw $err
        }

        # Find the bounds of each of the header columns
        $cols = @($output | Select -First 1 -Skip ($headerrow-1) | Select-String -Pattern "\w+" -AllMatches | %{
            $_.matches
        } | Select @{N='Name';E={$_.Value}},@{N='Start';E={$_.Index}})

        # Process output
        $output | Select -Skip ($headerrow+1) | %{
            $row = $_
            $o = New-Object PSObject

            for($i = 0 ; $i -lt $cols.Count ; $i++) {
                # If last item
                if(($i+1) -eq $cols.Count) {
                    $len = $row.Length - $cols[$i].Start
                }
                else {
                    $len = $cols[$i+1].Start - $cols[$i].Start
                }
                $o | Add-Member -MemberType NoteProperty -Name $cols[$i].Name -Value ($row.Substring($cols[$i].Start,$len) -replace '¦','').Trim()
            }
            $o
        }

        $Id = $null
    }
    end {}
}

function Get-WingetAppInformation {
    param(
        [Parameter()][String]$Id,
        [Parameter(ValueFromPipeline)]$InputObject
    )
    begin  {}
    process{    
        if(-not $Id) {
            if($InputObject.id) {
                $Id = $InputObject.Id
            } 
            else {
                $Id = $InputObject
            }
        }
        
        $output = (Invoke-Expression "& `"$env:winget`" show --id `"$Id`"")

        $pso = New-Object PSObject
        $lasttoplevel = $null
        $output -split "`r`n" | ?{$_ -match '\w'} | %{
            if($_ -match '^Found ([^\[]+)\s\[([^\]]+)\]') {
                $pso | Add-Member -MemberType NoteProperty -Name 'Name' -Value $matches[1]
                $pso | Add-Member -MemberType NoteProperty -Name 'id'   -Value $matches[2]
            }
            elseif($_ -match '^(\w[\w\s]+):\s(.+)$') {
                if($matches[2].Trim().Length -eq 0) {
                    $pso | Add-Member -MemberType NoteProperty -Name $matches[1] -Value $null
                }
                else {
                    $pso | Add-Member -MemberType NoteProperty -Name $matches[1] -Value $matches[2]
                }
                $lasttoplevel = $matches[1] 
            }
            else {
                #Add logic for sub-properties such as copyright,tags, and installer
            }
        }

        $Id = $null
        $pso
    }
    end    {}
}

function Update-WingetApp {
    param(
        [Parameter()][String]$Id,
        [Parameter(ValueFromPipeline)]$InputObject
    )
    begin  {}
    process{
        if(-not $Id) {
            if($InputObject.id) {
                $Id = $InputObject.Id
            } 
            else {
                $Id = $InputObject
            }
        }
        
        $output_info1 = Get-WingetApp -Id $Id
        $pso = $output_info1 | Select Name,id,@{N='From Version';E={$_.Version}}

        $output_upgrade = (Invoke-Expression "& `"$env:winget`" upgrade --id `"$Id`" --silent --accept-package-agreements --accept-source-agreements --include-unknown") -replace 'â€','' | ?{$_ -notmatch "^\s" -and $_.length -ne 0}
        
        $output_info2 = Get-WingetApp -Id $Id
        $pso | Select *,@{N='To Version';E={$output_info2.Version}},@{N='Update Message';E={$output_upgrade}}

        $Id = $null
    }
    end    {}
}

function Remove-WingetApp {
    param(
        [Parameter()][String]$Id,
        [Parameter(ValueFromPipeline)]$InputObject
    )
    begin  {}
    process{
        if(-not $Id) {
            if($InputObject.id) {
                $Id = $InputObject.Id
            } 
            else {
                $Id = $InputObject
            }
        }
        
        $output = (Invoke-Expression "& `"$env:winget`" uninstall --id `"$Id`" --silent --disable-interactivity --accept-source-agreements --purge") -replace 'â€','' | ?{$_ -notmatch "^\s" -and $_.length -ne 0}
        Write-Host $output
        $Id = $null
    }
    end    {}
}
