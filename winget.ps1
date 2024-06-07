$env:winget = (Get-ChildItem "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" | Sort-Object -Property LastWriteTime | Select-Object -Last 1).FullName

function Get-WingetApp {
    $output = (Invoke-Expression "& `"$env:winget`" list") -replace 'â€',''

    $cols = $null
    $output | ?{$_ -match "^Name\s+Id\s+Version\s+Available\s+Source"} | %{
        $cols = @{
            'Name'      = @{
                'start' = 0
                'length'= $_.IndexOf('Id')
            };
            'Id'        = @{
                'start' = $_.IndexOf('Id')
                'length'= $_.IndexOf('Version') - $_.IndexOf('Id')
            }
            'Version'   = @{
                'start' = $_.IndexOf('Version')
                'length'= $_.IndexOf('Available') - $_.IndexOf('Version')
            };
            'Available' = @{
                'start' = $_.IndexOf('Available')
                'length'= $_.IndexOf('Source') - $_.IndexOf('Available')
            };
            'Source'    = @{
                'start' = $_.IndexOf('Source')
                'length'= $_.length - $_.IndexOf('Source')
            }
        }
    }

    $output | ?{$_ -notmatch "^\s" -and $_ -notmatch "^\-{20}" -and $_ -notmatch "^Name\s+Id\s+Version\s+Available\s+Source" -and $_.length -gt 0} | %{
        $row = $_
        $o = New-Object PSObject

        $cols.Keys | %{
            $col = $_
            if(($cols[$col].start + $cols[$col].length) -gt $row.Length) {
                $o | Add-Member -MemberType NoteProperty -Name $col -Value ($row.Substring($cols[$col].start,($row.Length - $cols[$col].start)) -replace '¦','').Trim()
            }
            else {
                $o | Add-Member -MemberType NoteProperty -Name $col -Value ($row.Substring($cols[$col].start,$cols[$col].length) -replace '¦','').Trim()
            }
        }

        $o
    }
}

function Update-WingetApp {
    param(
        [Parameter()][String]$Id,
        [Parameter(ValueFromPipeline)]$InputObject
    )
    begin  {}
    process{
        if($InputObject.id) {
            $Id = $InputObject.Id
        } 
        else {
            $Id = $InputObject
        }
        
        $output = (Invoke-Expression "& `"$env:winget`" upgrade --id `"$Id`" --silent --accept-package-agreements --accept-source-agreements") -replace 'â€','' | ?{$_ -notmatch "^\s" -and $_.length -ne 0}
        Write-Host $output
    }
    end    {}
}

$output = Get-WingetApp | ?{$_.Name -match 'Google'} | Update-WingetApp
