# Install PSWindowsUpdate module, Fix several issues with Windows Update configurations, ensure WSUS server is reachable
# Built for an inconsistently managed environment

#--------------------------------------------
# Install PowerShell module

if(-not(Get-Module PSWindowsUpdate -ListAvailable)) { 
    Write-Host "Install PSWindowsUpdate module"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    Set-PSRepository -Name PSGallery -InstallationPolicy Untrusted
    #Register-PSRepository -Default
    
    Install-Module PSWindowsUpdate -Confirm:$false -Force
}

#--------------------------------------------
# Make sure windows update service is started, clear cache if stopped

if((Get-Service wuauserv).Status -eq 'Stopped') {
    #Remove-Item "$env:windir\SoftwareDistribution\DataStore\*" -Recurse
    #Start-Service wuauserv

    Reset-WUComponents
}

#--------------------------------------------
# Reset WSUS configuration

# Found off-domain machines which were configured for WSUS server: http://SESWSUS1.SES.NSW.GOV.AU:8530
# This resulted in updates not being found

$wsusServer = (Get-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\ -Name WUServer -ErrorAction SilentlyContinue).WUServer

if($wsusServer) {
    Write-Host "WSUS URL found: $wsusServer"
    $httpError = $null
    try {
        Write-Host "Attempting to request URL"
        $response = Invoke-WebRequest -Uri $wsusServer -ErrorAction Stop
        if($response.StatusCode -ne 200) {
            Write-Host "HTTP call proceeded but had internal error: $($response.StatusCode)"
        }
        else {
            Write-Host "Success!"
        }
    }
    catch {
        Write-Host 'Fail'
        $httpError = $_.Exception.Message
    }

    if($httpError -match 'The remote name could not be resolved') {
        Write-Host 'Stopping Windows Update Service'
        Stop-Service -Name wuauserv
        Write-Host 'Removing WSUS Configuration'
        Remove-Item HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate -Recurse
        Write-Host 'Starting Windows Update Service'
        Start-Service -name wuauserv    
        while((Get-Service wuauserv).Status -ne 'Running') {
            Write-Host 'Windows Update service not started yet..'
            sleep -Seconds 5
        }
    }
}

#--------------------------------------------
# Remove ability for user to pause updates, remove any pause in place, set to auto update
 
if(Get-Item -Path HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate -ErrorAction SilentlyContinue) {
    Set-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\   -Name SetDisablePauseUXAccess -Value 1
    Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU -Name NoAutoUpdate            -Value 0
    @('PauseUpdatesExpiryTime','PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime','PauseQualityUpdatesStartTime','PauseQualityUpdatesEndTime','PauseUpdatesStartTime') | %{
        Remove-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name $_ -ErrorAction SilentlyContinue
    }
}

#--------------------------------------------
# Schedule Windows Updates

try {
    if(Get-Module PSWindowsUpdate -ListAvailable) { 
        # Run update in one minute (allowing script to finish)
        Install-WindowsUpdate -AcceptAll -Install -ScheduleReboot (Get-Date -Date (Get-Date -Format "yyyy-MM-dd 23:00:00")) -ErrorAction Stop -Confirm:$false -Verbose -ScheduleJob (Get-Date).AddMinutes(1)
    }

    # Remove old tasks
    @('Windows Update','Wake computer','Restart computer') | ?{Get-ScheduledTask -TaskName $_ -ErrorAction SilentlyContinue} | %{
        Write-Host 'Removing Task:' $_
        Unregister-ScheduledTask -TaskName $_ -Confirm:$false | Out-Null
    }

    # Set a task to force windows update on start up and at 6pm Tuesday
    Write-Host 'Creating Task: Windows Update'
    $A = New-ScheduledTaskAction -WorkingDirectory "C:\Windows\System32\WindowsPowerShell\v1.0\" -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-Command {Install-WindowsUpdate -AcceptAll -Install -ScheduleReboot (Get-Date -Date (Get-Date -Format `"yyyy-MM-dd 23:00:00`")) -Confirm:$false -ErrorAction SilentlyContinue}"
    $T1 = New-ScheduledTaskTrigger -AtStartup
    $T2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -At 6pm
    $P = New-ScheduledTaskPrincipal -UserId "LOCALSERVICE" -LogonType ServiceAccount
    $S = New-ScheduledTaskSettingsSet
    $D = New-ScheduledTask -Action $A -Trigger @($T1,$T2) -Settings $S -Principal $P
    Register-ScheduledTask 'Windows Update' -InputObject $D | Out-Null

    # Set task to restart computer on Wednesday at 11:15pm, giving 15 minutes warning
    Write-Host 'Creating Task: Restart computer'
    $A = New-ScheduledTaskAction -WorkingDirectory "c:\windows\system32\" -Execute "c:\windows\system32\shutdown.exe" -Argument "/r /t 900"
    $T = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -At 11pm
    $P = New-ScheduledTaskPrincipal -UserId "LOCALSERVICE" -LogonType ServiceAccount
    $S = New-ScheduledTaskSettingsSet
    $D = New-ScheduledTask -Action $A -Trigger $T -Settings $S -Principal $P
    Register-ScheduledTask 'Restart computer' -InputObject $D | Out-Null
}
catch {
    Write-Host 'Error:' $_.Exception.Message
}