# Update Firefox to the latest version
# Downloads latest msi from mozilla.org, compares local executable to msi version, runs installation

# https://joelitechlife.ca/2021/04/01/getting-version-information-from-windows-msi-installer/
function Get-MsiVersion {
    param (
        [parameter(Mandatory=$true)] 
        [ValidateNotNullOrEmpty()] 
            [System.IO.FileInfo] $MSIPATH
    ) 
    if (!(Test-Path $MSIPATH.FullName)) { 
        throw "File '{0}' does not exist" -f $MSIPATH.FullName 
    } 
    try { 
        $WindowsInstaller = New-Object -com WindowsInstaller.Installer 
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($MSIPATH.FullName, 0)) 
        $Query = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
        $View = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $Database, ($Query)) 
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null) | Out-Null
        $Record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null ) 
        $Version = $Record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $Record, 1 ) 
        return $Version
    } catch { 
        throw "Failed to get MSI file version: {0}." -f $_
    }       
}

if(Test-Path 'C:\Program Files\Mozilla Firefox\firefox.exe') {
    Write-Host 'Firefox found'
    
    $filename = 'FirefoxSetup.msi'
    Invoke-WebRequest -Uri "https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win&lang=en-US" -OutFile $filename
    $msiVersion = Get-MsiVersion $filename
    $exeVersion = (Get-Item 'C:\Program Files\Mozilla Firefox\firefox.exe').VersionInfo.ProductVersion.Clone()
    
    if([System.Version]$msiVersion -gt [System.Version]$exeVersion) {
        Write-Host "Updating Firefox from version $exeVersion to $msiVersion"
        & MsiExec.exe /i $filename INSTALL_MAINTENANCE_SERVICE=true /qn
    }
}