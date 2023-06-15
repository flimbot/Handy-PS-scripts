# Harden Adobe Acrobat clients with registry hacks
# Based on https://media.defense.gov/2022/Jan/20/2002924940/-1/-1/1/CTR_CONFIGURING_ADOBE_ACROBAT_READER_20220120.PDF

function SetRegKey {
    param(
        [parameter(Mandatory=$true)][String]$Path,
        [parameter(Mandatory=$true)][String]$Name,
        [parameter(Mandatory=$true)][String]$Value,
        [parameter(Mandatory=$true)][String][ValidateSet('String','DWord','ExpandString','Binary','MultiString','Qword','Unknown')]$Type
    )
    $curVal = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue

    if((Get-Item -Path $Path -ErrorAction SilentlyContinue) -eq $null) {
        $folder = $Path | ?{$_ -match "\w+$"} | %{$Matches.Values}
        New-Item -Path ($Path -replace $folder,"") -Name $folder -ItemType Directory -Force | Out-Null
    }

    if($curVal -eq $null) {
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
    }
    elseif($curVal -ne $Value) {
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force | Out-Null
    }
}

$base = 'HKLM:\Software\Policies\Adobe\Adobe Acrobat\DC'

SetRegKey -Path "$base\FeatureLockDown" -Name 'bEnhancedSecurityStandalone' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bEnhancedSecurityInBrowser' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bProtectedMode' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'iProtectedView' -Value 2 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bEnableProtectedModeAppContainer' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bDisableTrustedSites' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'iFileAttachmentPerms' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bDisablePDFHandlerSwitching' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bEnableFlash' -Value 0 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bEnable3D' -Value 0 -Type DWord
SetRegKey -Path "$base\FeatureLockDown" -Name 'bUpdater' -Value 0 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cDefaultLaunchAttachmentPerms" -Name 'bEnhancedSecurityInBrowser' -Value 3 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cDefaultLaunchAttachmentPerms" -Name 'tBuiltInPermList' -Value 'Default' -Type String
SetRegKey -Path "$base\FeatureLockDown\TrustManager\cDefaultLaunchURLPerms" -Name 'iURLPerms' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cDefaultLaunchURLPerms" -Name 'iUnknownURLPerms' -Value 3 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cSharePoint" -Name 'bDisableSharePointFeatures' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cWebmailProfile" -Name 'bDisableWebmail' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cServices" -Name 'bUpdater' -Value 0 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cServices" -Name 'bToggleAdobeDocumentServices' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cServices" -Name 'bToggleAdobeSign' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cServices" -Name 'bTogglePrefSync' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cServices" -Name 'bToggleWebConnectors' -Value 1 -Type DWord
SetRegKey -Path "$base\FeatureLockDown\cCloud" -Name 'bAdobeSendPluginToggle' -Value 1 -Type DWord
SetRegKey -Path "$base\TrustManager" -Name 'bEnableAlwaysOutlookAttachmentProtectedView' -Value 0 -Type DWord
