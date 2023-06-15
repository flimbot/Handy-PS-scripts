# Handy-PS-scripts

These are just some PowerShell scripts and modeuls which I've created to make my work life a little easier.
Note that they may not be related systems, so I'm just keeping this repository generic.

I don't claim to be the best PowerShell scripter around, so don't expect anything magical.

I'm sure we could appropriate the following logic to be a generic module installer...
```PowerShell
$moduledir = $env:PSModulePath -split ';' | ?{$_ -match "\\$($env:username)\\"} | select -first 1
mkdir ($moduledir + "\AGOL\")
cp AGOL.psm1 ($moduledir + "\AGOL\AGOL.psm1")
```

## AGOL - ArcGIS Online module
 - Get-AGOL-Token-Obj
 - Invoke-AGOL-Request
 - Get-AGOL-Users
 - Get-AGOL-User-Subfolders
 - Get-AGOL-Item
 - Get-AGOL-Group
 - Get-AGOL-Current-User-Info
 - Get-AGOL-Item-Data
 - Get-AGOL-Item-Details

## OTWSMS - OpenText Web Management Server module
 - Resolve-OTWSMS-Error
 - Execute-OTWSMS-RQL
 - Identify-OTWSMS-Version
 - Logout-OTWSMS
 - Login-OTWSMS
 - Get-OTWSMS-Projects 
 - Open-OTWSMS-Session
 - Open-OTWSMS-Project
 - Load-OTWSMS-AsyncJobCategories
 - Load-OTWSMS-AsyncJobs
 - Get-OTWSMS-Pages
 - Get-OTWSMS-Page
 - Disconnect-OTWSMS-Page
 - Get-OTWSMS-ConnectionToPage
 - Get-OTWSMS-ConnectionFromPage
 - Remove-OTWSMS-Page
 - Remove-OTWSMS-PagePermanently
 
## MortgageChoiceBanking - Access to accounts in Mortgage Choice
**In development - Not a module yet**
I started some work on introducing a new module for Mortgage Choice online banking, which I use for some of my accounts.
It's only scratching the surface, but hoping to turn it into a script abstraction to start automating some payments.
It's also a test if everything is accessible using PowerShell (easier for me) to later develop again in TypeScript / Ionic.
**Mortgage Choice recently released a mobile application, so Ionic no longer needed. A command line bank automation may still be an interesting project though**

### Authentication
A little googling of some parameters in URLs lead me to discover that Mortgage Choice uses [Tivoli Access Manager](https://en.wikipedia.org/wiki/IBM_Tivoli_Access_Manager) at least for the authentication at the login screen.
Source which brought me to the conclusion: http://www-01.ibm.com/support/docview.wss?uid=swg1IY90727
