# Handy-PS-modules

These are just some PowerShell modules which I've created to make my work life a little easier.
Note that they may not be related systems, so I'm just keeping this repository generic.

I don't claim to be the best PowerShell scripter around, so don't expect anything magical.

I'm sure we could appropriate the following logic to be a generic module installer...
```PowerShell
$moduledir = $env:PSModulePath -split ';' | ?{$_ -match "\\$($env:username)\\"} | select -first 1
mkdir ($moduledir + "\AGOL\")
cp AGOL.psm1 ($moduledir + "\AGOL\AGOL.psm1")
```

## AGOL - ArcGIS Online module


## OTWMS - OpenText Web Management Server module
