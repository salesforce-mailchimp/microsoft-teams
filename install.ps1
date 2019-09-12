# Install NuGet
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force

# Install and import BitTitan.Runbooks.Modules for the Import-BT_Module and Import-ExternalModule functions
Install-Module BitTitan.Runbooks.Modules -Scope CurrentUser -AllowClobber -Force
Import-Module -Name "$($env:USERPROFILE)\Documents\WindowsPowerShell\Modules\BitTitan.Runbooks.Modules" -Force

# Import BitTitan.Runbooks modules
Import-BT_Module BitTitan.Runbooks.Common
Import-BT_Module BitTitan.Runbooks.Csv

Install-Module -Name MicrosoftTeams -RequiredVersion 1.0.1 -Scope CurrentUser -AllowClobber -Force
Install-Module AzureAD -RequiredVersion 2.0.2.26 -Scope CurrentUser -AllowClobber -Force -Verbose

# Update this so that it can be executed during build?yes
# aaa
# aaa
# aaa
# aaa