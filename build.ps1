Import-Module PowerShellGet

#Publish to PSGallery and install/import locally

Publish-Module -Path .\AMAG-SMSPowershell -Repository PSGallery -Verbose
Install-Module -Name AMAG-SMSPowershell -Repository PSGallery -Force
Import-Module -Name AMAG-SMSPowershell -Force 