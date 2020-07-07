$url = "https://dev.sp2019.ezcode.org/sites/marketing"
$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

Add-Type -Path $scriptPath\DLLs\Microsoft.SharePoint.Client.dll
Add-Type -Path $scriptPath\DLLs\Microsoft.SharePoint.Client.Runtime.dll

$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)

#[System.Net.CredentialCache]::DefaultCredentials
$context.Credentials = [System.Net.CredentialCache]::DefaultCredentials

$web = $context.Web
$context.Load($web)
$context.ExecuteQuery()

Write-Host $web.Title

$appId=[AppId]
$appSecret=[AppSecret]
$url=[siteurl]

Connect-PnPOnline -Url $url -AppId $appId -AppSecret $appSecret

Get-PnPUnifiedGroup
Get-PnPSiteCollectionAdmin
set-pnpsitec