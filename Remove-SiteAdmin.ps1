param(
    [Parameter(Mandatory=$false)] 
    [string] $scListFile = "SiteCollectionList.csv",
    [Parameter(Mandatory=$false)] 
    [string] $removeUserFile = "RemoveUsers.csv"
)

$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

#Generate log file
$logFile = "{0}\log_{1}.txt" -f $scriptPath, [System.Guid]::NewGuid().ToString()

#get default csv files
if($scListFile.IndexOf("\") -eq -1)
{
    $scListFile = "$scriptPath\$scListFile"
}
if($removeUserFile.IndexOf("\") -eq -1)
{
    $removeUserFile = "$scriptPath\$removeUserFile"
}

Add-Type -Path $scriptPath\DLLs\Microsoft.SharePoint.Client.dll
Add-Type -Path $scriptPath\DLLs\Microsoft.SharePoint.Client.Runtime.dll

#import the module.
if((Get-Module -Name "Module") -ne $null){
    Remove-Module -Name "Module"
}
Import-Module "$scriptPath\Module"
Write-Log -logFile $logFile -message "Imported module."

#get credentials
$credential = Get-PnPStoredCredential -Name "SPO"

#load CSV sc list
$scList = Import-Csv -Path $scListFile

#load CSV remove users
$removeUserArray=@()
$removeUsers = Import-Csv -Path $removeUserFile
foreach($removeUser in $removeUsers) 
{
    $removeUserArray+=$removeUser.RemoveUser
}

foreach($sc in $scList)
{
    $scUrl = $sc.OD4BSiteURL
    $isConnected=$false

    try {
        #get ClientContext
        $ctx = Get-SPOContext -url $scUrl -SPCredential $credential
        $web = $ctx.Web
        $site = $ctx.Site
        Load-CSOMProperties -object $web -propertyNames @("SiteUsers", "Title")
        Load-CSOMProperties -object $site -propertyNames @("Owner")
        $ctx.ExecuteQuery()
        $siteOwner = $site.Owner
        $isConnected=$true

        Write-Log -logFile $logFile -message "Connected to '$web.Title' at $scUrl."
    }
    catch {
        Write-Log -logFile $logFile -message "Cannot connect to $scUrl." -isError
    }

    if($isConnected)
    {
        foreach($user in $web.SiteUsers)
        {
            If($user.IsSiteAdmin)
            {
                If(($user.UserPrincipalName -ne $siteOwner.UserPrincipalName) -and ($removeUserArray -contains $user.UserPrincipalName))
                {
                    try
                    {
                        $user.IsSiteAdmin=$false
                        $user.Update()
                        $ctx.ExecuteQuery()
                        $msg = "{0} removed from {1}" -f $user.UserPrincipalName, $scUrl
                        Write-Log -logFile $logFile -message $msg
                    }
                    catch
                    {
                        $errMsg = "{0} removed from {1} failed. {2}" -f $user.UserPrincipalName, $scUrl, $_.Exception.Message
                        Write-Log -logFile $logFile -message $errMsg -isError
                    }
                }
            }
        }
    }
}
Write-Log -logFile $logFile -message "Completed."

