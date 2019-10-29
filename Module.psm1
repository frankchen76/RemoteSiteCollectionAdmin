#
# Module.psm1
#
function Get-SPOCredentials()
{
	param(
        [Parameter(Mandatory=$false)] 
		[string] $username=$null,
        [Parameter(Mandatory=$false)] 
        [string] $password="",
        [Parameter(Mandatory=$false)]
        [switch] $requireProxy=$false,
        [Parameter(Mandatory=$false)] 
		[string] $proxyusername=$null,
        [Parameter(Mandatory=$false)] 
		[string] $proxypassword=""
	)
	if($username -eq $null)
	{
		$username = Read-Host -Prompt "Please enter SPO username:"
	}
	if($password -eq "")
	{
		$securePassword = Read-Host -AsSecureString -Prompt "Please enter password:"
		$password=[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))

	}else
	{
		$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    }
    
    $proxyCredential=$null

    if($requireProxy)
    {
        $proxyusername=Read-Host -Prompt "Please enter proxy username:(press enter if no proxy needed.)"
        if($proxyusername -ne "")
        {
            $proxypassword=Read-Host -Prompt "Please enter proxy password:"
        }
        $proxyCredential = New-Object System.Net.NetworkCredential($proxyusername, $proxypassword)
    }

    $ret = [PSCustomObject]@{
        SPOCredential   = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword)
        ProxyCredential = $proxyCredential
    }

	$ret
}
function Get-SPOContext()
{
	param(
        [Parameter(Mandatory=$true)] 
		[string] $url,
		[Parameter(Mandatory=$true)] 
		[System.Net.ICredentials] $SPCredential,
		[Parameter(Mandatory=$false)] 
		[System.Net.NetworkCredential] $ProxyCredential=$null
    )
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
	$context.Credentials = $SPCredential
	if($ProxyCredential -ne $null)
	{
		$null= Register-ObjectEvent -InputObject $context -MessageData $ProxyCredential -EventName ExecutingWebRequest -Action{
	
			$EventArgs.WebRequestExecutor.WebRequest.Proxy.Credentials = $Event.MessageData
		}
	}
	
	$context
}

function Write-Log()
{
    param(
        [Parameter(Mandatory = $true)] 
        [string] $logFile,
        [Parameter(Mandatory = $true)] 
        [string] $message,
        [Parameter(Mandatory = $false)]
        [Switch] $isError = $false
    )
    $currentDate = Get-Date
    if($isError)
    {
        $msgType="ERROR";
    }else
    {
        $msgType="INFO";
    }
    $outputMessage = "[{0}]{1}" -f $currentDate.ToString(), $message
    $logMessage = "[{0}][{1}]{2}" -f $currentDate.ToString(), $msgType, $message
    if($isError)
    {
        Write-Error $outputMessage
    }else
    {
        Write-Host $outputMessage
    }
    if($logFile -ne "")
    {
        Add-Content -Path $logFile -Value $logMessage
    }
}

function Remove-SPOListItems()
{
	param(
        [Parameter(Mandatory=$true)] $ctx,
        [Parameter(Mandatory=$true)] $listname
    )

	$listInstance = $ctx.Web.Lists.GetByTitle($listname);
	$query = New-Object -TypeName Microsoft.SharePoint.Client.CamlQuery
	$query.ViewXml="<View><ViewFields><FieldRef Name='ID' /></ViewFields><RowLimit>5000</RowLimit></View>"
	$done=$false
	while($done -ne $true)
	{
		#$listItems = $listInstance.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
		$listItems = $listInstance.GetItems($query)
		$ctx.Load($listItems)
		$ctx.ExecuteQuery()
		$total=$listItems.Count
		$done=$total -eq 0

		if($total -gt 0)
		{
			for($i=$total-1; $i -ge 0; $i--)
			{
				Write-Host "Delete ID $($listItems[$i].ID) ..."
				$listItems[$i].DeleteObject()
			
				$mod = $i % 100
				if($mod -eq 0)
				{
					$ctx.ExecuteQuery()
				}
			}
			$ctx.ExecuteQuery()
			Write-Host "$total rows were deleted under $listname."
		}
	}
}

function global:Load-CSOMProperties {
    [CmdletBinding(DefaultParameterSetName='ClientObject')]
    param (
        # The Microsoft.SharePoint.Client.ClientObject to populate.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, ParameterSetName = "ClientObject")]
        [Microsoft.SharePoint.Client.ClientObject]
        $object,
 
        # The Microsoft.SharePoint.Client.ClientObject that contains the collection object.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, ParameterSetName = "ClientObjectCollection")]
        [Microsoft.SharePoint.Client.ClientObject]
        $parentObject,
 
        # The Microsoft.SharePoint.Client.ClientObjectCollection to populate.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1, ParameterSetName = "ClientObjectCollection")]
        [Microsoft.SharePoint.Client.ClientObjectCollection]
        $collectionObject,
 
        # The object properties to populate
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "ClientObject")]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "ClientObjectCollection")]
        [string[]]
        $propertyNames,
 
        # The parent object's property name corresponding to the collection object to retrieve (this is required to build the correct lamda expression).
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = "ClientObjectCollection")]
        [string]
        $parentPropertyName,
 
        # If specified, execute the ClientContext.ExecuteQuery() method.
        [Parameter(Mandatory = $false, Position = 4)]
        [switch]
        $executeQuery
    )
 
    begin { }
    process {
        if ($PsCmdlet.ParameterSetName -eq "ClientObject") {
            $type = $object.GetType()
        } else {
            $type = $collectionObject.GetType() 
            if ($collectionObject -is [Microsoft.SharePoint.Client.ClientObjectCollection]) {
                $type = $collectionObject.GetType().BaseType.GenericTypeArguments[0]
            }
        }
 
        $exprType = [System.Linq.Expressions.Expression]
        $parameterExprType = [System.Linq.Expressions.ParameterExpression].MakeArrayType()
        $lambdaMethod = $exprType.GetMethods() | ? { $_.Name -eq "Lambda" -and $_.IsGenericMethod -and $_.GetParameters().Length -eq 2 -and $_.GetParameters()[1].ParameterType -eq $parameterExprType }
        $lambdaMethodGeneric = Invoke-Expression "`$lambdaMethod.MakeGenericMethod([System.Func``2[$($type.FullName),System.Object]])"
        $expressions = @()
 
        foreach ($propertyName in $propertyNames) {
            $param1 = [System.Linq.Expressions.Expression]::Parameter($type, "p")
            try {
                $name1 = [System.Linq.Expressions.Expression]::Property($param1, $propertyName)
            } catch {
                Write-Error "Instance property '$propertyName' is not defined for type $type"
                return
            }
            $body1 = [System.Linq.Expressions.Expression]::Convert($name1, [System.Object])
            $expression1 = $lambdaMethodGeneric.Invoke($null, [System.Object[]] @($body1, [System.Linq.Expressions.ParameterExpression[]] @($param1)))
  
            if ($collectionObject -ne $null) {
                $expression1 = [System.Linq.Expressions.Expression]::Quote($expression1)
            }
            $expressions += @($expression1)
        }
 
 
        if ($PsCmdlet.ParameterSetName -eq "ClientObject") {
            $object.Context.Load($object, $expressions)
            if ($executeQuery) { $object.Context.ExecuteQuery() }
        } else {
            $newArrayInitParam1 = Invoke-Expression "[System.Linq.Expressions.Expression``1[System.Func````2[$($type.FullName),System.Object]]]"
            $newArrayInit = [System.Linq.Expressions.Expression]::NewArrayInit($newArrayInitParam1, $expressions)
 
            $collectionParam = [System.Linq.Expressions.Expression]::Parameter($parentObject.GetType(), "cp")
            $collectionProperty = [System.Linq.Expressions.Expression]::Property($collectionParam, $parentPropertyName)
 
            $expressionArray = @($collectionProperty, $newArrayInit)
            $includeMethod = [Microsoft.SharePoint.Client.ClientObjectQueryableExtension].GetMethod("Include")
            $includeMethodGeneric = Invoke-Expression "`$includeMethod.MakeGenericMethod([$($type.FullName)])"
 
            $lambdaMethodGeneric2 = Invoke-Expression "`$lambdaMethod.MakeGenericMethod([System.Func``2[$($parentObject.GetType().FullName),System.Object]])"
            $callMethod = [System.Linq.Expressions.Expression]::Call($null, $includeMethodGeneric, $expressionArray)
             
            $expression2 = $lambdaMethodGeneric2.Invoke($null, @($callMethod, [System.Linq.Expressions.ParameterExpression[]] @($collectionParam)))
 
            $parentObject.Context.Load($parentObject, $expression2)
            if ($executeQuery) { $parentObject.Context.ExecuteQuery() }
        }
    }
    end { }
}