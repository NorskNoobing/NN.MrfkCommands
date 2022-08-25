function Send-TeamsMessage {
    param (
        [Parameter(Mandatory)]$Content,
        [Parameter(Mandatory)]$HookURI
    )

    $InvokeRestMethodSplat = @{
        Body        = [PSCustomObject]@{
            text = $Content
        } | ConvertTo-Json
            
        ContentType = 'Application/Json'
        Method      = 'POST'
        Uri         = $HookURI
    }
    Invoke-RestMethod @InvokeRestMethodSplat
}