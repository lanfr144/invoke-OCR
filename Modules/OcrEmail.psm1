function Expand-Template {
    param([string]$Template, [hashtable]$Variables)
    Write-Verbose "Entering $($MyInvocation.MyCommand.Name)"
    Write-Debug "Entering $($MyInvocation.MyCommand.Name) - Parameters: $($PSBoundParameters | Out-String)"
    Write-Information "Entering $($MyInvocation.MyCommand.Name)" -InformationAction Continue
    $result = $Template
    foreach ($key in $Variables.Keys) {
        $result = $result -replace [regex]::Escape("`${$key}"), $Variables[$key]
    }
    return $result
}

Export-ModuleMember -Function Expand-Template
