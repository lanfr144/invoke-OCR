function Expand-Template {
    param([string]$Template, [hashtable]$Variables)
    $result = $Template
    foreach ($key in $Variables.Keys) {
        $result = $result -replace [regex]::Escape("`${$key}"), $Variables[$key]
    }
    return $result
}

Export-ModuleMember -Function Expand-Template
