function Show-ErrorPopup {
    param([string]$Title, [string]$Message)
    try {
        # Try BurntToast module first (best experience)
        if (Get-Module -ListAvailable -Name BurntToast -ErrorAction SilentlyContinue) {
            Import-Module BurntToast -ErrorAction SilentlyContinue
            New-BurntToastNotification -Text $Title, $Message -AppLogo $null -ErrorAction Stop
            return
        }

        # Fallback: native .NET toast (Windows 10/11, requires loaded assemblies)
        $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] 2>$null
        if ($?) {
            $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
            $textNodes = $template.GetElementsByTagName("text")
            $textNodes.Item(0).AppendChild($template.CreateTextNode($Title)) | Out-Null
            $textNodes.Item(1).AppendChild($template.CreateTextNode($Message)) | Out-Null
            $toast = [Windows.UI.Notifications.ToastNotification]::new($template)
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Invoke-OCR").Show($toast)
            return
        }
    }
    catch { }

    # Last resort: msg.exe (works on Windows Pro/Enterprise)
    try {
        Start-Process "msg.exe" -ArgumentList "* `"${Title}: $Message`"" -WindowStyle Hidden -ErrorAction SilentlyContinue
    }
    catch { }
}

Export-ModuleMember -Function Show-ErrorPopup
