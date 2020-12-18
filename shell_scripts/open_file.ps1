. .\shell_scripts\notifier.ps1

# Execute the macro
function Invoke-Macro {
    Param (
        [__ComObject]
        $excel,
        [string]
        $macro,
        [Boolean]
        $notify,
        [string]
        $file_label,
        [string]
        $notifier,
        [string]
        $notification="{0} executed."
    )

    $excel.Run($macro)
    
    if ($notify) { 
        $message = ($notification + " [" + (Get-Date).ToString('T') + "]") -f $macro
        Show-Notification $notifier $file_label $message
        Start-Sleep 1
    }
}

# Open the excel file
function Open-File {
    Param (
        [System.Xml.XmlLinkedNode]
        $config,
        [System.Xml.XmlLinkedNode]
        $file
    )

    $excel = New-Object -comobject Excel.Application

    # Open excel sheet
    $workbook = $excel.Workbooks.Open($file.path)

    $excel.Visible = if ($config.FileVisible -eq "True") { $true } else { $false }

    # Trigger the macros
    $file.Macros.Macro | % { Invoke-Macro $excel $_ ($config.NotifyFor -eq "Macros") $file.Label $config.Notifier $config.NotificationMessage }
    
    $excel.Workbooks.Close()
    $excel.Quit()

    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Start-Sleep 2

    if ($config.NotifyFor -eq "Files") { 
        $message = $config.NotificationMessage + " [" + (Get-Date).ToString('T') + "]"
        Show-Notification $config.Notifier $file.Label $message 
    }
}