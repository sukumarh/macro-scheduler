$path = $MyInvocation.MyCommand.Path
$dir = Split-Path $path
Push-Location $dir

# Load configuration
[xml]$config = Get-Content .\config.xml

. .\shell_scripts\open_file.ps1

# Open each file in the configuration
$config.Configuration.Files.File | % { Open-File $config.Configuration $_ }

Pop-Location