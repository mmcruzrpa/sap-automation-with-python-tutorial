$exclude = @("venv", "sap_automation.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "sap_automation.zip" -Force