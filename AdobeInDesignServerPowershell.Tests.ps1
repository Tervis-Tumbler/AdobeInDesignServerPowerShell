ipmo -force AdobeInDesignServerPowerShell

Set-InDesignServerComputerName -ComputerName INF-InDesign01.tervis.prv
Invoke-InDesignServerRunScript -ScriptText "test"

Get-InDesignServerWSDL


Invoke-WebRequest -Uri (
    Get-InDesignServerWSDLURI -ComputerName inf-InDesign01 -Port 8080
)