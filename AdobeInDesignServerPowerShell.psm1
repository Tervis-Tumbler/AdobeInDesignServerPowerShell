#http://wwwimages.adobe.com/www.adobe.com/content/dam/acom/en/devnet/indesign/sdk/cc/server/intro-to-indesign-server-cc.pdf
#https://helpx.adobe.com/indesign/using/indesign-server.html

function Set-InDesignServerComputerName {
    param (
        $ComputerName
    )
    $Script:ComputerName = $ComputerName
    $Script:Proxy = New-WebServiceProxy -Class InDesignServer -Namespace InDesignServer -Uri (
        Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
    )
}

function Get-InDesignServerComputerName {
    $Script:ComputerName
}

function Get-InDesignServerInstallPath {
    param (
        [Switch]$Remote
    )
    $LocalPath = "C:\Program Files\Adobe\Adobe InDesign CC Server 2018"
    if (-not $Remote) {
        $LocalPath
    } else {
        $ComputerName = Get-InDesignServerComputerName
        $LocalPath | ConvertTo-RemotePath -ComputerName $ComputerName
    }  
}

function Get-InDesignServerService {
    $ComputerName = Get-InDesignServerComputerName
    Get-Service -ComputerName $ComputerName -Name "InDesignServerService x64"
}

function Install-InDesignServerService {
    param (

    )
    .\InDesignServerService.exe /install
    Get-InDesignServerService | 
    Set-Service -StartupType Automatic #-Credential $(New-Crednetial -Username system) Only is PS 6+
}

function Install-InDesignServerMMCSnapIn {
    regsvr32.exe .\InDesignServerMMC64.dll
}

function Start-InDesignServerService {
    Get-InDesignServerService | Start-Service
}

function Restart-InDesignServerService {
    Get-InDesignServerService | Restart-Service -Force
}

function New-InDesignServerInstance {
    param (
        $Port
    )
    $GUID = New-Guid | Select-Object -ExpandProperty GUID
    New-Item -Path HKLM:\SYSTEM\CurrentControlSet\Services\InDesignCCServer2017WinService -Name $Guid
    $ComputerName = Get-InDesignServerComputerName
    $Name = "InDesignServer $Port $GUID"
    New-TervisFirewallRule -ComputerName $ComputerName -DisplayName $Name -Name $Name -LocalPort $Port -Direction Inbound -Action Allow -Group InDesignServer
}

function Get-InDesignServerWSDLURI {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$Port
    )
    process {
        "http://$($ComputerName):$Port/service?wsdl"
    }
}

function Get-InDesignServerWSDL {
    param (
        $ComputerName = (Get-InDesignServerComputerName),
        $Port = 8080
    )
    Invoke-WebRequest -Uri (
        Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
    ) | 
    Select-Object -ExpandProperty Content |
    Replace-ContentValue -OldValue @"
<SOAP:address location="http://localhost:$Port"/>
"@ -NewValue @"
<SOAP:address location="http://$($ComputerName):$Port"/>
"@
}

function Get-InDesignServerWebServiceProxy {
    if ($Script:Proxy) {
        $Script:Proxy
    } else {
        Throw "No proxy object found, Set-InDesignServerComputerName needs to be called first"
    }
}

function Invoke-InDesignServerAPI {
    param (
        $MethodName,
        $Parameter,
        $Property
    )
    $Proxy = Get-InDesignServerWebServiceProxy

    if (-not $Parameter) {
        if ($Property) {
            $Parameter = New-Object -TypeName InDesignServer."$($MethodName)Parameters" -Property $Property
        } else {
            $Parameter = New-Object -TypeName InDesignServer."$($MethodName)Parameters"
        }
    }
    $Response = $Proxy.$MethodName($Parameter)
    $Response.result
}

function Invoke-InDesignServerRunScript {
    param (
        $ScriptText,
        $ScriptLanguage,
        $ScriptFile,
        $ScriptArgs
    )    
    #Invoke-InDesignServerAPI -MethodName RunScript -Property $PSBoundParameters

    $Proxy = Get-InDesignServerWebServiceProxy
    $Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property $PSBoundParameters
    $ErrorString = ""
    $Results = New-Object -TypeName InDesignServer.Data

    $Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
    $Response.result
}

function Invoke-InDesignServerJSX {
    param (
        $ScriptConent
    )
    Invoke-InDesignServerRunScript -ScriptText $ScriptConent -ScriptLanguage "JavaScript"
}

function Set-InDesingServerJobOption {
    param (
        $LocalPathToJobOptions
    )
    $LocalPathToJobOptions
    $PathToJobOptionsFolder = "C:\Program Files\Adobe\Adobe InDesign CC Server 2018\Resources\Adobe PDF\settings\mul"
    $InDesignServerComputerName = Get-InDesignServerComputerName
    $PathToJobOptionsFolderRemote = $PathToJobOptionsFolder | ConvertTo-RemotePath -ComputerName $InDesignServerComputerName
    Copy-Item -Path $LocalPathToJobOptions -Destination $PathToJobOptionsFolderRemote
}