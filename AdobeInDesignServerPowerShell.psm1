#http://wwwimages.adobe.com/www.adobe.com/content/dam/acom/en/devnet/indesign/sdk/cc/server/intro-to-indesign-server-cc.pdf
#https://helpx.adobe.com/indesign/using/indesign-server.html

function Set-InDesignServerComputerName {
    param (
        $ComputerName
    )
    $Script:ComputerName = $ComputerName
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
    Invoke-WebRequest -Uri (
        Get-InDesignServerWSDLURI -ComputerName (
            Get-InDesignServerComputerName
        ) -Port 8080
    )
}

function Invoke-ProgisticsAPI {
    param (
        $MethodName,
        $Parameter,
        $Property
    )

    $Proxy = New-WebServiceProxy -Uri (
        Get-InDesignServerWSDLURI -ComputerName (
            Get-InDesignServerComputerName
        ) -Port 8080
    ) -Class InDesignServer -Namespace InDesignServer

    if (-not $Parameter) {
        if ($Property) {
            $Parameter = New-Object -TypeName Progistics."$($MethodName)Request" -Property $Property
        } else {
            $Parameter = New-Object -TypeName Progistics."$($MethodName)Request"
        }
    }
    $Response = $Proxy.$MethodName($Parameter)
    $Response.result
}

function Invoke-InDesignServerRunScript {
    param (
        $ScriptConent
    )

    $Proxy
    
    $RunScriptParameters = New-Object -TypeName InDesignServer.RunScriptParameters
    $Proxy.RunScript($RunScriptParameters)
}