#http://wwwimages.adobe.com/www.adobe.com/content/dam/acom/en/devnet/indesign/sdk/cc/server/intro-to-indesign-server-cc.pdf
#https://helpx.adobe.com/indesign/using/indesign-server.html

function Set-InDesignServerComputerName {
    param (
        $ComputerName
    )
    $Script:ComputerName = $ComputerName
    # $Script:Proxy = New-WebServiceProxy -Class InDesignServer -Namespace InDesignServer -Uri (
    #     Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
    # )
}

function Get-InDesignServerComputerName {
    $Script:ComputerName
}

function Get-InDesignServerInstallPath {
    param (
        [Switch]$Remote
    )
    $ComputerName = Get-InDesignServerComputerName
    $InDesignServerInstallPath = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-ChildItem -LiteralPath "C:\Program Files\Adobe\" |
        Where-Object Name -Match "InDesign CC Server" |
        Select-Object -ExpandProperty FullName
    }
    
    if (-not $Remote) {
        $InDesignServerInstallPath
    } else {
        $InDesignServerInstallPath | ConvertTo-RemotePath -ComputerName $ComputerName
    }
}

function Get-InDesignServerService {
    $ComputerName = Get-InDesignServerComputerName
    Get-Service -ComputerName $ComputerName -Name "InDesignServerService x64"
}

function Install-InDesignServerService {
    $InDesignServerInstallPath = Get-InDesignServerInstallPath
    
    $ComputerName = Get-InDesignServerComputerName
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        & "$Using:InDesignServerInstallPath\InDesignServerService.exe" /install
    }

    Get-InDesignServerService |
    Set-Service -StartupType Automatic #-Credential $(New-Crednetial -Username system) Only is PS 6+
}

function Install-InDesignServerMMCSnapIn {
    $ComputerName = Get-InDesignServerComputerName
    $InDesignServerInstallPath = Get-InDesignServerInstallPath -ComputerName $ComputerName
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        regsvr32.exe /s $InDesignServerInstallPath\InDesignServerMMC64.dll
    }
}

function Start-InDesignServerService {
    Get-InDesignServerService | Start-Service
}

function Stop-InDesignServerService {
    Get-InDesignServerService | Stop-Service
}

function Restart-InDesignServerService {
    Get-InDesignServerService | Restart-Service -Force
}

function New-InDesignServerInstance {
    param (
        [Parameter(ValueFromPipelineByPropertyName)]$Port = 8080,
        [Parameter(ValueFromPipelineByPropertyName)]$ComputerName = (Get-InDesignServerComputerName),
        $RemoteAddress
    )
    begin {
        #This seems to be a magic string specific to each version of InDesignServer
        #Need a mapping for version numbers to magic strings and then a way to pull out version number from server to make this dynamic
        $ServiceRegistryKeyName = "InDesignCCServer2019WinService" 
        $InDesignServerServiceRegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceRegistryKeyName"

        $PortsCurrentlyUsed = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-ChildItem -Path $Using:InDesignServerServiceRegistryPath | 
            Get-ItemProperty | 
            Select-Object -ExpandProperty Port
        }
    }
    process {
        if ($Port -notin $PortsCurrentlyUsed) {
            $GUID = New-Guid | Select-Object -ExpandProperty GUID
        
            Invoke-Command -ComputerName $ComputerName -ScriptBlock {                
                $RegistryKeyPath = "$Using:InDesignServerServiceRegistryPath\$Using:GUID"
                New-Item -Path $RegistryKeyPath
                New-ItemProperty -Path $RegistryKeyPath -Name CommandLineArgs
                New-ItemProperty -Path $RegistryKeyPath -Name MaximumFailureCount -Value 10
                New-ItemProperty -Path $RegistryKeyPath -Name MaximumFailureIntervalInMinutes -Value 1440
                New-ItemProperty -Path $RegistryKeyPath -Name Port -Value $Using:Port
                New-ItemProperty -Path $RegistryKeyPath -Name TrackFailures -Value 1
            }
        }
    
        $Name = "InDesignServer instance on port $Port"
        $RemoteAddressParameter = $PSBoundParameters | ConvertFrom-PSBoundParameters -AsHashTable -Property RemoteAddress
        New-TervisFirewallRule -ComputerName $ComputerName -DisplayName $Name -Name $Name -LocalPort $Port -Direction Inbound -Action Allow -Group InDesignServer @RemoteAddressParameter
    }
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
        $Port = 8080,
        [Switch]$WithoutFix
    )
    $WSDL = Invoke-WebRequest -Uri (
        Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
    ) | 
    Select-Object -ExpandProperty Content 
    
    if ($WithoutFix) {
        $WSDL
    } else {
        $WSDL |
        Replace-ContentValue -OldValue @"
<SOAP:address location="http://localhost:$Port"/>
"@ -NewValue @"
<SOAP:address location="http://$($ComputerName):$Port"/>
"@
    }
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