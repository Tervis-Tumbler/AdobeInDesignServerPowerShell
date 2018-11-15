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