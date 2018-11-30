ipmo -force AdobeInDesignServerPowerShell

Set-InDesignServerComputerName -ComputerName INF-InDesign01.tervis.prv
Invoke-InDesignServerRunScript -ScriptText "test"

Get-InDesignServerWSDL

$ScriptText = @"
if (app.name !== 'Adobe InDesign Server') {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
}

function relinkAndEmbedByUnlinking (param) {
    var document = app.documents[0];
    var link = document.links.itemByName(param.itemName)
    var imageFile = File(param.imageFilePath);

    link.relink(imageFile)
    link.unlink() //This is actually embedding the linked image, really poorly named method
}

var templateFile = File("C:/ProgramData/PowerShellApplication/Production/TervisWebToPrintIllustratorServer/Templates/Print/InDesign/16oz-cstm-print.idml");
app.open(templateFile);

relinkAndEmbedByUnlinking({
    itemName: "Color.tif",
    imageFilePath: "C:/ThirdRun/16DWT-11157868-8-Color.png"
})

relinkAndEmbedByUnlinking({
    itemName: "WhiteInk.tif",
    imageFilePath: "C:/ThirdRun/16DWT-11157868-8-WhiteInkOpacityMask.png"
})

var pdfFile = new File("C:/ThirdRun/16DWT-11157868-8.pdf");

document.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, app.pdfExportPresets.itemByName("TervisWebToPrint"));
document.close();
"@

ipmo -force AdobeInDesignServerPowerShell
Set-InDesignServerComputerName -ComputerName INF-InDesign01.tervis.prv

$ComputerName = "INF-InDesign01.tervis.prv"
$Proxy = New-WebServiceProxy -Class InDesignServer -Namespace InDesignServer -Uri (
    Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
)
$Proxy.Url = "http://$($ComputerName):8080/"

$Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property @{
    ScriptText = $ScriptText
    ScriptLanguage = "javascript"
}


$Response.result$Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property @{
    ScriptText = @"
app.pdfExportPresets.itemByName("[TervisWebToPrint]").name;
"@
    ScriptLanguage = "javascript"
}

$ErrorString = ""
$Results = New-Object -TypeName InDesignServer.Data

$Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)

Invoke-InDesignServerRunScript -ScriptText $ScriptText