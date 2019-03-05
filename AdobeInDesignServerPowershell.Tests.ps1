ipmo -force AdobeInDesignServerPowerShell

Set-InDesignServerComputerName -ComputerName INF-InDesign01.tervis.prv

while ($True) {Get-InDesignServerWSDL}


$JSXFileContent = $OrderDetail | Get-WebToPrintInDesignJSX
$ScriptText = $JSXFileContent | Replace-ContentValue -OldValue "Users/c.magnuson/test/" -NewValue ""
$ScriptText = @"
if (app.name !== 'Adobe InDesign Server') {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
}

function relinkAndEmbedByUnlinking (param) {
    var link = param.document.links.itemByName(param.itemName)
    var imageFile = File(param.imageFilePath);

    link.relink(imageFile)
    link.unlink() //This is actually embedding the linked image, really poorly named method
}

var templateFile = File("C:/ProgramData/PowerShellApplication/Production/TervisWebToPrintIllustratorServer/Templates/Print/InDesign/16oz-cstm-print.idml");
app.open(templateFile);
var document = app.documents[0];

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "Color.tif",
    imageFilePath: "C:/ThirdRun/16DWT-11157868-8-Color.png"
})

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "WhiteInk.tif",
    imageFilePath: "C:/ThirdRun/16DWT-11157868-8-WhiteInkOpacityMask.png"
})

var pdfFile = new File("C:/ThirdRun/16DWT-11157868-8.pdf");

document.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, app.pdfExportPresets.itemByName("[TervisWebToPrint]"));
document.close();
"@

ipmo -force AdobeInDesignServerPowerShell
Set-InDesignServerComputerName -ComputerName INF-InDesign01.tervis.prv

$ComputerName = "INF-InDesign01.tervis.prv"
$Proxy = New-WebServiceProxy -Class InDesignServer -Namespace InDesignServer -Uri (
    Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
)
$Proxy.Url = "http://$($ComputerName):8080/"
$Proxy.Url = "http://$($ComputerName):8082/"

$Proxy2 = New-WebServiceProxy -Class InDesignServer -Namespace InDesignServer -Uri (
    Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
)
$Proxy2.Url = "http://$($ComputerName):8080/"


$Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property @{
    ScriptText = $ScriptText
    ScriptLanguage = "javascript"
}

$Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property @{
    ScriptText = @"
$.writeln("test");
"@
    ScriptLanguage = "javascript"
}

# $Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property @{
#     ScriptText = @"
# app.pdfExportPresets.itemByName("[TervisWebToPrint]").name
# "@
#     ScriptLanguage = "javascript"
# }


$ErrorString = ""
$Results = New-Object -TypeName InDesignServer.Data
$SessionID = $Proxy.BeginSessionAsync()
$SessionID = $Proxy.BeginSession()
$Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
$Response
$ErrorString
$Results
$proxy.EndSession($SessionID)

Invoke-InDesignServerRunScript -ScriptText $ScriptText



$ComputerName = "INF-InDesign01.tervis.prv"
$Proxy = New-WebServiceProxy -Class InDesignServer -Namespace InDesignServer -Uri (
    Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port 8080
)
$Proxy.Url = "http://$($ComputerName):8080/"

$Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property @{
    ScriptText = @"
$.writeln("test");
"@
    ScriptLanguage = "javascript"
}

$ErrorString = ""
$Results = New-Object -TypeName InDesignServer.Data

$SessionID = $Proxy.BeginSession()
$Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
$Response
$ErrorString
$Results
$proxy.EndSession($SessionID)




function Test-InDesignMultiInstancePerf {
    Set-InDesignServerComputerName -ComputerName INF-InDesign02.tervis.prv
    $ComputerName = "INF-InDesign02.tervis.prv"
    $Port = 8083

    $Proxy = New-WebServiceProxy -Class "InDesignServer$Port" -Namespace "InDesignServer$Port" -Uri (
        Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port $Port
    )
    $Proxy.Url = "http://$($ComputerName):$Port/"
    $Proxy

    $Parameter = New-Object -TypeName "InDesignServer$Port.RunScriptParameters" -Property @{
        ScriptText = @"
"test"
"@
        ScriptLanguage = "javascript"
    }

    $ErrorString = ""
    $Results = New-Object -TypeName "InDesignServer$Port.Data"

    $Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
    $Response
    $ErrorString
    $Results.data


    Measure-command {
        1..10000 | % {
            $Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
        }
    }
}

function Test-NewInDesignServerInstance {
    Set-InDesignServerComputerName -ComputerName INF-InDesign02.tervis.prv
    $ComputerName = "INF-InDesign02.tervis.prv"
    $Port = 8081

    $Proxy = New-WebServiceProxy -Class "InDesignServer$Port" -Namespace "InDesignServer$Port" -Uri (
        Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port $Port
    )
    $Proxy.Url = "http://$($ComputerName):$Port/"
    $Proxy

    $Parameter = New-Object -TypeName "InDesignServer$Port.RunScriptParameters" -Property @{
        ScriptText = @"
"test"
"@
        ScriptLanguage = "javascript"
    }

    $ErrorString = ""
    $Results = New-Object -TypeName "InDesignServer$Port.Data"

    $SessionID = $Proxy.BeginSession()
    $SessionID
    $Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
    $Response
    $ErrorString
    $Results.data
    $proxy.EndSession($SessionID)

    Measure-command {
        1..10000 | % {
            $SessionID = $Proxy.BeginSession()
            $Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
            $proxy.EndSession($SessionID)
        }
    }

    Start-RSJob -ScriptBlock {
        $ComputerName = "INF-InDesign02.tervis.prv"
        $Port = 8081

        $Proxy = New-WebServiceProxy -Class "InDesignServer$Port" -Namespace "InDesignServer$Port" -Uri (
            Get-InDesignServerWSDLURI -ComputerName $ComputerName -Port $Port
        )
        $Proxy.Url = "http://$($ComputerName):$Port/"
        $Parameter = New-Object -TypeName "InDesignServer$Port.RunScriptParameters" -Property @{
            ScriptText = @"
"test"
"@
            ScriptLanguage = "javascript"
        }

        $ErrorString = ""
        $Results = New-Object -TypeName "InDesignServer$Port.Data"
        $SessionID = $Proxy.BeginSession()
        1..1000 | % {
            $Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [ref]$Results)
        }
        $proxy.EndSession($SessionID)
    }
}

$ScriptText = @"
$.writeln("test");
"@
$InDesignServerInstance = Select-InDesignServerInstance -SelectionMethod Random
$Result = Invoke-InDesignServerRunScriptDirectlyWithInlineSOAP -ScriptText $ScriptText -InDesignServerInstance $InDesignServerInstance


$Body = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
        <RunScript xmlns="http://ns.adobe.com/InDesign/soap/">
            <runScriptParameters xmlns="">
                <scriptText>$.writeln("test")
$.writeln("test")
</scriptText>
                <scriptLanguage>javascript</scriptLanguage>
                <scriptFile xsi:nil="true" />
            </runScriptParameters>
        </RunScript>
    </soap:Body>
</soap:Envelope>
"@

$Body = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
        <RunScript xmlns="http://ns.adobe.com/InDesign/soap/">
            <runScriptParameters xmlns="">
                <scriptText>app.serverSettings.sessionTimeout</scriptText>
                <scriptLanguage>javascript</scriptLanguage>
                <scriptFile xsi:nil="true" />
            </runScriptParameters>
        </RunScript>
    </soap:Body>
</soap:Envelope>
"@


$Body = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
        <RunScript xmlns="http://ns.adobe.com/InDesign/soap/">
            <runScriptParameters xmlns="">
                <scriptText>if (app.name !== 'Adobe InDesign Server') {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
}

function relinkAndEmbedByUnlinking (param) {
    var link = param.document.links.itemByName(param.itemName)
    while (link !== null) {
        var imageFile = File(param.imageFilePath)

        link.relink(imageFile)
        link.unlink() //This is actually embedding the linked image, really poorly named method
        link = param.document.links.itemByName(param.itemName)
    }
}

var templateFile = File("C:/ProgramData/PowerShellApplication/Production/TervisWebToPrintTemplates/Templates/Print/InDesign/6SIP-cstm-print.idml");
app.open(templateFile);
var document = app.documents[0];

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "Color.png",
    imageFilePath: "C:/windows/temp/a0162000-046f-49be-992c-b00a68ac6c4e/ColorImage.png"
})

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "WhiteInk.png",
    imageFilePath: "C:/windows/temp/a0162000-046f-49be-992c-b00a68ac6c4e/WhiteInkImage.png"
})

if ("") {
    relinkAndEmbedByUnlinking({
        document: document,
        itemName: "ColorVuMark.png",
        imageFilePath: ""
    })
} else {
    var objectsList = app.findObject();
    for (var i = 0; i < objectsList.length; i++) {
        if (
            objectsList[i].contentType === ContentType.GRAPHIC_TYPE &&
            objectsList[i].images.count &&
            objectsList[i].images[0].itemLink.name === "ColorVuMark.png"
        ) {
            objectsList[i].remove()
        }
    }
}

if ("") {
    var textFrames = document.pages[0].textFrames
    textFrames[0].contents = ""
    textFrames[1].contents = ""
}

var pdfFile = new File("C:/windows/temp/a0162000-046f-49be-992c-b00a68ac6c4e/PDFFile.pdf")

if (app.name !== 'Adobe InDesign Server') {
    document.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, app.pdfExportPresets.itemByName("TervisWebToPrint"))
} else {
    document.exportFile(ExportFormat.PDF_TYPE, pdfFile, app.pdfExportPresets.itemByName("[TervisWebToPrint]"))
}

document.close()</scriptText>
                <scriptLanguage>javascript</scriptLanguage>
                <scriptFile xsi:nil="true" />
            </runScriptParameters>
        </RunScript>
    </soap:Body>
</soap:Envelope>
"@



$Body = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
        <RunScript xmlns="http://ns.adobe.com/InDesign/soap/">
            <runScriptParameters xmlns="">
                <scriptText>
                $ScriptText
                </scriptText>
                <scriptLanguage>javascript</scriptLanguage>
                <scriptFile xsi:nil="true" />
            </runScriptParameters>
        </RunScript>
    </soap:Body>
</soap:Envelope>
"@

$Body = @"
<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"><soap:Body><RunScript xmlns="http://ns.adobe.com/InDesign/soap/"><runScriptParameters xmlns=""><scriptText>if (app.name !== 'Adobe InDesign Server') {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
}

function relinkAndEmbedByUnlinking (param) {
    var link = param.document.links.itemByName(param.itemName)
    while (link !== null) {
        var imageFile = File(param.imageFilePath)
    
        link.relink(imageFile)
        link.unlink() //This is actually embedding the linked image, really poorly named method
        link = param.document.links.itemByName(param.itemName)
    }
}

var templateFile = File("C:/ProgramData/PowerShellApplication/Production/TervisWebToPrintTemplates/Templates/Print/InDesign/10WAV-cstm-print.idml");
app.open(templateFile);
var document = app.documents[0];

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "Color.png",
    imageFilePath: "C:/windows/temp/9bc25487-c51b-49f8-883c-a5062e8b5b5c/ColorImage.png"
})

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "WhiteInk.png",
    imageFilePath: "C:/windows/temp/9bc25487-c51b-49f8-883c-a5062e8b5b5c/WhiteInkImage.png"
})

if ("") {
    relinkAndEmbedByUnlinking({
        document: document,
        itemName: "ColorVuMark.png",
        imageFilePath: ""
    })
} else {
    var objectsList = app.findObject();
    for (var i = 0; i &lt; objectsList.length; i++) {
        if (
            objectsList[i].contentType === ContentType.GRAPHIC_TYPE &amp;&amp;
            objectsList[i].images.count &amp;&amp; 
            objectsList[i].images[0].itemLink.name === "ColorVuMark.png"
        ) {
            objectsList[i].remove()
        }
    }
}

if ("11424551/1") {
    var textFrames = document.pages[0].textFrames
    textFrames[0].contents = "11424551/1"
    textFrames[1].contents = "11424551/1"
}

var pdfFile = new File("C:/windows/temp/9bc25487-c51b-49f8-883c-a5062e8b5b5c/PDFFile.pdf")

if (app.name !== 'Adobe InDesign Server') {
    document.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, app.pdfExportPresets.itemByName("[TervisWebToPrint]"))
} else {
    document.exportFile(ExportFormat.PDF_TYPE, pdfFile, app.pdfExportPresets.itemByName("[TervisWebToPrint]"))
}

document.close()</scriptText><scriptLanguage>javascript</scriptLanguage><scriptFile xsi:nil="true" /></runScriptParameters></RunScript></soap:Body></soap:Envelope>
"@

$Body = @"
<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"><soap:Body><RunScript xmlns="http://ns.adobe.com/InDesign/soap/"><runScriptParameters xmlns=""><scriptText>if (app.name !== 'Adobe InDesign Server') {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
}

function relinkAndEmbedByUnlinking (param) {
    var link = param.document.links.itemByName(param.itemName)
    while (link !== null) {
        var imageFile = File(param.imageFilePath)

        link.relink(imageFile)
        link.unlink() //This is actually embedding the linked image, really poorly named method
        link = param.document.links.itemByName(param.itemName)
    }
}

var templateFile = File("C:/ProgramData/PowerShellApplication/Production/TervisWebToPrintTemplates/Templates/Print/InDesign/6SIP-cstm-print.idml");
app.open(templateFile);
var document = app.documents[0];

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "Color.png",
    imageFilePath: "C:/windows/temp/a0162000-046f-49be-992c-b00a68ac6c4e/ColorImage.png"
})

relinkAndEmbedByUnlinking({
    document: document,
    itemName: "WhiteInk.png",
    imageFilePath: "C:/windows/temp/a0162000-046f-49be-992c-b00a68ac6c4e/WhiteInkImage.png"
})

if ("") {
    relinkAndEmbedByUnlinking({
        document: document,
        itemName: "ColorVuMark.png",
        imageFilePath: ""
    })
}

if ("") {
    var textFrames = document.pages[0].textFrames
    textFrames[0].contents = ""
    textFrames[1].contents = ""
}

var pdfFile = new File("C:/windows/temp/a0162000-046f-49be-992c-b00a68ac6c4e/PDFFile.pdf")

if (app.name !== 'Adobe InDesign Server') {
    document.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, app.pdfExportPresets.itemByName("TervisWebToPrint"))
} else {
    document.exportFile(ExportFormat.PDF_TYPE, pdfFile, app.pdfExportPresets.itemByName("[TervisWebToPrint]"))
}

document.close()</scriptText><scriptLanguage>javascript</scriptLanguage><scriptFile xsi:nil="true" /></runScriptParameters></RunScript></soap:Body></soap:Envelope>
"@