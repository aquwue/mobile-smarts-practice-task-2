param(
    [string]$SourceDirectory = $PSScriptRoot,
    [string]$DatabaseDirectory = 'C:\ProgramData\Cleverence\Platform\Databases\Empty app',
    [string[]]$DocumentNumbers = @('006', '007', '008')
)

$ErrorActionPreference = 'Stop'
$utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
$documentsDirectory = Join-Path $DatabaseDirectory 'Data\Documents'

foreach ($number in $DocumentNumbers) {
    $sourcePath = Join-Path $SourceDirectory "day2-interface-check-$number.json"
    $sourceText = [System.IO.File]::ReadAllText($sourcePath, [System.Text.Encoding]::UTF8)
    $sourceDocument = ($sourceText | ConvertFrom-Json)[0]
    $targetPath = Join-Path $documentsDirectory "doc_$($sourceDocument.id).xml"

    if (Test-Path -LiteralPath $targetPath) {
        throw "Document already exists: $($sourceDocument.id)"
    }

    $now = (Get-Date).ToString('o')
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = $utf8WithoutBom
    $settings.Indent = $true
    $settings.OmitXmlDeclaration = $false

    $writer = [System.Xml.XmlWriter]::Create($targetPath, $settings)
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('Document')
        $writer.WriteAttributeString('xmlns', 'clr', $null, 'http://schemas.cleverence.ru/clr')
        $writer.WriteAttributeString('barcode', [string]$sourceDocument.barcode)
        $writer.WriteAttributeString('createDate', $now)
        $writer.WriteAttributeString('lastChangeDate', $now)
        $writer.WriteAttributeString('description', [string]$sourceDocument.description)
        $writer.WriteAttributeString('deviceId', '')
        $writer.WriteAttributeString('deviceIP', '')
        $writer.WriteAttributeString('deviceName', '')
        $writer.WriteAttributeString('documentTypeName', [string]$sourceDocument.documentTypeName)
        $writer.WriteAttributeString('id', [string]$sourceDocument.id)
        $writer.WriteAttributeString('inProcess', 'True')
        $writer.WriteAttributeString('name', [string]$sourceDocument.name)
        $writer.WriteAttributeString('priority', [string]$sourceDocument.priority)
        $writer.WriteAttributeString('appointment', [string]$sourceDocument.appointment)
        $writer.WriteAttributeString('userId', [string]$sourceDocument.appointment)
        $writer.WriteAttributeString('userName', [string]$sourceDocument.appointment)

        $writer.WriteStartElement('Fields')
        $writer.WriteEndElement()

        $writer.WriteStartElement('States')
        $writer.WriteStartElement('DocumentState')
        $writer.WriteAttributeString('finished', 'False')
        $writer.WriteAttributeString('finishedDate', '0001-01-01T00:00:00.0000000')
        $writer.WriteAttributeString('inProcess', 'True')
        $writer.WriteAttributeString('inProcessDate', $now)
        $writer.WriteAttributeString('modified', 'False')
        $writer.WriteAttributeString('modifiedDate', '0001-01-01T00:00:00.0000000')
        $writer.WriteAttributeString('processingTime', '00:00:00')
        $writer.WriteAttributeString('userId', [string]$sourceDocument.appointment)
        $writer.WriteEndElement()
        $writer.WriteEndElement()

        $writer.WriteStartElement('Tables')
        $writer.WriteEndElement()

        $writer.WriteStartElement('ExpectedLines')
        foreach ($line in $sourceDocument.declaredItems) {
            $writer.WriteStartElement('DocumentItem')
            $writer.WriteAttributeString('uid', [string]$line.uid)
            $writer.WriteAttributeString('inventoryItemId', [string]$line.productId)
            $writer.WriteAttributeString('unitOfMeasureId', [string]$line.packingId)
            $writer.WriteAttributeString('actualQuantity', '0')
            $writer.WriteAttributeString('expectedQuantity', [string]$line.declaredQuantity)
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('ActualLines')
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }

    Write-Output "Created $($sourceDocument.id)"
}
