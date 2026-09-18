#requires -Version 5.1
<#
.SYNOPSIS
Exports Swagger 2.0 JSON request and response structures to a real .xlsx workbook.
.EXAMPLE
.\Export-SwaggerToExcel.ps1 -SwaggerPath 'C:\Users\naray\Downloads\swagger.json' -OutputPath '.\SwaggerFields.xlsx'
.NOTES
Uses built-in .NET ZIP/XML APIs. Excel and additional PowerShell modules are not required.
Required means required within the immediate containing object, when that object is present.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SwaggerPath,
    [string]$OutputPath = '.\SwaggerFields.xlsx',
    [switch]$ExcludeRequestHeaders,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'

function Get-Value($Object, [string]$Name) {
    if ($null -ne $Object) {
        $property = $Object.PSObject.Properties[$Name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}
function Resolve-Reference($Node, [string[]]$Chain = @()) {
    $reference = Get-Value $Node '$ref'
    if (-not $reference) { return $Node }
    if ($Chain -ccontains $reference) { throw "Circular reference alias: $reference" }
    if (-not $reference.StartsWith('#/')) { throw "External references are not supported: $reference" }
    $target = $script:Document
    foreach ($part in $reference.Substring(2).Split('/')) {
        $key = [Uri]::UnescapeDataString($part).Replace('~1', '/').Replace('~0', '~')
        $target = Get-Value $target $key
        if ($null -eq $target) { throw "Unresolved reference: $reference" }
    }
    Resolve-Reference $target ($Chain + $reference)
}
function Add-Row($Rows, [string]$Field, [string]$Type = '', [string]$Required = '', [int]$Style = 0) {
    $Rows.Add([pscustomobject]@{ Values = @($Field, $Type, $Required); Style = $Style })
}
function Write-Schema($Rows, $Schema, [string]$Name, [string]$Required = '', [int]$Depth = 0, [string[]]$Stack = @()) {
    if ($Depth -gt 60) { throw "Schema exceeds the supported nesting depth at $Name" }
    $prefix = '  ' * $Depth
    $reference = Get-Value $Schema '$ref'
    if ($reference -and ($Stack -ccontains $reference)) {
        Add-Row $Rows "$prefix$Name" "recursive reference: $reference" $Required
        return
    }
    if ($reference) { $Stack = $Stack + $reference }
    $node = Resolve-Reference $Schema
    foreach ($keyword in @('allOf', 'oneOf', 'anyOf')) {
        if (Get-Value $node $keyword) { throw "Schema composition '$keyword' is not supported at $Name. Flatten the schema first." }
    }
    $type = Get-Value $node 'type'
    $properties = Get-Value $node 'properties'
    if (-not $type) {
        if ($properties -or (Get-Value $node 'additionalProperties')) { $type = 'object' }
        else { $type = 'any' }
    }
    $format = Get-Value $node 'format'
    $displayType = $type
    if ($format) { $displayType = "$type ($format)" }
    if ($type -eq 'object') {
        Add-Row $Rows "$prefix<<$Name>>" 'object' $Required 2
        if ($properties) {
            $requiredNames = @(Get-Value $node 'required')
            foreach ($property in $properties.PSObject.Properties) {
                if ((Get-Value $property.Value 'readOnly') -eq $true -and $script:Direction -eq 'request') { continue }
                $status = 'Optional'
                if ($requiredNames -ccontains $property.Name) { $status = 'Required' }
                Write-Schema $Rows $property.Value $property.Name $status ($Depth + 1) $Stack
            }
        }
        $additional = Get-Value $node 'additionalProperties'
        if ($additional -is [pscustomobject]) {
            Write-Schema $Rows $additional '{additional key}' 'Optional' ($Depth + 1) $Stack
        } elseif ($additional -eq $true) {
            Add-Row $Rows "$prefix  {additional key}" 'any' 'Optional'
        }
        Add-Row $Rows "$prefix<</$Name>>" '' '' 2
    } elseif ($type -eq 'array') {
        Add-Row $Rows "$prefix<<$Name>>" 'array' $Required 2
        $items = Get-Value $node 'items'
        if ($null -eq $items) { throw "Array is missing items at $Name" }
        Write-Schema $Rows $items "$Name[]" '' ($Depth + 1) $Stack
        Add-Row $Rows "$prefix<</$Name>>" '' '' 2
    } else {
        Add-Row $Rows "$prefix$Name" $displayType $Required
    }
}
function Get-SchemaName($Schema, [string]$Fallback) {
    $reference = Get-Value $Schema '$ref'
    if ($reference) { return [Uri]::UnescapeDataString(($reference -split '/')[-1]).Replace('~1', '/').Replace('~0', '~') }
    return $Fallback
}
function Escape-Xml([string]$Text) { [System.Security.SecurityElement]::Escape($Text) }
function Add-ZipText($Zip, [string]$Name, [string]$Text) {
    $entry = $Zip.CreateEntry($Name)
    $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}
function Get-SheetXml($Rows) {
    if ($Rows.Count -gt 1048576) { throw 'Worksheet exceeds the Excel row limit.' }
    $xml = [System.Text.StringBuilder]::new()
    [void]$xml.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/><cols><col min="1" max="1" width="70" customWidth="1"/><col min="2" max="2" width="28" customWidth="1"/><col min="3" max="3" width="23" customWidth="1"/></cols><sheetData>')
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $r = $i + 1
        $row = $Rows[$i]
        $height = 22
        foreach ($value in $row.Values) {
            if ($value.Length -gt 65) { $height = [Math]::Max($height, 16 * [Math]::Ceiling($value.Length / 65)) }
        }
        [void]$xml.Append("<row r=`"$r`" ht=`"$height`" customHeight=`"1`">")
        for ($c = 0; $c -lt 3; $c++) {
            $address = ([string][char](65 + $c)) + $r
            $value = Escape-Xml $row.Values[$c]
            [void]$xml.Append("<c r=`"$address`" s=`"$($row.Style)`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$value</t></is></c>")
        }
        [void]$xml.Append('</row>')
    }
    [void]$xml.Append('</sheetData></worksheet>')
    $xml.ToString()
}

$inputFile = (Resolve-Path -LiteralPath $SwaggerPath).ProviderPath
$outputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([System.IO.Path]::GetExtension($outputFile) -ne '.xlsx') { throw 'OutputPath must end with .xlsx.' }
if ($inputFile -eq $outputFile) { throw 'Input and output paths must differ.' }
if ((Test-Path -LiteralPath $outputFile) -and -not $Force) { throw "Output already exists. Use -Force to replace it: $outputFile" }
$script:Document = Get-Content -LiteralPath $inputFile -Raw -Encoding UTF8 | ConvertFrom-Json
if ((Get-Value $script:Document 'swagger') -ne '2.0') { throw 'This script supports Swagger 2.0 JSON. OpenAPI 3.x and YAML are not supported.' }
$requests = [System.Collections.Generic.List[object]]::new()
$responses = [System.Collections.Generic.List[object]]::new()
Add-Row $requests 'Source Fields' 'Fields DataType' 'Required/Optional' 1
Add-Row $responses 'Source Fields' 'Fields DataType' 'Required/Optional' 1
$operationCount = 0
foreach ($path in $script:Document.paths.PSObject.Properties) {
    $pathItem = Resolve-Reference $path.Value
    foreach ($method in @('get', 'put', 'post', 'delete', 'options', 'head', 'patch')) {
        $operation = Get-Value $pathItem $method
        if (-not $operation) { continue }
        $operationCount++
        $label = $method.ToUpperInvariant() + ' ' + $path.Name
        Add-Row $requests $label '' '' 3
        Add-Row $responses $label '' '' 3
        $script:Direction = 'request'
        # An operation parameter overrides a path parameter with the same name and location.
        $parameters = [System.Collections.Specialized.OrderedDictionary]::new([System.StringComparer]::Ordinal)
        foreach ($parameter in (@(Get-Value $pathItem 'parameters') + @(Get-Value $operation 'parameters'))) {
            if ($null -eq $parameter) { continue }
            $p = Resolve-Reference $parameter
            $parameters[($p.in + ':' + $p.name)] = $p
        }
        foreach ($p in $parameters.Values) {
            if ($ExcludeRequestHeaders -and $p.in -eq 'header') { continue }
            $status = 'Optional'
            if ($p.required -eq $true -or $p.in -eq 'path') { $status = 'Required' }
            if ($p.in -eq 'body') {
                Write-Schema $requests $p.schema (Get-SchemaName $p.schema $p.name) $status
            } else {
                Write-Schema $requests $p ("$($p.name) [$($p.in)]") $status
            }
        }
        $script:Direction = 'response'
        foreach ($response in $operation.responses.PSObject.Properties) {
            $r = Resolve-Reference $response.Value
            Add-Row $responses ("HTTP $($response.Name): $($r.description)") '' '' 3
            $schema = Get-Value $r 'schema'
            if ($schema) { Write-Schema $responses $schema (Get-SchemaName $schema 'body') }
            else { Add-Row $responses '(No response body schema)' }
            $headers = Get-Value $r 'headers'
            if ($headers) {
                foreach ($header in $headers.PSObject.Properties) {
                    Write-Schema $responses $header.Value ("$($header.Name) [response header]")
                }
            }
        }
    }
}
if ($operationCount -eq 0) { throw 'No supported HTTP operations were found in paths.' }

# Build an Office Open XML package. All schema text is stored as literal strings.
Add-Type -AssemblyName System.IO.Compression
$parent = [System.IO.Path]::GetDirectoryName($outputFile)
[void][System.IO.Directory]::CreateDirectory($parent)
$tempFile = Join-Path $parent ([Guid]::NewGuid().ToString() + '.tmp')
$stream = $null
$zip = $null
try {
    $stream = [System.IO.File]::Open($tempFile, [System.IO.FileMode]::CreateNew)
    $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
    Add-ZipText $zip '[Content_Types].xml' '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    Add-ZipText $zip '_rels/.rels' '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    Add-ZipText $zip 'xl/workbook.xml' '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Requests" sheetId="1" r:id="rId1"/><sheet name="Responses" sheetId="2" r:id="rId2"/></sheets></workbook>'
    Add-ZipText $zip 'xl/_rels/workbook.xml.rels' '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    Add-ZipText $zip 'xl/styles.xml' @'
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F607A"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8F0F4"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FF7D8D96"/></left><right style="thin"><color rgb="FF7D8D96"/></right><top style="thin"><color rgb="FF7D8D96"/></top><bottom style="thin"><color rgb="FF7D8D96"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4"><xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="2" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@
    Add-ZipText $zip 'xl/worksheets/sheet1.xml' (Get-SheetXml $requests)
    Add-ZipText $zip 'xl/worksheets/sheet2.xml' (Get-SheetXml $responses)
    $zip.Dispose(); $zip = $null
    $stream.Dispose(); $stream = $null
    Move-Item -LiteralPath $tempFile -Destination $outputFile -Force:$Force
} finally {
    if ($zip) { $zip.Dispose() }
    if ($stream) { $stream.Dispose() }
    if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile }
}
Write-Output "Created $outputFile ($operationCount operation(s); $($requests.Count - 1) request rows; $($responses.Count - 1) response rows)."
