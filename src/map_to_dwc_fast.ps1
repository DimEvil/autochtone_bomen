function Convert-Lambert72ToWGS84($x, $y) {
    # Basic Lambert 72 to WGS84 conversion for Belgium (EPSG:31370 to EPSG:4326)
    $n = 0.7716421928
    $F = 1.8132976357
    $a = 6378388
    $e = 0.08199188998
    $x0 = 150000.013
    $y0 = 5400088.438
    $lambda0 = 0.076042943 # 4 deg 21 min 24.983 sec

    $dx = $x - $x0
    $dy = $y0 - $y
    $rho = [Math]::Sqrt($dx * $dx + $dy * $dy)
    $theta = [Math]::Atan2($dx, $dy)

    $lambda = $lambda0 + $theta / $n
    
    $lat = 2 * [Math]::Atan([Math]::Pow($F * $a / $rho, 1/$n)) - [Math]::PI / 2
    for ($i = 0; $i -lt 5; $i++) {
        $sinLat = [Math]::Sin($lat)
        $lat = 2 * [Math]::Atan([Math]::Pow($F * $a / $rho, 1/$n) * [Math]::Pow((1 + $e * $sinLat) / (1 - $e * $sinLat), $e/2)) - [Math]::PI / 2
    }

    $longitude = $lambda * 180 / [Math]::PI
    $latitude = $lat * 180 / [Math]::PI

    return @($latitude, $longitude)
}

function Extract-Year($projectCod) {
    if ($projectCod -match "(\d{4})") {
        return $matches[1]
    } elseif ($projectCod -match "(\d{2})") {
        $yr = [int]$matches[1]
        if ($yr -gt 50) { return "19$yr" } else { return "20$yr" }
    }
    return ""
}

function Extract-FullDate($nummer, $fallbackYear) {
    if ($nummer -match "^(\d{2})(\d{2})(\d{2})") {
        $yr = [int]$matches[1]
        $mo = [int]$matches[2]
        $da = [int]$matches[3]
        
        $fullYr = if ($yr -gt 50) { 1900 + $yr } else { 2000 + $yr }
        
        if ($mo -ge 1 -and $mo -le 12 -and $da -ge 1 -and $da -le 31) {
            if ($fallbackYear -ne "" -and $fullYr -ne [int]$fallbackYear) {
                return $fallbackYear
            }
            return "{0:D4}-{1:D2}-{2:D2}" -f $fullYr, $mo, $da
        }
    }
    return $fallbackYear
}

function Get-InheemsDetails($code) {
    if ([string]::IsNullOrWhiteSpace($code)) {
        return @{ em = ""; rem = "" }
    }
    
    $code = $code.ToLower().Trim()
    $remParts = @()
    $em = "uncertain"

    if ($code -match "a") { $remParts += "Autochthonous origin"; $em = "native" }
    if ($code -match "b") { $remParts += "Probably autochthonous origin"; $em = "native" }
    if ($code -match "c") { $remParts += "Possibly autochthonous origin"; $em = "native" }
    if ($code -match "p") { $remParts += "Planted / Introduced"; $em = "introduced" }
    if ($code -match "s") { $remParts += "Spontaneous origin"; $em = "uncertain" }

    $remarks = ($remParts -join "; ") + " (code: $code)"
    return @{ em = $em; rem = $remarks }
}

$dbPath = "c:\Users\dimitri_brosens\Documents\Github\autochtone_bomen\data\databankABS_v11-03-2013.mdb"
$outputPath = "c:\Users\dimitri_brosens\Documents\Github\autochtone_bomen\data\dwc_occurrence.csv"

Write-Host "Connecting to database..."
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
$connection.Open()

$sql = "SELECT [o].[ProjCodLocNr], [o].[PROJECTCOD], [o].[NUMMER], [o].[COORD_HOR], [o].[COORD_VERT], [o].[GEMEENTE], [o].[LOCATIE], [c].[Latijnse_naam], [c].[Nederlandse_naam], [bs].[AANTAL], [bs].[INHEEMS] FROM ([tblOpnames] AS [o] INNER JOIN [tblBomenStruiken] AS [bs] ON [o].[ProjCodLocNr] = [bs].[ProjCodLocNr]) INNER JOIN [cdeBomenStruiken] AS [c] ON [bs].[SOORT] = [c].[SOORT_code]"

$command = $connection.CreateCommand()
$command.CommandText = $sql
$adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
$dataset = New-Object System.Data.DataSet
[void]$adapter.Fill($dataset)
$connection.Close()

$results = New-Object System.Collections.Generic.List[PSObject]
$count = 0
$total = $dataset.Tables[0].Rows.Count
Write-Host "Processing $total records with coordinate rounding (5 decimals)..."

foreach ($row in $dataset.Tables[0].Rows) {
    $lat = ""
    $lon = ""
    if (![Convert]::IsDBNull($row.COORD_HOR) -and ![Convert]::IsDBNull($row.COORD_VERT)) {
        try {
            $coords = Convert-Lambert72ToWGS84 ($row.COORD_HOR * 1000) ($row.COORD_VERT * 1000)
            # Rounded to 5 decimals as requested
            $lat = [Math]::Round($coords[0], 5)
            $lon = [Math]::Round($coords[1], 5)
        } catch {}
    }

    $projYear = Extract-Year $row.PROJECTCOD
    $fullDate = Extract-FullDate $row.NUMMER $projYear
    
    $locParts = @()
    if (![Convert]::IsDBNull($row.GEMEENTE)) { $locParts += $row.GEMEENTE }
    if (![Convert]::IsDBNull($row.LOCATIE)) { $locParts += $row.LOCATIE }
    $locality = $locParts -join ": "

    $inheemsInfo = Get-InheemsDetails $row.INHEEMS

    $dwcRecord = [PSCustomObject]@{
        occurrenceID = "INBO:VBP:ABS:OCC:" + $row.ProjCodLocNr.ToString().Replace(" ", "") + ":" + $count
        basisOfRecord = "HumanObservation"
        scientificName = $row.Latijnse_naam
        vernacularName = $row.Nederlandse_naam
        kingdom = "Plantae"
        eventDate = $fullDate
        year = $projYear
        decimalLatitude = $lat
        decimalLongitude = $lon
        locality = $locality
        individualCount = $row.AANTAL
        establishmentMeans = $inheemsInfo.em
        occurrenceRemarks = $inheemsInfo.rem
        datasetName = "Autochtone boomsoorten en struiken in Vlaanderen"
        institutionCode = "INBO"
        collectionCode = "ABS-monitoring"
        rightsHolder = "Research Institute for Nature and Forest (INBO)"
        license = "http://creativecommons.org/publicdomain/zero/1.0/"
        language = "en"
    }
    $results.Add($dwcRecord)
    $count++
    
    if ($count % 10000 -eq 0) {
        Write-Host "Processed $count / $total records"
    }
}

$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding utf8
Write-Host "Refined export completed to $outputPath"
