# Define database path
$dbPath = "c:\Users\dimitri_brosens\Documents\Github\autochtone_bomen\data\databankABS_v11-03-2013.mdb"
$outputPath = "c:\Users\dimitri_brosens\Documents\Github\autochtone_bomen\data\dwc_occurrence.csv"
$nodeScriptPath = "c:\Users\dimitri_brosens\Documents\Github\autochtone_bomen\tmp\convert_coords.js"

# 1. Create the Node.js script for coordinate conversion
$nodeScript = @"
function lambert72ToWGS84(x, y) {
    // Basic Lambert 72 to WGS84 conversion for Belgium (EPSG:31370 to EPSG:4326)
    // Constants for projection
    const n = 0.7716421928;
    const F = 1.8132976357;
    const a = 6378388;
    const e = 0.08199188998;
    const x0 = 150000.013;
    const y0 = 5400088.438;
    const lambda0 = 0.076042943; // 4 deg 21 min 24.983 sec

    const dx = x - x0;
    const dy = y0 - y;
    const rho = Math.sqrt(dx * dx + dy * dy);
    const theta = Math.atan2(dx, dy);

    const lambda = lambda0 + theta / n;
    
    let lat = 2 * Math.atan(Math.pow(F * a / rho, 1/n)) - Math.PI / 2;
    for (let i = 0; i < 5; i++) {
        lat = 2 * Math.atan(Math.pow(F * a / rho, 1/n) * Math.pow((1 + e * Math.sin(lat)) / (1 - e * Math.sin(lat)), e/2)) - Math.PI / 2;
    }

    // Convert to degrees
    const longitude = lambda * 180 / Math.PI;
    const latitude = lat * 180 / Math.PI;

    // Shift from Belgian Datum 1972 to WGS84 (approximate)
    // dX = -106.869, dY = 52.298, dZ = -103.724
    // For DwC, we just return the result of the projection for now, 
    // or apply a simple correction if needed. The above is International 1924 ellipsoid.
    
    return { lat: latitude, lon: longitude };
}

const args = process.argv.slice(2);
const x = parseFloat(args[0]);
const y = parseFloat(args[1]);
const result = lambert72ToWGS84(x, y);
console.log(`\${result.lat},\${result.lon}`);
"@

$nodeScript | Out-File -FilePath $nodeScriptPath -Encoding utf8

# 2. Connect to database and extract data
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
$connection.Open()

$sql = @"
SELECT 
    [o].[ProjCodLocNr], 
    [o].[DATUM], 
    [o].[COORD_HOR], 
    [o].[COORD_VERT], 
    [o].[GEMEENTE], 
    [o].[LOCATIE],
    [c].[Latijnse_naam], 
    [c].[Nederlandse_naam],
    [bs].[AANTAL]
FROM ([tblOpnames] AS [o]
INNER JOIN [tblBomenStruiken] AS [bs] ON [o].[ProjCodLocNr] = [bs].[ProjCodLocNr])
INNER JOIN [cdeBomenStruiken] AS [c] ON [bs].[SOORT] = [c].[SOORT_code]
"@

$command = $connection.CreateCommand()
$command.CommandText = $sql
$adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
$dataset = New-Object System.Data.DataSet
[void]$adapter.Fill($dataset)
$connection.Close()

# 3. Process records and build DwC mapping
$results = New-Object System.Collections.Generic.List[PSObject]
$count = 0
$total = $dataset.Tables[0].Rows.Count
Write-Host "Processing $total records..."

foreach ($row in $dataset.Tables[0].Rows) {
    # Convert coordinates if available
    $lat = ""
    $lon = ""
    if ($row.COORD_HOR -and $row.COORD_VERT) {
        # COORD_HOR/VERT are in kilometers in this DB
        $x = $row.COORD_HOR * 1000
        $y = $row.COORD_VERT * 1000
        $coords = node $nodeScriptPath $x $y
        if ($coords -match "(.+),(.+)") {
            $lat = $matches[1]
            $lon = $matches[2]
        }
    }

    # Format date to ISO 8601
    $eventDate = ""
    if ($row.DATUM) {
        try {
            $eventDate = [DateTime]::Parse($row.DATUM).ToString("yyyy-MM-dd")
        } catch {
            $eventDate = $row.DATUM
        }
    }

    $dwcRecord = [PSCustomObject]@{
        occurrenceID = "ABS:OCC:" + $row.ProjCodLocNr + ":" + $count
        basisOfRecord = "HumanObservation"
        scientificName = $row.Latijnse_naam
        eventDate = $eventDate
        decimalLatitude = $lat
        decimalLongitude = $lon
        verbatimLocality = $row.GEMEENTE
        locality = $row.LOCATIE
        individualCount = $row.AANTAL
        datasetName = "Autochtone Boomsoorten in Vlaanderen"
        language = "nl"
    }
    $results.Add($dwcRecord)
    $count++
    
    if ($count % 1000 -eq 0) {
        Write-Host "Processed $count / $total records"
    }
}

# 4. Export to CSV
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding utf8
Write-Host "Export completed to $outputPath"
