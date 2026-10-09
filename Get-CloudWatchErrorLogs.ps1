<#
.SYNOPSIS
    Extracts logs that match known error patterns from AWS CloudWatch Logs Insights.
.DESCRIPTION
    Specify only the start and end date/time.
    Searches each error pattern and combines the results.
    For logs that match multiple patterns, the longest pattern takes priority.
    Outputs the results as JSON and as an Excel file based on the labeled _template.xlsx.
.EXAMPLE
    .\Get-CloudWatchErrorLogs.ps1 -StartDateTime "2026-10-01 09:00" -EndDateTime "2026-10-01 10:00"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$StartDateTime,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$EndDateTime
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$AwsProfile = 'awsvsp-psd-basic-964988438754'
$AwsRegion = 'ap-northeast-1'
$LogGroup = '/aws/containerinsights/tmr-dev-eks/application'
$Namespace = 'tmrdev'
$QueryLimit = 100000
$PollSeconds = 2
$TimeoutMinutes = 15

$ErrorPatterns = @(
    'SIGSEGV',
    'OutOfMemory',
    'out of memory',
    'Java heap space',
    'panic:',
    'JobExecutionException',
    'unhandled Exception',
    'SSL error',
    'SAS/TK is aborting',
    'OAuth token is expired',
    'OOMKilling',
    'Operation timed out',
    'INTERNAL_SERVER_ERROR',
    'I/O error on',
    'ServletOutputStream failed to write:',
    'Gateway Time-out',
    'No space left',
    'endpoints have no available addresses',
    'Traceback',
    'timed out after 60',
    'DiskPressure',
    'NodeHasDiskPressure',
    'FreeDiskSpaceFailed',
    'EvictionThresholdMet',
    'ImageGCFailed',
    'Insufficient ephemeral-storage',
    'NodeHasMemoryPressure',
    'OOMKilled',
    'SystemOOM',
    'NodeHasPIDPressure',
    'FailedScheduling',
    'NodeNotReady',
    'NetworkUnavailable',
    'PROCESS_EVENT_FAILED',
    'Exception'
)

function Convert-ToJstDateTimeOffset {
    param(
        [string]$Text,
        [string]$ParameterName
    )

    try {
        $parsed = Get-Date $Text
    }
    catch {
        throw "$ParameterName has an invalid format. Example: 2026-10-01 09:00"
    }

    [DateTimeOffset]::new(
        $parsed.Year,
        $parsed.Month,
        $parsed.Day,
        $parsed.Hour,
        $parsed.Minute,
        $parsed.Second,
        [TimeSpan]::FromHours(9)
    )
}

function Invoke-AwsCliJson {
    param([string[]]$Arguments)

    $output = & aws @Arguments 2>&1

    if ($LASTEXITCODE -ne 0) {
        throw "AWS CLI command failed.`n$(($output | Out-String).Trim())"
    }

    $text = ($output | Out-String).Trim()

    if ([string]::IsNullOrWhiteSpace($text)) {
        throw 'AWS CLI returned no response.'
    }

    $text | ConvertFrom-Json
}

function Convert-CloudWatchRows {
    param(
        $Results,
        [string]$Pattern,
        [System.Collections.Generic.HashSet[string]]$SeenEventKeys
    )

    $sourceColumns = @(
        '@timestamp',
        'kubernetes.pod_name',
        'kubernetes.container_name',
        'log_processed.level',
        'log_processed.source',
        'kubernetes.namespace_name',
        '@message'
    )

    foreach ($eventFields in $Results) {
        $values = @{}

        foreach ($pair in $eventFields) {
            $values[[string]$pair.field] = [string]$pair.value
        }

        $eventKey = if ($values.ContainsKey('@ptr') -and
                        -not [string]::IsNullOrWhiteSpace([string]$values['@ptr'])) {
            [string]$values['@ptr']
        }
        else {
            @(
                [string]$values['@timestamp'],
                [string]$values['kubernetes.pod_name'],
                [string]$values['kubernetes.container_name'],
                [string]$values['log_processed.source'],
                [string]$values['@message']
            ) -join ([char]0x1F)
        }

        # Patterns are searched from longest to shortest, so an existing event is not added again for a shorter pattern.
        if (-not $SeenEventKeys.Add($eventKey)) {
            continue
        }

        $row = [ordered]@{
            'Pattern' = $Pattern
        }

        foreach ($column in $sourceColumns) {
            $row[$column] = if ($values.ContainsKey($column)) {
                $values[$column]
            }
            else {
                $null
            }
        }

        [pscustomobject]$row
    }
}

function Invoke-CloudWatchPatternQuery {
    param(
        [string]$Pattern,
        [string]$QueryString,
        [long]$StartEpoch,
        [long]$EndEpoch
    )

    $startResponse = Invoke-AwsCliJson @(
        'logs', 'start-query',
        '--profile', $AwsProfile,
        '--region', $AwsRegion,
        '--log-group-name', $LogGroup,
        '--start-time', [string]$StartEpoch,
        '--end-time', [string]$EndEpoch,
        '--query-string', $QueryString,
        '--limit', [string]$QueryLimit,
        '--output', 'json',
        '--no-cli-pager'
    )

    $queryId = [string]$startResponse.queryId

    if ([string]::IsNullOrWhiteSpace($queryId)) {
        throw "Could not get the query ID. Pattern: $Pattern"
    }

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    do {
        Start-Sleep -Seconds $PollSeconds

        $queryResponse = Invoke-AwsCliJson @(
            'logs', 'get-query-results',
            '--profile', $AwsProfile,
            '--region', $AwsRegion,
            '--query-id', $queryId,
            '--output', 'json',
            '--no-paginate',
            '--no-cli-pager'
        )

        $status = [string]$queryResponse.status

        if ((Get-Date) -gt $deadline) {
            throw "The query did not complete within ${TimeoutMinutes} minutes. Pattern: $Pattern"
        }
    }
    while ($status -in @('Scheduled', 'Running', 'Unknown'))

    if ($status -ne 'Complete') {
        throw "The query did not complete successfully. Pattern: $Pattern / Status: $status"
    }

    $statistics = $queryResponse.statistics
    $allResults = [System.Collections.Generic.List[object]]::new()

    foreach ($result in @($queryResponse.results)) {
        $allResults.Add($result)
    }

    $nextToken = if ($queryResponse.PSObject.Properties.Name -contains 'nextToken') {
        [string]$queryResponse.nextToken
    }
    else {
        $null
    }

    while (-not [string]::IsNullOrWhiteSpace($nextToken)) {
        $pageResponse = Invoke-AwsCliJson @(
            'logs', 'get-query-results',
            '--profile', $AwsProfile,
            '--region', $AwsRegion,
            '--query-id', $queryId,
            '--next-token', $nextToken,
            '--output', 'json',
            '--no-paginate',
            '--no-cli-pager'
        )

        foreach ($result in @($pageResponse.results)) {
            $allResults.Add($result)
        }

        $nextToken = if ($pageResponse.PSObject.Properties.Name -contains 'nextToken') {
            [string]$pageResponse.nextToken
        }
        else {
            $null
        }
    }

    [pscustomobject]@{
        Pattern        = $Pattern
        QueryId        = $queryId
        Results        = $allResults.ToArray()
        ReturnedCount  = $allResults.Count
        RecordsMatched = if ($null -ne $statistics) { $statistics.recordsMatched } else { $null }
        RecordsScanned = if ($null -ne $statistics) { $statistics.recordsScanned } else { $null }
        BytesScanned   = if ($null -ne $statistics) { $statistics.bytesScanned } else { $null }
    }
}

function Export-RowsToExcel {
    param(
        [object[]]$Rows,
        [string]$Path
    )

    $columns = @(
        'Pattern',
        '@timestamp',
        'kubernetes.pod_name',
        'kubernetes.container_name',
        'log_processed.level',
        'log_processed.source',
        'kubernetes.namespace_name',
        '@message'
    )

    $templateFile = Join-Path $PSScriptRoot '_template.xlsx'

    if (-not (Test-Path -LiteralPath $templateFile)) {
        throw "_template.xlsx was not found: $templateFile"
    }

    Copy-Item -LiteralPath $templateFile -Destination $Path -Force

    $absolutePath = [IO.Path]::GetFullPath($Path)
    $excel = $null
    $workbook = $null
    $worksheet = $null
    $headerRange = $null
    $dataRange = $null
    $timestampRange = $null
    $usedRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open(
            $absolutePath,
            0,
            $false
        )

        if ($workbook.ReadOnly) {
            throw 'The copied template was opened as read-only.'
        }

        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = 'CloudWatch Logs'
        $worksheet.Cells.Clear()

        # Set the background color of the entire sheet to white.
        $worksheet.Cells.Interior.Color = 16777215

        $headerValues = New-Object 'object[,]' 1, $columns.Count

        for ($c = 0; $c -lt $columns.Count; $c++) {
            $headerValues[0, $c] = $columns[$c]
        }

        $headerRange = $worksheet.Range(
            $worksheet.Cells.Item(1, 1),
            $worksheet.Cells.Item(1, $columns.Count)
        )

        $headerRange.Value2 = $headerValues
        $headerRange.Font.Bold = $true
        $headerRange.Interior.Color = 16777215
        $headerRange.AutoFilter() | Out-Null

        if ($Rows.Count -gt 0) {
            $dataValues = New-Object 'object[,]' $Rows.Count, $columns.Count

            for ($r = 0; $r -lt $Rows.Count; $r++) {
                for ($c = 0; $c -lt $columns.Count; $c++) {
                    $columnName = $columns[$c]
                    $property = $Rows[$r].PSObject.Properties[$columnName]

                    if ($null -eq $property -or $null -eq $property.Value) {
                        $dataValues[$r, $c] = ''
                        continue
                    }

                    # Set column B (@timestamp) as an Excel date value.
                    if ($columnName -eq '@timestamp') {
                        $timestampText = [string]$property.Value
                        $timestampValue = [DateTimeOffset]::MinValue

                        if (
                            [DateTimeOffset]::TryParse(
                                $timestampText,
                                [Globalization.CultureInfo]::InvariantCulture,
                                [Globalization.DateTimeStyles]::AllowWhiteSpaces,
                                [ref]$timestampValue
                            )
                        ) {
                            # Convert date/time values with an offset, such as UTC, to local time.
                            $dataValues[$r, $c] =
                                $timestampValue.LocalDateTime.ToOADate()
                        }
                        else {
                            # If the value cannot be parsed, keep the original string.
                            $dataValues[$r, $c] = $timestampText
                        }
                    }
                    else {
                        $dataValues[$r, $c] = [string]$property.Value
                    }
                }
            }

            $dataRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 1),
                $worksheet.Cells.Item($Rows.Count + 1, $columns.Count)
            )

            $dataRange.Value2 = $dataValues
            $dataRange.Interior.Color = 16777215

            # Set the display format of column B to yyyy/mm/dd hh:mm:ss.
            $timestampRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 2),
                $worksheet.Cells.Item($Rows.Count + 1, 2)
            )

            $timestampRange.NumberFormat = 'yyyy/mm/dd hh:mm:ss'
        }

        $usedRange = $worksheet.UsedRange
        $usedRange.Interior.Color = 16777215
        $usedRange.EntireColumn.AutoFit() | Out-Null

        # @message column
        $worksheet.Columns.Item(8).ColumnWidth = 100
        $worksheet.Columns.Item(8).WrapText = $true

        # Show the sheet and freeze the first row.
        $worksheet.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true

        # Hide the Excel gridlines.
        $excel.ActiveWindow.DisplayGridlines = $false

        # Save over the copied template
        # and keep its sensitivity label.
        $workbook.Save()
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
        }

        if ($excel) {
            $excel.Quit()
        }

        foreach ($obj in @(
            $usedRange,
            $timestampRange,
            $dataRange,
            $headerRange,
            $worksheet,
            $workbook,
            $excel
        )) {
            if ($null -ne $obj) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
            }
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if (-not (Get-Command aws -ErrorAction SilentlyContinue)) {
    throw 'AWS CLI was not found.'
}

$start = Convert-ToJstDateTimeOffset $StartDateTime 'StartDateTime'
$end = Convert-ToJstDateTimeOffset $EndDateTime 'EndDateTime'

if ($end -le $start) {
    throw 'EndDateTime must be later than StartDateTime.'
}

$startEpoch = $start.ToUnixTimeSeconds()
$endEpoch = $end.ToUnixTimeSeconds()

# Keep the original order, but search longer patterns first.
$patternCandidates = @(
    for ($index = 0; $index -lt $ErrorPatterns.Count; $index++) {
        [pscustomobject]@{
            Pattern = [string]$ErrorPatterns[$index]
            Length  = ([string]$ErrorPatterns[$index]).Length
            Index   = $index
        }
    }
)

$sortedPatterns = @(
    $patternCandidates |
        Sort-Object -Property `
            @{ Expression = 'Length'; Descending = $true }, `
            @{ Expression = 'Index';  Descending = $false }
)

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputDirectory = Join-Path $PSScriptRoot "CloudWatchLogs_$stamp"
New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null

$queryPath = Join-Path $outputDirectory 'queries.txt'
$querySummaries = [System.Collections.Generic.List[object]]::new()
$allRows = [System.Collections.Generic.List[object]]::new()
$seenEventKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

foreach ($patternInfo in $sortedPatterns) {
    $pattern = [string]$patternInfo.Pattern

    # / is the Logs Insights regular expression delimiter, so replace it with ., which matches any single character.
    $queryPattern = $pattern -replace '/', '.'

    $queryString = @"
fields @timestamp,
       kubernetes.pod_name,
       kubernetes.container_name,
       log_processed.level,
       log_processed.source,
       kubernetes.namespace_name,
       @message,
       @ptr
| filter @message like /(?i)($queryPattern)/
| sort @timestamp desc
| limit $QueryLimit
"@

    @(
        "===== Pattern: $pattern =====",
        $queryString,
        ''
    ) | Add-Content -Path $queryPath -Encoding UTF8

    Write-Host "Searching: $pattern" -ForegroundColor Cyan

    $patternResult = Invoke-CloudWatchPatternQuery `
        -Pattern $pattern `
        -QueryString $queryString `
        -StartEpoch $startEpoch `
        -EndEpoch $endEpoch

    $beforeCount = $allRows.Count

    foreach ($row in @(Convert-CloudWatchRows `
            -Results $patternResult.Results `
            -Pattern $pattern `
            -SeenEventKeys $seenEventKeys)) {
        $allRows.Add($row)
    }

    $uniqueAdded = $allRows.Count - $beforeCount

    $querySummaries.Add([pscustomobject]@{
        Pattern        = $pattern
        QueryId        = $patternResult.QueryId
        ReturnedCount  = $patternResult.ReturnedCount
        UniqueAdded    = $uniqueAdded
        DuplicateCount = $patternResult.ReturnedCount - $uniqueAdded
        RecordsMatched = $patternResult.RecordsMatched
        RecordsScanned = $patternResult.RecordsScanned
        BytesScanned   = $patternResult.BytesScanned
    })

    Write-Host "  Retrieved: $($patternResult.ReturnedCount) results / New: ${uniqueAdded} results" -ForegroundColor DarkGray

    if ($patternResult.ReturnedCount -ge $QueryLimit) {
        Write-Warning "Results for pattern '$pattern' reached the limit of $QueryLimit. You may need to split the time range."
    }
}

$rows = @($allRows.ToArray() | Sort-Object '@timestamp' -Descending)

$summaryPath = Join-Path $outputDirectory 'cloudwatch-query-summary.json'
$querySummaries |
    ConvertTo-Json -Depth 5 |
    Set-Content -Path $summaryPath -Encoding UTF8

$rowsJsonPath = Join-Path $outputDirectory 'cloudwatch-logs.json'
$rows |
    ConvertTo-Json -Depth 5 |
    Set-Content -Path $rowsJsonPath -Encoding UTF8

$xlsxPath = Join-Path $outputDirectory 'cloudwatch-logs.xlsx'
$excelCreated = $false

if ($rows.Count -gt 0) {
    try {
        Export-RowsToExcel -Rows $rows -Path $xlsxPath
        $excelCreated = $true
    }
    catch {
        Write-Warning "Could not create the Excel file. Details: $($_.Exception.Message)"
    }
}

Write-Host "Completed: $($rows.Count) results / $outputDirectory" -ForegroundColor Green
Write-Host "Queries: $queryPath"
Write-Host "JSON: $rowsJsonPath"
Write-Host "Summary: $summaryPath"

if ($excelCreated) {
    Write-Host "Excel: $xlsxPath"
}
elseif ($rows.Count -eq 0) {
    Write-Host 'No Excel file was created because the search returned 0 results.' -ForegroundColor Yellow
}
