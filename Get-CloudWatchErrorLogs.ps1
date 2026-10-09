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

$AwsProfile = 'awsvsp-psd-basic-999999999999'
$AwsRegion = 'ap-northeast-1'
$LogGroup = '/aws/containerinsights/tla-dev-eks/application'
# Namespace is configured but is intentionally not used in the current search.
$Namespace = 'tladev'
$QueryLimit = 100000
$PollSeconds = 2
$TimeoutMinutes = 15

function Get-ErrorPatterns {

    $defaultPatterns = @(
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

    $templateFile = Join-Path `
        $PSScriptRoot `
        '_template.xlsx'

    if (-not (Test-Path -LiteralPath $templateFile)) {
        Write-Warning (
            "_template.xlsx was not found. " +
            "Using built-in error patterns."
        )

        return $defaultPatterns
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open(
            [IO.Path]::GetFullPath($templateFile),
            0,
            $true
        )

        try {
            $worksheet =
                $workbook.Worksheets.Item('ErrorPatterns')
        }
        catch {

            Write-Warning (
                "Worksheet 'ErrorPatterns' was not found. " +
                "Using built-in error patterns."
            )

            return $defaultPatterns
        }

        $patterns =
            [System.Collections.Generic.List[string]]::new()

        $row = 1

        while ($true) {

            $value =
                [string]$worksheet.Cells.Item($row,1).Text

            $value = $value.Trim()

            if ([string]::IsNullOrWhiteSpace($value)) {
                break
            }

            $patterns.Add($value)

            $row++
        }

        if ($patterns.Count -eq 0) {

            Write-Warning (
                "No patterns were found in worksheet " +
                "'ErrorPatterns'. Using built-in error patterns."
            )

            return $defaultPatterns
        }

        Write-Host (
            "Loaded $($patterns.Count) error patterns " +
            "from worksheet 'ErrorPatterns'."
        )

        return @($patterns.ToArray())
    }
    catch {

        Write-Warning (
            "Failed to read worksheet 'ErrorPatterns'. " +
            "Using built-in error patterns. Details: " +
            $_.Exception.Message
        )

        return $defaultPatterns
    }
    finally {

        if ($workbook) {
            $workbook.Close($false)
        }

        if ($excel) {
            $excel.Quit()
        }

        foreach ($obj in @(
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

$ErrorPatterns = Get-ErrorPatterns

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

function Add-PatternPivotSheet {
    param(
        $Workbook,
        $SourceWorksheet,
        [int]$LastRow
    )

    $pivotSheet = $null
    $existingSheet = $null
    $pivotCache = $null
    $pivotTable = $null
    $rowField = $null
    $messageField = $null
    $dataField = $null
    $destinationRange = $null

    try {
        # Remove an existing Pivot sheet if it exists.
        try {
            $existingSheet = $Workbook.Worksheets.Item('Pivot')
        }
        catch {
            $existingSheet = $null
        }

        if ($null -ne $existingSheet) {
            $existingSheet.Delete()

            [void][Runtime.InteropServices.Marshal]::ReleaseComObject(
                $existingSheet
            )
            $existingSheet = $null
        }

        # Insert Pivot sheet immediately after
        # the CloudWatch Logs sheet.
        $pivotSheet = $Workbook.Worksheets.Add()
        $pivotSheet.Name = 'Pivot'

        # Build the source range in R1C1 format.
        # The source worksheet contains eight columns from A through H.
        $sourceSheetName = $SourceWorksheet.Name.Replace("'", "''")
        $sourceData = "'$sourceSheetName'!R1C1:R${LastRow}C12"

        # Create the PivotTable cache.
        # 1 represents xlDatabase.
        $pivotCache = $Workbook.PivotCaches().Create(
            1,
            $sourceData
        )

        # Create the PivotTable beginning at cell A3.
        $destinationRange = $pivotSheet.Range('A3')

        $pivotTable = $pivotCache.CreatePivotTable(
            $destinationRange,
            'ErrorPatternPivot'
        )

        # Add Pattern as the row field.
        # 1 represents xlRowField.
        $rowField = $pivotTable.PivotFields('Pattern')
        $rowField.Orientation = 1
        $rowField.Position = 1

        # Count @message for each Pattern.
        # -4112 represents xlCount.
        $messageField = $pivotTable.PivotFields('@message')
        $dataField = $pivotTable.AddDataField(
            $messageField,
            'Count of @message',
            -4112
        )

        $dataField.NumberFormat = '#,##0'
        
        # Set the Pivot sheet appearance.
        $pivotSheet.Cells.Interior.Color = 16777215
        $pivotSheet.Cells.Font.Name = 'Calibri'

        # Adjust the Pivot sheet layout.
        $pivotSheet.Columns.AutoFit() | Out-Null
        $pivotSheet.Activate()
        $pivotSheet.Range('A3').Select() | Out-Null
    }
    finally {
        foreach ($obj in @(
            $dataField,
            $messageField,
            $rowField,
            $pivotTable,
            $destinationRange,
            $pivotCache,
            $pivotSheet,
            $existingSheet
        )) {
            if ($null -ne $obj) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $obj
                )
            }
        }
    }
}

function Set-WorksheetOrder {
    param(
        $Workbook
    )

    $logSheet = $null
    $pivotSheet = $null
    $errorPatternSheet = $null

    try {
        $logSheet =
            $Workbook.Worksheets.Item('CloudWatch Logs')

        $pivotSheet =
            $Workbook.Worksheets.Item('Pivot')

        #
        # Desired order:
        #
        # 1. CloudWatch Logs
        # 2. Pivot
        # 3. ErrorPatterns
        #

        #$logSheet.Move(
        #    $Workbook.Worksheets.Item(1)
        #)

        #$pivotSheet.Move(
        #    $null,
        #    $logSheet
        #)
    }
    catch {
        Write-Warning (
            "Failed to reorder worksheets. Details: " +
            $_.Exception.Message
        )
    }
    finally {
        foreach ($obj in @(
            $errorPatternSheet,
            $pivotSheet,
            $logSheet
        )) {
            if ($null -ne $obj) {
                [void][Runtime.InteropServices.Marshal]::
                    ReleaseComObject($obj)
            }
        }
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
        'year',
        'month',
        'day',
        'hour',
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

    Copy-Item `
        -LiteralPath $templateFile `
        -Destination $Path `
        -Force

    $absolutePath = [IO.Path]::GetFullPath($Path)

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $headerRange = $null
    $dataRange = $null
    $timestampRange = $null
    $datePartRange = $null
    $yearRange = $null
    $monthRange = $null
    $dayRange = $null
    $hourRange = $null
    $messageColumn = $null
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
            throw 'The template copy was opened as read-only.'
        }

        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = 'CloudWatch Logs'
        $worksheet.Cells.Clear()

        # Set the worksheet background color to white.
        $worksheet.Cells.Interior.Color = 16777215

        # Create the header row.
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

        # Align the headers to the top-left.
        # -4160 represents xlTop.
        # -4131 represents xlLeft.
        $headerRange.VerticalAlignment = -4160
        $headerRange.HorizontalAlignment = -4131

        $headerRange.AutoFilter() | Out-Null

        if ($Rows.Count -gt 0) {
            $dataValues = New-Object `
                'object[,]' `
                $Rows.Count, `
                $columns.Count

            for ($r = 0; $r -lt $Rows.Count; $r++) {
                for ($c = 0; $c -lt $columns.Count; $c++) {
                    $columnName = $columns[$c]

                    # The date-part columns are populated with Excel formulas.
                    if ($columnName -in @(
                        'year',
                        'month',
                        'day',
                        'hour'
                    )) {
                        $dataValues[$r, $c] = ''
                        continue
                    }

                    $property =
                        $Rows[$r].PSObject.Properties[$columnName]

                    if (
                        $null -eq $property -or
                        $null -eq $property.Value
                    ) {
                        $dataValues[$r, $c] = ''
                        continue
                    }

                    # Store @timestamp as an Excel datetime value.
                    if ($columnName -eq '@timestamp') {
                        $timestampText = [string]$property.Value
                        $timestampValue = [DateTimeOffset]::MinValue
                        if (
                            [DateTimeOffset]::TryParse(
                                $timestampText,
                                [Globalization.CultureInfo]::
                                    InvariantCulture,
                                (
                                    [Globalization.DateTimeStyles]::AllowWhiteSpaces -bor
                                    [Globalization.DateTimeStyles]::AssumeUniversal
                                ),
                                [ref]$timestampValue
                            )
                        ) {
                            # Convert the UTC timestamp to the local time
                            # of the computer running this script.
                            $localTimestamp =
                                $timestampValue.ToLocalTime()

                            $dataValues[$r, $c] =
                                $localTimestamp.DateTime.ToOADate()
                        }
                        else {
                            # Keep the original value if parsing fails.
                            $dataValues[$r, $c] = $timestampText
                        }
                    }
                    else {
                        $dataValues[$r, $c] =
                            [string]$property.Value
                    }
                }
            }

            $lastDataRow = $Rows.Count + 1

            $dataRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 1),
                $worksheet.Cells.Item(
                    $lastDataRow,
                    $columns.Count
                )
            )

            $dataRange.Value2 = $dataValues
            $dataRange.Interior.Color = 16777215

            # Align all log data to the top-left.
            $dataRange.VerticalAlignment = -4160
            $dataRange.HorizontalAlignment = -4131

            # Format column B as yyyy/mm/dd hh:mm:ss.
            $timestampRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 2),
                $worksheet.Cells.Item($lastDataRow, 2)
            )

            $timestampRange.NumberFormat =
                'yyyy/mm/dd hh:mm:ss'

            # Add formulas for year, month, day, and hour.
            # FormulaR1C1 allows all rows to be updated at once.
            $yearRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 3),
                $worksheet.Cells.Item($lastDataRow, 3)
            )

            $monthRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 4),
                $worksheet.Cells.Item($lastDataRow, 4)
            )

            $dayRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 5),
                $worksheet.Cells.Item($lastDataRow, 5)
            )

            $hourRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 6),
                $worksheet.Cells.Item($lastDataRow, 6)
            )

            # Each formula references @timestamp in column B.
            $yearRange.FormulaR1C1 = '=YEAR(RC[-1])'
            $monthRange.FormulaR1C1 = '=MONTH(RC[-2])'
            $dayRange.FormulaR1C1 = '=DAY(RC[-3])'
            $hourRange.FormulaR1C1 = '=HOUR(RC[-4])'

            # Format year, month, day, and hour as numbers.
            $datePartRange = $worksheet.Range(
                $worksheet.Cells.Item(2, 3),
                $worksheet.Cells.Item($lastDataRow, 6)
            )

            $datePartRange.NumberFormat = '0'
        }

        $usedRange = $worksheet.UsedRange
        $usedRange.Interior.Color = 16777215

        # Align the entire used range to the top-left.
        $usedRange.VerticalAlignment = -4160
        $usedRange.HorizontalAlignment = -4131

        $usedRange.EntireColumn.AutoFit() | Out-Null

        # Configure the @message column.
        # @message is now column L because four columns were added.
        $messageColumn = $worksheet.Columns.Item(12)
        $messageColumn.ColumnWidth = 100
        $messageColumn.WrapText = $true

        # Freeze the header row.
        $worksheet.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true

        # Hide Excel gridlines.
        $excel.ActiveWindow.DisplayGridlines = $false

        # Add the Pivot sheet.
        if ($Rows.Count -gt 0) {
            Add-PatternPivotSheet `
                -Workbook $workbook `
                -SourceWorksheet $worksheet `
                -LastRow ($Rows.Count + 1)

            # Add the executed command to the Pivot sheet.
            $pivotWorksheet = $workbook.Worksheets.Item('Pivot')

            $commandText = ".\Get-CloudWatchErrorLogs.ps1 " +
                "-StartDateTime `"$StartDateTime`" " +
                "-EndDateTime `"$EndDateTime`""

            $pivotWorksheet.Cells.Item(1, 1).Value2 = $commandText

            [void][Runtime.InteropServices.Marshal]::
                ReleaseComObject($pivotWorksheet)
        }

        # Change the order of the sheets
        Set-WorksheetOrder -Workbook $workbook

        # Save the copied template and keep the sensitivity label.
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
            $messageColumn,
            $hourRange,
            $dayRange,
            $monthRange,
            $yearRange,
            $datePartRange,
            $timestampRange,
            $dataRange,
            $headerRange,
            $worksheet,
            $workbook,
            $excel
        )) {
            if ($null -ne $obj) {
                [void][Runtime.InteropServices.Marshal]::
                    ReleaseComObject($obj)
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
