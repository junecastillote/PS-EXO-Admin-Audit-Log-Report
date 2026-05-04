function Get-ExoAdminAuditLogs {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)]
        [datetime] $StartDate,

        [Parameter(Mandatory, Position = 1)]
        [datetime] $EndDate,

        [Parameter(Position = 2)]
        [ValidateRange(500, 5000)]
        [int] $PageSize = 500,

        [Parameter()]
        [bool] $ShowProgress = $true,

        [Parameter()]
        [int] $MaxRetryCount = 3
    )

    # -------------------------
    # Initialization
    # -------------------------
    $reportDate = (Get-Date).ToUniversalTime()
    $recordType = 'ExchangeAdmin'
    $retryCount = 0
    $sessionID = (New-Guid).Guid
    $results = [System.Collections.Generic.List[object]]::new()
    $pageIndex = 0
    $totalResult = 0

    try {
        $null = (Get-OrganizationConfig -ErrorAction Stop).DisplayName
    }
    catch {
        SayError "Not connected to Exchange Online."
        return
    }

    SayInfo "Using the following parameters:"
    Say "......................................................................"
    Say "Start Date: $($StartDate)"
    Say "End Date: $($EndDate)"
    Say "Page Size: $($PageSize)"
    # Say "Display Progress Bar: $($ShowProgress)"
    Say "Maximum Retries: $($MaxRetryCount)"
    Say "Search Session Id: $($sessionID)"
    Say "......................................................................"

    if ($EndDate -le $StartDate) {
        SayError "EndDate must be greater than StartDate."
        return
    }

    # -------------------------
    # Helper ScriptBlocks
    # -------------------------
    $ExtractAuditLogs = {
        Search-UnifiedAuditLog `
            -SessionId $sessionID `
            -SessionCommand ReturnLargeSet `
            -StartDate $StartDate `
            -EndDate $EndDate `
            -Formatted `
            -RecordType $recordType `
            -ResultSize $PageSize
    }

    function IsResultProblematic ($inputObject) {
        return (
            $inputObject[-1].ResultIndex -eq -1 -and
            $inputObject[-1].ResultCount -eq 0
        )
    }

    SayInfo "Starting audit logs extraction."

    # -------------------------
    # Initial Page + Retry Logic
    # -------------------------
    do {
        $currentPage = @(& $ExtractAuditLogs)

        if (-not $currentPage) {
            # SayInfo "No results found."
            return
        }

        if (IsResultProblematic $currentPage) {
            if ($retryCount -ge $MaxRetryCount) {
                SayWarning "Result metadata still invalid after $MaxRetryCount retries."
                return
            }

            $retryCount++
            $sessionID = (New-Guid).Guid
            SayInfo "Retry #$retryCount (new SessionId: $sessionID)"
        }

    } while (IsResultProblematic $currentPage)

    # -------------------------
    # Paging Loop
    # -------------------------
    do {
        $pageIndex++
        $totalResult += $currentPage.Count
        SayInfo "Progress: Page $pageIndex, Total = $totalResult"

        $results.AddRange(
            @(
                $currentPage | Select-Object *, @{
                    Name       = 'ReportDate'
                    Expression = { $($reportDate) }
                }, @{
                    Name       = 'StartDate'
                    Expression = { $($StartDate.ToUniversalTime()) }
                }, @{
                    Name       = 'EndDate'
                    Expression = { $($EndDate.ToUniversalTime()) }
                }
            )
        )

        $currentPage = @(& $ExtractAuditLogs)

    } while (
        $currentPage.Count -gt 0
    )

    SayInfo "Audit logs extraction complete."

    return $results
}