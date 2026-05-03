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

    # Keep caller state safe
    $oldProgressPreference = $ProgressPreference
    $ProgressPreference = 'Continue'

    if ($PSVersionTable.PSEdition -eq 'Core') {
        $PSStyle.Progress.View = 'Classic'
    }

    try {
        $null = (Get-OrganizationConfig -ErrorAction Stop).DisplayName
    }
    catch {
        SayError "Not connected to Exchange Online."
        return
    }

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

    # -------------------------
    # Initial Page + Retry Logic
    # -------------------------
    do {
        $currentPage = @(& $ExtractAuditLogs)

        if (-not $currentPage) {
            SayInfo "No results found."
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

    $maxResultCount = $currentPage[-1].ResultCount
    if ($maxResultCount -lt 1) { return }

    # -------------------------
    # Paging Loop
    # -------------------------
    do {

        $currentIndex = $currentPage[-1].ResultIndex
        $percent = [math]::Min(100, ($currentIndex * 100) / $maxResultCount)

        if ($ShowProgress) {
            Write-Progress `
                -Activity "Getting Exchange Admin Audit Log [$StartDate - $EndDate]" `
                -Status "Progress: $currentIndex of $maxResultCount ($([math]::Round($percent,2))%)" `
                -PercentComplete $percent
        }

        $results.AddRange(
            @(
                $currentPage | Select-Object *, @{
                    Name  = 'ReportDate'
                    Value = $reportDate
                }, @{
                    Name  = 'StartDate'
                    Value = $StartDate.ToUniversalTime()
                }, @{
                    Name  = 'EndDate'
                    Value = $EndDate.ToUniversalTime()
                }
            )
        )

        $currentPage = @(& $ExtractAuditLogs)

    } while (
        $currentPage.Count -gt 0 -and
        $currentIndex -lt $maxResultCount
    )

    if ($ShowProgress) {
        Write-Progress -Activity "Getting Exchange Admin Audit Log" -Completed
    }

    $ProgressPreference = $oldProgressPreference

    SayInfo "Audit logs extraction complete."

    return $results
}