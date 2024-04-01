Function Get-ExoAdminAuditLogs {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)]
        $StartDate,

        [Parameter(Mandatory, Position = 1)]
        $EndDate,

        [Parameter(Position = 2)]
        [int]
        $PageSize = 500,

        [Parameter()]
        [bool]
        $ShowProgress = $true,

        [Parameter()]
        [int]
        $MaxRetryCount = 3
    )

    $FormatEnumerationLimit = -1

    ## Define the session ID and record type to use with the Search-UnifiedAuditLog cmdlet.
    $sessionID = (New-Guid).GUID
    $recordType = 'ExchangeAdmin'

    $retryCount = 0
    # $maxRetryCount = 3

    ## Set progress bar visibility
    $ProgressPreference = 'Continue'

    ## Set progress bar style if PowerShell Core
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $PSStyle.Progress.View = 'Classic'
    }

    #Region - Is Exchange Connected?
    try {
        $null = (Get-OrganizationConfig -ErrorAction STOP).DisplayName
    }
    catch [System.Management.Automation.CommandNotFoundException] {
        SayWarning "It looks like you forgot to connect to Remote Exchange PowerShell. You should do that first."
        Return $null
    }
    catch {
        SayError "Something is wrong. You can see the error below. Fix it and try again."
        SayError $_.Exception.Message
        Return $null
    }
    #EndRegion

    #Region ExtractAuditLogs
    Function ExtractAuditLogs {
        Search-UnifiedAuditLog -SessionId $sessionID -SessionCommand ReturnLargeSet -StartDate $startDate -EndDate $endDate -Formatted -RecordType $recordType -ResultSize $PageSize
    }

    #EndRegion
    SayInfo "Using the following parameters:"
    Say "......................................................................"
    Say "Start Date: $($StartDate)"
    Say "End Date: $($EndDate)"
    Say "Page Size: $($PageSize)"
    Say "Display Progress Bar: $($ShowProgress)"
    Say "Maximum Retries: $($MaxRetryCount)"
    Say "......................................................................"

    if ([datetime]($StartDate) -eq [datetime]$EndDate) {
        SayError "The StartDate and EndDate cannot be the same values."
        return $null
    }

    if ([datetime]($EndDate) -le [datetime]($StartDate)) {
        SayError "The EndDate value cannot be older than the StartDate value."
        return $null
    }

    Function IsResultProblematic {
        param (
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            $inputObject
        )
        if ($inputObject[-1].ResultIndex -eq -1 -and $inputObject[-1].ResultCount -eq 0) {
            return $true
        }
        else {
            return $false
        }
    }

    #Region Initial Records

    ## This code region retrieves the initial records based on the specified page size.
    if ($ShowProgress) {
        Write-Progress -Activity "Getting Exchange Admin Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: Getting the initial $($pageSize) records based on the page size (0%)" -PercentComplete 0 -ErrorAction SilentlyContinue
    }

    SayInfo "Progress: Getting the initial $($pageSize) records based on the page size (0%)"
    do {
        $currentPageResult = @(ExtractAuditLogs)

        if ($currentPageResult.Count -lt 1) {
            SayInfo "No results found"
            return $null
        }

        ## In some instances, the ResultIndex and ResultCount returned shows -1 and 0 respectively.
        ## When this happens, the output will not be accurate, so the script will retry the retrieval 2 more times.
        if ($retryCount -gt $maxRetryCount) {
            SayWarning "The result's total count and indexes are problematic after $($maxRetryCount) retries. This may be a temporary error. Try again after a few minutes."
            return $null
        }

        if (($isProblematic = IsResultProblematic -inputObject $currentPageResult) -and ($retryCount -le $maxRetryCount)) {
            $retryCount++
            $sessionID = (New-Guid).Guid
            SayInfo "Retry # $($retryCount)"
        }
    }
    while ($isProblematic)

    ## Initialize the maximum results available variable once.
    $maxResultCount = $($currentPageResult[-1].ResultCount)
    SayInfo "Total entries: $($maxResultCount)"

    ## Set the current page result count.
    $currentPageResultCount = $($currentPageResult[-1].ResultIndex)
    ## Compute the completion percentage
    $percentComplete = ($currentPageResultCount * 100) / $maxResultCount
    ## Display the progress
    if ($ShowProgress) {
        Write-Progress -Activity "Getting Exchange Admin Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: $($currentPageResultCount) of $($maxResultCount) ($([math]::round($percentComplete,2))%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue
    }
    SayInfo "Progress: $($currentPageResultCount) of $($maxResultCount) ($([math]::round($percentComplete,2))%)"
    ## Display the current page results
    $currentPageResult #| Select-Object CreationDate, UserIds, Operations, AuditData, ResultIndex

    #EndRegion Initial 100 Records

    ## Retrieve the rest of the audit log entries
    do {
        $currentPageResult = @(ExtractAuditLogs)
        if ($currentPageResult) {
            ## Set the current page result count.
            $currentPageResultCount = $($currentPageResult[-1].ResultIndex)
            ## Compute the completion percentage
            $percentComplete = ($currentPageResultCount * 100) / $maxResultCount
            ## Display the progress
            if ($ShowProgress) {
                Write-Progress -Activity "Getting Exchange Admin Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: $($currentPageResultCount) of $($maxResultCount) ($([math]::round($percentComplete,2))%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue
            }
            Sayinfo "Progress: $($currentPageResultCount) of $($maxResultCount) ($([math]::round($percentComplete,2))%)"
            ## Display the current page results
            $currentPageResult #| Select-Object CreationDate, UserIds, Operations, AuditData, ResultIndex
        }
    }
    while (
        ## Continue running while the last ResultIndex in the current page is less than the ResultCount value.
        ## Note: "ResultIndex" is not ZERO-based.
        ($currentPageResultCount -lt $maxResultCount) -or ($currentPageResult.Count -gt 0)
    )

    if ($ShowProgress) {
        Write-Progress -Activity "Getting Exchange Admin Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: $($currentPageResultCount) of $($maxResultCount) ($([math]::round($percentComplete,2))%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue -Completed
    }

    SayInfo "Audit logs extraction complete."
}


