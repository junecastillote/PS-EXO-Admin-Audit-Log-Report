function Write-ExoAdminAuditReport {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        $InputObject,

        [string] $Organization,
        [int]    $TruncateLongValue,
        [string] $OutHtml,
        [switch] $ConvertToLocalTime
    )

    begin {

        Add-Type -AssemblyName System.Web
        $FormatEnumerationLimit = -1

        if (-not $Organization) {
            try {
                $Organization = (Get-OrganizationConfig -ErrorAction Stop).DisplayName
            }
            catch {
                SayError "Not connected to Exchange Online and -Organization was not provided."
                return
            }
        }

        if ($OutHtml) {
            New-Item -ItemType File -Path $OutHtml -Force | Out-Null
        }

        $dateCollection = [System.Collections.Generic.List[datetime]]::new()
        $logCount = 0

        $moduleInfo = $PSCmdlet.MyInvocation.MyCommand.Module
        $html = Get-Content (Join-Path $moduleInfo.ModuleBase 'source\private\template.html') -Raw
        # $css = Get-Content (Join-Path $moduleInfo.ModuleBase 'source\private\style.css') -Raw
        $title = "Exchange Admin Audit Log Report for $Organization"
        $reportDate = Get-Date

        # Time handling (single source of truth)
        $toDisplayTime = {
            param($dt)
            if ($ConvertToLocalTime) { $dt.ToLocalTime() } else { $dt }
        }

        $timeZone = if ($ConvertToLocalTime) {
            [System.TimeZoneInfo]::Local
        }
        else {
            [System.TimeZoneInfo]::Utc
        }

        # HTML builders
        $htmlRows = [System.Text.StringBuilder]::new()
    }

    process {

        foreach ($item in $InputObject) {

            $audit = $item.AuditData | ConvertFrom-Json

            $creationDate = & $toDisplayTime $item.CreationDate
            $startDate = & $toDisplayTime $item.StartDate
            $endDate = & $toDisplayTime $item.EndDate

            $null = $dateCollection.Add($creationDate)

            [void]$htmlRows.AppendLine('<tr><td>')
            [void]$htmlRows.AppendLine("<b>Time:</b> $($creationDate.ToString("yyyy-MM-dd HH:mm:ss"))<br>")
            [void]$htmlRows.AppendLine("<b>Record Id:</b> $($audit.Id)<br>")
            [void]$htmlRows.AppendLine("<b>Admin Id:</b> $($audit.UserId)<br>")
            [void]$htmlRows.AppendLine("<b>Target Object:</b> $($audit.ObjectId)<br>")
            [void]$htmlRows.AppendLine('</td><td>')
            [void]$htmlRows.AppendLine("<b>$($audit.Operation)</b><br><br>")

            foreach ($param in $audit.Parameters) {

                $value = $param.Value

                if ($TruncateLongValue -and $value.Length -gt $TruncateLongValue) {
                    $value = $value.Substring(0, $TruncateLongValue) + '...'
                }

                $value = [System.Web.HttpUtility]::HtmlEncode($value)

                [void]$htmlRows.AppendLine("<b>$($param.Name):</b> $value<br>")
            }

            [void]$htmlRows.AppendLine('</td></tr>')

            $logCount++
        }
    }

    end {

        if ($logCount -eq 0) {
            SayError "The report data is empty."
            return
        }

        $dateCollection.Sort()
        $latest = $dateCollection[-1]
        $oldest = $dateCollection[0]

        $rowHead = "<tr><td><b>Event</b></td><td><b>Commands and Parameters</b></td></tr>"
        $html = $html.Replace("vTitle", $title
        ).Replace("vOrganization", $Organization
        ).Replace("vTimeZone", $($timeZone.DisplayName)
        ).Replace("vReportDate", $($reportDate.ToString("yyyy-MM-dd HH:mm:ss"))
        ).Replace("vStartDate", $($startDate.ToString("yyyy-MM-dd HH:mm:ss"))
        ).Replace("vEndDate", $($endDate.ToString("yyyy-MM-dd HH:mm:ss"))
        ).Replace("vCount", $logCount
        ).Replace("vNewest", $($latest.ToString("yyyy-MM-dd HH:mm:ss"))
        ).Replace("vOldest", $($oldest.ToString("yyyy-MM-dd HH:mm:ss"))
        ).Replace("vRowItems", $($htmlRows.ToString())
        ).Replace("vRepoUrl", $($moduleInfo.ProjectURI.AbsoluteUri)
        ).Replace("vModuleInfo", $("$($moduleInfo.Name) v$($moduleInfo.Version)")
        ).Replace("vRowHead", $rowHead)

        if ($OutHtml) {
            $html | Out-File $OutHtml -Encoding UTF8
            SayInfo "Report written to $((Resolve-Path $OutHtml).Path)"
        }
        else {
            $html
        }

        SayInfo "Audit logs HTML report complete."
    }
}