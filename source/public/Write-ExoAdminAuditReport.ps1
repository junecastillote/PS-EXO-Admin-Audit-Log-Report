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
        $css = Get-Content (Join-Path $moduleInfo.ModuleBase 'source\private\style.css') -Raw
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

        $html = @"
<html>
<head>
<title>$title</title>
<style>
$css
</style>
</head>
<body>

<table id="tbl">
<tr><th class="section">Exchange Admin Activity Audit Report</th></tr>
<tr><td class="head"><b>$Organization</b></td></tr>
</table>

<table id="tbl">
<tr><td>
<b>Report</b><br>
&nbsp;&nbsp;&nbsp;&gt; Time zone : $($timeZone.DisplayName)<br>
&nbsp;&nbsp;&nbsp;&gt; Date generated : $($reportDate.ToString("yyyy-MM-dd HH:mm:ss"))<br>
&nbsp;&nbsp;&nbsp;&gt; Start date : $($startDate.ToString("yyyy-MM-dd HH:mm:ss"))<br>
&nbsp;&nbsp;&nbsp;&gt; End date : $($endDate.ToString("yyyy-MM-dd HH:mm:ss"))<br>
<b>Result</b><br>
&nbsp;&nbsp;&nbsp;&gt; Count : $logCount<br>
&nbsp;&nbsp;&nbsp;&gt; Oldest : $($oldest.ToString("yyyy-MM-dd HH:mm:ss"))<br>
&nbsp;&nbsp;&nbsp;&gt; Newest : $($latest.ToString("yyyy-MM-dd HH:mm:ss"))<br>
</td></tr>
</table>

<table id="tbl">
<tr><td><b>Event</b></td><td><b>Commands and Parameters</b></td></tr>
$($htmlRows.ToString())
</table>

<table id="tbl">
<tr>
<td class="head">
<a href="$($moduleInfo.ProjectURI.AbsoluteUri)" target="_blank">
$($moduleInfo.Name) v$($moduleInfo.Version)
</a>
</td>
</tr>
</table>

</body>
</html>
"@

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