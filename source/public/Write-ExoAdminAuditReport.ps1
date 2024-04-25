Function Write-ExoAdminAuditReport {
    [CmdletBinding()]
    param (
        [parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline
        )]
        [ValidateNotNullOrEmpty()]
        $InputObject,

        [parameter()]
        [string]
        $Organization,

        [parameter()]
        [int]
        $TruncateLongValue,

        [parameter()]
        [string]
        $OutHtml,

        [parameter()]
        [Switch]
        $ConvertToLocalTime
    )
    Begin {

        $FormatEnumerationLimit = -1

        #Region - Is Exchange Connected?
        if (!($Organization)) {
            SayInfo "You did not specify the name of the organization."
            try {
                SayInfo "Attempting to get the organization name."
                $Organization = (Get-OrganizationConfig -ErrorAction STOP).DisplayName
                SayInfo "Found it! Your organization name is $($Organization)"
            }
            catch [System.Management.Automation.CommandNotFoundException] {
                SayWarning "It looks like you forgot to connect to Remote Exchange PowerShell. You should do that first."
                SayWarning "Or you can just specify your organization name next time so that I don't have to look for it for you. The parameter is -Organization <organization name>."
                return $null
            }
            catch {
                SayError "Something is wrong. You can see the error below. I can't tell you how to fix it, but you should fix it before retrying."
                SayError $_.Exception.Message
                return $null
            }
        }

        #EndRegion

        if ($OutHtml) {
            New-Item -ItemType File -Path $OutHtml -Force -ErrorAction Stop | Out-Null
        }

        # For use later to determine the oldest and newest entry
        $dateCollection = [System.Collections.ArrayList]@()

        # $ModuleInfo = Get-Module PsExoAdminAuditLogReport
        $ModuleInfo = $PSCmdlet.MyInvocation.MyCommand.Module
        # $tz = ([System.TimeZoneInfo]::Local).DisplayName.ToString().Split(" ")[0]
        # $today = Get-Date -Format "MMMM dd, yyyy HH:mm"
        # $today = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $css = Get-Content (($ModuleInfo.ModuleBase.ToString()) + '\source\private\style.css') -Raw
        $title = "Exchange Admin Audit Log Report for $($Organization)"

        $logCount = 0

        if ($ConvertToLocalTime) {
            $timeZone = [System.TimeZoneInfo]::Local
        }
        else {
            $timeZone = [System.TimeZoneInfo]::Utc
        }
    }

    Process {
        foreach ($item in $InputObject) {
            $audit_data = ($item.AuditData | ConvertFrom-Json)
            # $dateCollection += ($audit_data.CreationTime)
            if ($ConvertToLocalTime) {
                $creationDate = $item.CreationDate.ToLocalTime()
                $startDate = $item.StartDate.ToLocalTime()
                $endDate = $item.EndDate.ToLocalTime()
                $reportDate = $item.ReportDate.ToLocalTime()
            }
            else {
                $creationDate = $item.CreationDate
                $startDate = $item.StartDate
                $endDate = $item.EndDate
                $reportDate = $item.ReportDate
            }
            $null = $dateCollection.Add($creationDate)
            $html2 += '<tr><td>'
            $html2 += '<b>Time: </b>' + (Get-Date $creationDate -Format "yyyy-MM-dd HH:mm:ss") + '<br>'
            $html2 += '<b>Record Id: </b>' + $audit_data.Id + '<br>'
            $html2 += '<b>Admin Id: </b>' + $audit_data.UserId + '<br>'
            $html2 += '<b>Target Object: </b>' + $audit_data.ObjectId + '<br>'
            $html2 += '</td>'
            $html2 += '<td><b>' + $audit_data.Operation + '</b><br><br>'
            foreach ($param in $audit_data.Parameters) {
                if ($TruncateLongValue) {
                    if ($param.Value.length -gt $TruncateLongValue) {
                        $paramValue = ((($param.Value).ToString().SubString(0, $TruncateLongValue)) + "...")
                    }
                    else {
                        $paramValue = $param.Value
                    }
                }
                else {
                    $paramValue = $param.Value
                }

                ## This line prevents rendering the HTML code in auto-reply messages.
                $paramValue = $paramValue.ToString().Replace('<', '&lt;').Replace('>', '&gt;')

                $html2 += ('<b>' + $param.Name + ':</b> ' + $paramValue + '<br>')
            }
            $html2 += '</td></tr>'
            $logCount++
        }
    }
    End {

        if ($logCount -eq 0) {
            SayError "The report data is empty."
            Return $null
        }

        $dateCollection = $dateCollection | Sort-Object
        $latest_item_date = $dateCollection[0]
        $oldest_item_date = $dateCollection[-1]
        SayInfo "Your report covers the period of $($startDate) to $($endDate)"
        SayInfo "I am creating your HTML report now...."
        # $html1 = @()
        $html1 += '<html><head><title>' + $title + '</title>'
        $html1 += '<style type="text/css">'
        $html1 += $css
        $html1 += '</style></head>'
        $html1 += '<body>'
        $html1 += '<table id="tbl">'
        # $html1 += '<tr><td class="head"></td></tr>'
        $html1 += '<tr><th class="section">Exchange Admin Activity Audit Report</th></tr>'
        $html1 += '<tr><td class="head"><b>' + $Organization + '</b></td></tr>'
        # $html1 += '<tr><td class="head"></td></tr>'
        $html1 += '</table>'
        $html1 += '<table id="tbl">'
        # $html1 += '<tr><td></td><br></tr>'
        # $html1 += '<tr><td class="head"><b>' + 'Summary' + '</b></td></tr>'
        # $html1 += '<tr><td></td><br></tr>'
        $html1 += '<tr><td><b>Report - </b><br>'
        $html1 += '<b>' + ('&nbsp;' * 5) + ' > Time zone : </b>' + $timezone.DisplayName + '<br>'
        $html1 += '<b>' + ('&nbsp;' * 5) + ' > Date generated : </b>' + $reportDate.ToString("yyyy-MM-dd HH:mm:ss") + '<br>'
        $html1 += '<b>' + ('&nbsp;' * 5) + ' > Start date : </b>' + $startDate.ToString("yyyy-MM-dd HH:mm:ss") + '<br>'
        $html1 += '<b>' + ('&nbsp;' * 5) + ' > End date : </b>' + $endDate.ToString("yyyy-MM-dd HH:mm:ss") + '<br>'
        $html1 += '<b>Result - </b><br>'
        $html1 += '<b>' + ('&nbsp;' * 5) + ' > Count : </b>' + $logCount + '<br>'
        $html1 += '<b>' + ('&nbsp;' * 5) + ' > Oldest : </b>' + $oldest_item_date.ToString("yyyy-MM-dd HH:mm:ss") + '<br>'
        $html1 += '<b>' + ('&nbsp;' * 5) + ' > Newest : </b>' + $latest_item_date.ToString("yyyy-MM-dd HH:mm:ss") + '<br>'
        $html1 += '</td></tr>'
        $html1 += '</table>'
        $html1 += '<table id="tbl">'
        # $html1 += '<tr><td class="head"><b>' + 'Details' + '</b></td></tr>'
        # $html1 += '<tr><td></td><td></td><br></tr>'
        $html1 += '<tr><td><b>Event</td><td><b>Commands and Parameters</b></td></tr>'

        $html3 += '</table>'
        $html3 += '<table id="tbl">'
        $html3 += '<tr><td class="head"></td></tr>'
        $html3 += '<tr><td class="head"></td></tr>'
        $html3 += '<tr><td class="head"><a href="' + $ModuleInfo.ProjectURI.AbsoluteUri + '" target="_blank">' + $ModuleInfo.Name.ToString() + ' v' + $ModuleInfo.Version.ToString() + ' </td></a><br>'
        $html3 += '<tr><td class="head"></td></tr>'
        $html3 += '</body></html>'

        $htmlBody = ($html1 + $html2 + $html3) -join "`n"
        if ($OutHtml) {
            try {
                $htmlBody | Out-File $OutHtml -Encoding UTF8 -Force -ErrorAction Stop
                SayInfo "You can find the report at $((Resolve-Path $OutHtml).Path)."
                # return $htmlBody
            }
            catch {
                SayError "Something is wrong. You can see the error below. Because of it I cannot save your report to file. Please fix it."
                SayError $_.Exception.Message
                return $null
            }
        }
        else {
            SayInfo "I've created the report object for you, which is basically just an HTML code in my memory."
            SayInfo "If you wanted to save the report to an HTML file, you should use the -ReportFile <path to report.html> parameter."
            SayInfo "Or, you can just pipe the report out to file like ' | Out-File report.html'. But you should already know how to do that."
            $htmlBody
        }
        SayInfo "Audit logs HTML report complete."
        Say "......................................................................"
        Say "Report time      : $($reportDate.ToString("yyyy-MM-dd HH:mm:ss"))"
        Say "Report coverage  - "
        Say "         Start   : $($startDate.ToString("yyyy-MM-dd HH:mm:ss") )"
        Say "         End     : $($endDate.ToString("yyyy-MM-dd HH:mm:ss") )"
        Say "Result coverage  - "
        Say "         Latest  : $(($latest_item_date).ToString("yyyy-MM-dd HH:mm:ss") )"
        Say "         Oldest  : $(($oldest_item_date).ToString("yyyy-MM-dd HH:mm:ss") )"
        Say "Result count     : $($logCount)"
        Say "......................................................................"
    }
}