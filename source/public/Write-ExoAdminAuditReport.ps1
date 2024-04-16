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
        $ReportFile
    )
    Begin {
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

        if ($ReportFile) {
            New-Item -ItemType File -Path $ReportFile -Force -ErrorAction Stop | Out-Null
        }

        # For use later to determine the oldest and newest entry
        $dateCollection = @()

        $ModuleInfo = Get-Module PsExoAdminAuditLogReport
        # $tz = ([System.TimeZoneInfo]::Local).DisplayName.ToString().Split(" ")[0]
        $today = Get-Date -Format "MMMM dd, yyyy HH:mm (zzzz)"
        $css = Get-Content (($ModuleInfo.ModuleBase.ToString()) + '\source\private\style.css') -Raw
        $title = "Exchange Admin Audit Log Report for $($Organization)"

        $logCount = 0
    }

    Process {
        foreach ($item in ($InputObject)) {
            $audit_data = ($item.AuditData | ConvertFrom-Json)
            $dateCollection += $audit_data.CreationTime
            $html2 += '<tr><td>'
            $html2 += '<b>Time: </b>' + (Get-Date $item.CreationDate -Format "yyyy-MM-dd hh:mm:ss (zzzz)") + '<br>'
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
                $html2 += ('<b>' + $param.Name + ':</b> ' + $paramValue + '<br>')
            }
            $html2 += '</td></tr>'
            $logCount = $logCount + 1
        }
    }
    End {

        if ($logCount -eq 0) {
            SayError "The report data is empty."
            Return $null
        }

        $dateCollection = $dateCollection | Sort-Object
        $startDate = $dateCollection[0]
        $endDate = $dateCollection[-1]
        SayInfo "Your report covers the period of $($startDate) to $($endDate)"
        SayInfo "I am creating your HTML report now...."
        #$html1 = @()
        $html1 += '<html><head><title>' + $title + '</title>'
        $html1 += '<style type="text/css">'
        $html1 += $css
        $html1 += '</style></head>'
        $html1 += '<body>'
        $html1 += '<table id="tbl">'
        $html1 += '<tr><td class="head"</td></tr>'
        $html1 += '<tr><th class="section">Exchange Admin Activity Audit Report</th></tr>'
        $html1 += '<tr><td class="head"><b>' + $Organization + '</b></td></tr>'
        $html1 += '<tr><td class="head"</td></tr>'
        $html1 += '</table>'

        $html1 += '<table id="tbl">'
        $html1 += '<tr><td class="head"><b>' + 'Summary' + '</b></td></tr>'
        $html1 += '<tr><td><b>Report Time:</b> ' + $today + '<br>'
        $html1 += '<b>Coverage:</b><br>'
        $html1 += '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Oldest:</b> ' + $startDate.ToString("yyyy-MM-dd HH:mm:ss (zzzz)") + '<br>'
        $html1 += '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Newset:</b> ' + $endDate.ToString("yyyy-MM-dd HH:mm:ss (zzzz)") + '<br>'
        $html1 += '<b>Activity Count:</b> ' + $logCount
        $html1 += '</td></tr>'
        $html1 += '<tr><td class="head"</td></tr>'
        $html1 += '</table>'
        $html1 += '<table id="tbl">'
        $html1 += '<tr><td class="head"><b>' + 'Details' + '</b></td></tr>'
        $html1 += '<tr><td><b>Event</td><td><b>Commands and Parameters</b></td></tr>'

        $html3 += '</table>'
        $html3 += '<table id="tbl">'
        $html3 += '<tr><td class="head"</td></tr>'
        $html3 += '<tr><td class="head"</td></tr>'
        $html3 += '<tr><td class="head"><a href="' + $ModuleInfo.ProjectURI.AbsoluteUri + '" target="_blank">' + $ModuleInfo.Name.ToString() + ' v' + $ModuleInfo.Version.ToString() + ' </td></a><br>'
        $html3 += '<tr><td class="head"</td></tr>'
        $html3 += '</body></html>'

        $htmlBody = ($html1 + $html2 + $html3) -join "`n"
        if ($ReportFile) {
            try {
                $htmlBody | Out-File $ReportFile -Encoding UTF8 -Force -ErrorAction Stop
                SayInfo "You can find the report at $((Resolve-Path $ReportFile).Path)."
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
        Say "Report time      : $($today)"
        Say "Activity count   : $($logCount)"
        Say "Report coverage  - "
        Say "         Oldest  : $($startDate.ToString("yyyy-MM-dd HH:mm:ss (zzzz)") )"
        Say "         Newest  : $($endDate.ToString("yyyy-MM-dd HH:mm:ss (zzzz)") )"
        Say "......................................................................"
    }
}