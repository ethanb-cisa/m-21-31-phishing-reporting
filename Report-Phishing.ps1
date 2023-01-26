param (    
    [DateTime]
    $DateToReport = (Get-Date),

    [Parameter(Mandatory)] 
    [MailAddress]
    $SenderUPN,
    
    [Parameter(Mandatory)]
    [MailAddress]
    $RecipientUPN,

    [Parameter(ParameterSetName = "AppAuth")]
    [guid]
    $AppId,

    [Parameter(ParameterSetName = "AppAuth")]
    [string]
    $CertificateThumb,

    [Parameter(ParameterSetName = "AppAuth")]
    [string]
    $EXOOrganization
)

$LogFileName = "log-ReportedPhishing-" + $DateToReport.ToString("yyyy-MM-dd")
$LogFilePath = Join-Path -Path $PSScriptRoot -ChildPath $LogFileName

#################
#DRAFT DO NOT USE
#################

function Connect-Microsoft365 {
    
    if ( -not (Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Opened" -and $_.Availability -eq "Available"}) ) {
        if ($PSCmdlet.ParameterSetName -eq "AppAuth") {
            Connect-ExchangeOnline -CertificateThumbprint $CertificateThumb -AppId $AppId -Organization $EXOOrganization
        }
        else {
            Connect-ExchangeOnline -UserPrincipalName $SenderUPN        
        }
    }

    if ("Mail.Send" -notin (Get-MgContext).Scopes) {
        if ($PSCmdlet.ParameterSetName -eq "AppAuth") {
            Connect-MgGraph -TenantId $EXOOrganization -CertificateThumbprint $CertificateThumb -ClientId $AppId | Out-Null
        }
        else {
            Connect-MgGraph -Scopes Mail.Send | Out-Null
        }
    }
}

function Get-UnreportedMessages {
    [CmdletBinding()]
    param (
        [Parameter()]
        [DateTime]
        $DateToReport
    )


    $DaysQuarantineMessages = @()
    
    $MoreResults = $True
    $PageCount = 1
    While ($MoreResults){
        Get-QuarantineMessage `
            -StartReceivedDate $DateToReport.Date `
            -EndReceivedDate $DateToReport.Date.AddDays(1).AddTicks(-1) `
            -Type Phish `
            -Direction Inbound `
            -Reported $false `
            -ReleaseStatus "NotReleased" `
            -PageSize 1000 `
            -Page $PageCount `
        | % { $DaysQuarantineMessages += $_}

        Get-QuarantineMessage `
            -StartReceivedDate $DateToReport.Date `
            -EndReceivedDate $DateToReport.Date.AddDays(1).AddTicks(-1) `
            -Type HighConfPhish `
            -Direction Inbound `
            -Reported $false `
            -ReleaseStatus "NotReleased" `
            -PageSize 1000 `
            -Page $PageCount `
        | % { $DaysQuarantineMessages += $_}

        #Check for more results pages
        if (
        (Get-QuarantineMessage `
            -StartReceivedDate $DateToReport.Date `
            -EndReceivedDate $DateToReport.Date.AddDays(1).AddTicks(-1) `
            -Type Phish `
            -Direction Inbound `
            -Reported $false `
            -ReleaseStatus "NotReleased" `
            -PageSize 1000 `
            -Page ($PageCount+1)) `
        -or `
        (Get-QuarantineMessage `
            -StartReceivedDate $DateToReport.Date `
            -EndReceivedDate $DateToReport.Date.AddDays(1).AddTicks(-1) `
            -Type HighConfPhish `
            -Direction Inbound `
            -Reported $false `
            -ReleaseStatus "NotReleased" `
            -PageSize 1000 `
            -Page ($PageCount+1))
        ) {
            $PageCount++
        }
        else {
            $MoreResults = $False
        }
        
    }

    if (Test-Path -Path $LogFilePath) {
        $DaysReportedMessages = Get-Content -Path $LogFilePath
    }
    $DaysUnreportedMessages = @()

    ForEach ($Message in $DaysQuarantineMessages) {
        if ($Message.Identity -notin $DaysReportedMessages) {
            $DaysUnreportedMessages += $Message
        }
    }

    return $DaysUnreportedMessages
}

function Send-EmailsToCISA {
    param (
        [PSObject[]]
        $QurantineMessagesToReport
    )

    $i=1
    ForEach ($Message in $QurantineMessagesToReport) {
        
        try {
            $FullMessage = Export-QuarantineMessage -Identity $Message.Identity

            $params = @{
                Message = @{
                    Subject = "Federal phishing email submission"
                    Body = @{
                        ContentType = "Text"
                        Content = "Phishing email reported as attachment per M-21-31. Sent via CISA's script."
                    }
                    ToRecipients = @(
                        @{
                            EmailAddress = @{
                                Address = $RecipientUPN
                            }
                        }
                    )
                    Attachments = @(
                        @{
                            "@odata.type" = "#microsoft.graph.fileAttachment"
                            Name = $Message.Subject + ".eml"
                            ContentType = "text/plain"
                            ContentBytes = $FullMessage.eml
                        }
                    )
                }
            }
        }

        catch {
            Write-Error $_
        }

        try {
            Send-MgUserMail -UserId $SenderUPN -BodyParameter $params -ErrorAction Stop

            $Message.Identity | Out-File -FilePath $LogFilePath -Append

            $Status = "Sent " + $i + " of " + $QurantineMessagesToReport.Count
            Write-Host $Status
            $i++

            #EXO has a 30 messages/min rate limit on sent mail. This ensures we stay under it. 
            Start-Sleep -Seconds 3
        }

        catch {
            Write-Error $_
        }
    }
}

Connect-Microsoft365

[PSObject[]]$UnreportedMessages = Get-UnreportedMessages -DateToReport $DateToReport

if ($UnreportedMessages) {
    Send-EmailsToCISA -QurantineMessagesToReport $UnreportedMessages
}
else {
    Write-Host "Nothing new to report for $($DateToReport.ToString("yyyy-MM-dd"))"
}
