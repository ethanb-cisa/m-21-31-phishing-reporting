param (
    [Parameter(Mandatory)]    
    [DateTime]
    $DateToReport = (Get-Date),

    [Parameter(Mandatory)] 
    [MailAddress]
    $SenderUPN,
    
    [Parameter(Mandatory)]
    [MailAddress]
    $ReceipientUPN
)

#################
#DRAFT DO NOT USE

# TODO: App auth 
#################

function Connect-Microsoft365 {
    
    if ( -not (Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Opened" -and $_.Availability -eq "Available"}) ) {
        
        Connect-ExchangeOnline -UseRPSSession -UserPrincipalName $SenderUPN        
    }

    if ("Mail.Send" -notin (Get-MgContext).Scopes) {
        
        Connect-MgGraph -Scopes Mail.Send | Out-Null
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
            -PageSize 1000 `
            -Page $PageCount `
        | % { $DaysQuarantineMessages += $_}

        Get-QuarantineMessage `
            -StartReceivedDate $DateToReport.Date `
            -EndReceivedDate $DateToReport.Date.AddDays(1).AddTicks(-1) `
            -Type HighConfPhish `
            -Direction Inbound `
            -Reported $false `
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
            -PageSize 1000 `
            -Page ($PageCount+1)) `
        -or `
        (Get-QuarantineMessage `
            -StartReceivedDate $DateToReport.Date `
            -EndReceivedDate $DateToReport.Date.AddDays(1).AddTicks(-1) `
            -Type HighConfPhish `
            -Direction Inbound `
            -Reported $false `
            -PageSize 1000 `
            -Page ($PageCount+1))
        ) {
            $PageCount++
        }
        else {
            $MoreResults = $False
        }
        
    }

    if (Test-Path -Path $DateToReport.ToString("yyyy-MM-dd")) {
        $DaysReportedMessages = Get-Content -Path $DateToReport.ToString("yyyy-MM-dd")
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
                        Content = "Phishing email reported as attachment per M-21-31."
                    }
                    ToRecipients = @(
                        @{
                            EmailAddress = @{
                                Address = $ReceipientUPN
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
            Send-MgUserMail -UserId $SenderUPN -BodyParameter $params

            $Message.Identity | Out-File -FilePath $DateToReport.ToString("yyyy-MM-dd") -Append

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
