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

$Version = "0.1.0"

$LogFileName = "log-ReportedPhishing-" + $DateToReport.ToString("yyyy-MM-dd")
$LogFilePathPart = Join-Path -path $PSScriptRoot -ChildPath "logs" 
$LogFilePath = Join-Path -Path $LogFilePathPart -ChildPath $LogFileName


#################
#DRAFT DO NOT USE
#################

function Connect-Microsoft365 {
    
    if ( -not (Get-ConnectionInformation | Where-Object {$_.ConnectionUri -eq "https://outlook.office365.com" -and $_.State -eq "Connected" -and $_.TokenStatus -eq "Active"}) ) {
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

function ConvertTo-EncryptedZip {
    param (
        $Message
    )

    $FullMessage = Export-QuarantineMessage -Identity $Message.Identity
    $B64DecodedEML = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($FullMessage.eml))

    $ZipFilePath = $PSScriptRoot + "\encrypted_zips\" + $Message.Identity.Replace("\","_") + ".zip" 
    $7zipParams = @("a", "-si", $ZipFilePath, "-tzip", "-pCISA-PHISHING-REPORT" )
    $B64DecodedEML | & $PSScriptRoot\7za.exe $7zipParams | Out-Null

    $B64EncryptedZip = [Convert]::ToBase64String((Get-Content $ZipFilePath -Encoding Byte))

    return @{"Base64Data" = $B64EncryptedZip; 
             "FilePath" = (Get-ChildItem -Path $ZipFilePath)
            }
}

function Send-EmailsToCISA {
    param (
        [PSObject[]]
        $QurantineMessagesToReport
    )

    $i=1
    ForEach ($Message in $QurantineMessagesToReport) {
        
        try {

            $Zip = ConvertTo-EncryptedZip -Message $Message

            $params = @{
                Message = @{
                    Subject = "M-21-31 Federal phishing email submission"
                    Body = @{
                        ContentType = "Text"
                        Content = "This phishing email is reported as required by M-21-31. This zip is encrypted with the password from CISA's M-21-31 phishing reporting script (version $Version)."
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
                            Name = $Message.Subject + ".zip"
                            ContentType = "text/plain"
                            ContentBytes = $Zip.Base64Data
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
            
            Remove-Item -Path $Zip.FilePath.PSPath 

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
