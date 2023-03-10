# CISA's M-21-31 EOP Phishing Reporting Script

# DRAFT DO NOT USE

This PowerShell script automates reporting of phishing emails detected by Microsoft's Exchange Online Protection (EOP) to CISA. Agencies must report all phishing emails to CISA, as required by M-21-31.

## Features:
- Downloads ***quarantined*** emails marked as `Phishing` or `High Confidence Phishing` from EOP and sends them to CISA's federal phishing reporting mailbox. Emails released from quarantine are not sent.
- Encrypts quarantined emails into a zip archive (using 7zip command line tool)
- Supports user and app-based execution methods.
- Tracks sent emails to ensure submissions are not duplicated.
- Rate limiting to stay under Exchange Online's [sending limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#receiving-and-sending-limits). This is ~28,000 submissions/day capacity.
- Allows users to specify the day to search.


## Requirements

- PowerShell v5.1 or higher (PowerShell 7 might work)
- Module: `ExchangeOnlineManagement` 3.0.0 or higher
- Module: `Microsoft.Graph.Users.Actions`
- Microsoft commercial and GCC are supported. (GCC-High/DOD are not supported.)
- The ability to send encrypted zip files ***outbound***. (Some Data Loss Prevention/email security appliances block this functionality.)

### User-based execution permissions

The user needs:
- `Security Admin` or higher to download the emails
- An Exchange mailbox to send the emails

The Microsoft Graph PowerShell application needs:
- Delegated permission `Mail.Send` in Microsoft Graph

In this method, the script authenticates and send emails as the user specified in `SenderUPN`.

### Application-based execution permissions

The executing application needs:
- Application permission `Exchange.ManageAsApp` in Office 365 Exchange Online
- Application permission `Mail.Send` in Microsoft Graph
- `Security Admin` or higher to download the emails

In this method, the script sends emails as the user specified in `SenderUPN`. The application is used for authentication.

To run the script unattended you need to create and permission an application in AAD. See [App-only authentication for unattended scripts in Exchange Online PowerShell and Security & Compliance PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) for instructions.

> **Warning**
> An application with the `Mail.Send` application permission can send mail as ***any*** user. You should treat the application as highly privileged. Additionally, you can limit the mailboxes the app can access via [Application Access Policies](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access). 

## Instructions

Download this repository.

1. Click "Code"
2. Click "Download as zip"
3. Extract to a folder
4. Open a PowerShell terminal in that folder.
5. See examples below.

### Example 1: Report emails from current day as a user
```PowerShell
.\Report-Phishing.ps1 -SenderUPN "bob@agency.gov" -RecipientUPN "federal.phishing.report@us-cert.gov"
```

### Example 2: Report emails from specific day as a user
```PowerShell
.\Report-Phishing.ps1 -DateToReport "2023-01-01" -SenderUPN "bob@agency.gov" -RecipientUPN "federal.phishing.report@us-cert.gov"
```

### Example 3: Report emails from current day with application certificate
```PowerShell
.\Report-Phishing.ps1 -SenderUPN "bob@agency.gov" -RecipientUPN "federal.phishing.report@us-cert.gov" -AppId <AppId> -CertificateThumb <CertificateThumbprint>
```

## Parameters

### ***-AppId*** 

The Azure Active Directory (AAD) client ID (sometimes called Application ID) to use when authenticating as an application. Required when application authentication is used.

|           |     |
|---------------|---------|
| Type          | GUID    |
| Mandatory     | No      |
| Default Value | None    |
| ParameterSet  | AppAuthCert and AppAuthSecret |

### ***-CertificateThumb***

The certificate thumprint associated with the AAD application specified in `AppId`. The certificate must be installed locally. Required when application authentication is used.

|           |     |
|---------------|---------|
| Type          | String  |
| Mandatory     | No      |
| Default Value | None    |
| ParameterSet  | AppAuthCert |

### ***-ClientSecret***

The client application's, specified in `AppId`, secret key. 
|           |     |
|---------------|---------|
| Type          | String  |
| Mandatory     | No      |
| Default Value | None    |
| ParameterSet  | AppAuthSecret|

### ***-DateToReport***

The day to search for quarantined phishing emails. Includes the entire day in local time zone.

|               |             |
|---------------|-------------|
| Type          | DateTime    |
| Mandatory     | No          |
| Default Value | Current day |
| ParameterSet  | All         |

### ***-RecipientUPN***

The email address the phishing email will be sent to. This should always be `federal.phishing.report@us-cert.gov`. This parameter is included so you can test functionality before sending.

|               |             |
|---------------|-------------|
| Type          | MailAddress |
| Mandatory     | Yes         |
| Default Value | None        |
| ParameterSet  | All         |

### ***-SenderUPN***
The email address the phishing email will be sent from. With user authentication, this is always the user's email address. With application authentication, it can be any Exchange Online mailbox.

|               |             |
|---------------|-------------|
| Type          | MailAddress |
| Mandatory     | Yes         |
| Default Value | None        |
| ParameterSet  | All         |
