# CISA's M-21-31 EOP Phishing Reporting Script

# DRAFT DO NOT USE

This PowerShell script automates reporting of phishing emails detected by Microsoft's Exchange Online Protection (EOP) to CISA. Agencies must report all phishing emails to CISA, as required by M-21-31.

## Features:
- Downloads ***quarantined*** emails marked as `Phishing` or `High Confidence Phishing` from EOP and sends them to CISA's federal phishing reporting mailbox. Emails released from quarantine are not sent.
- Supports user and app-based execution methods.
- Tracks emails sent emails to ensure submissions are not duplicated.
- Rate limiting to stay under Exchange Online's [sending limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#receiving-and-sending-limits). This is ~28,000 submissions/day capacity.
- Allows users to specify the day to search.


## Requirements

- PowerShell v5.1 or higher (PowerShell 7 might work)
- Module: ExchangeOnlineManagement
- Module: Microsoft.Graph.Users.Actions
- [WinRM basic auth](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#turn-on-basic-authentication-in-winrm) enabled on the executing endpoint. This is to download the messages from EXO via remote PowerShell.
- Microsoft commercial and GCC support. (GCC-High/DOD are not supported.)

### User Permissions

The user needs `Security Admin` or higher and must have an Exchange mailbox.

The script will authenticate and send emails as the user specified in `SenderUPN`.

### App Permissions

Scopes:
- `Exchange.ManageAsApp` in Office 365 Exchange Online
- `Mail.Send` in Microsoft Graph

Azure AD Roles:
- `Security Admin` or higher

The script will send emails to CISA as the user specified in `SenderUPN`.

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

### Example 3: Report emails from current day with application
```PowerShell
.\Report-Phishing.ps1 -SenderUPN "bob@agency.gov" -RecipientUPN "federal.phishing.report@us-cert.gov" -AppId <AppId> -CertificateThumb <CertificateThumbprint> -EXOOrganization <Agency Microsoft domain>
```

## Parameters

### ***-AppId*** 

The Azure Active Directory (AAD) client ID (sometimes called Application ID) to use when authenticating as an application. Required when application authentication is used.

|           |     |
|---------------|---------|
| Type          | GUID    |
| Mandatory     | No      |
| Default Value | None    |
| ParameterSet  | AppAuth |

### ***-CertificateThumb***

The certificate thumprint associated with the AAD application specified in `AppId`. The certificate must be installed locally. Required when application authentication is used.

|           |     |
|---------------|---------|
| Type          | String  |
| Mandatory     | No      |
| Default Value | None    |
| ParameterSet  | AppAuth |

### ***-DateToReport***

The day to search for quarantined phishing emails. Includes the entire day in local time zone.

|               |             |
|---------------|-------------|
| Type          | DateTime    |
| Mandatory     | No          |
| Default Value | Current day |
| ParameterSet  | All         |

### ***-EXOOrganization***

One of the domains associated with the agency's tenant. For example: `usdhs.onmicrosoft.com` or `cisa.dhs.gov`. Required when application authentication is used.

|           |     |
|---------------|---------|
| Type          | String  |
| Mandatory     | No      |
| Default Value | None    |
| ParameterSet  | AppAuth |

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


## Configuring application authentication

To run the script unattended you need to create and permission an application in AAD. See [App-only authentication for unattended scripts in Exchange Online PowerShell and Security & Compliance PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) for instructions.
