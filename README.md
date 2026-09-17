# Microsoft 365 mail sender for Sitefinity CMS

>**Latest supported version**: Sitefinity CMS 15.4.8637.0

>**IMPORTANT**: This repository may not be compatible with the latest or your current Sitefinity CMS version. If you want to use the repository with a specific Sitefinity CMS version, either upgrade the code from this repository or your Sitefinity CMS project to ensure compatibility.<br/>
The dev team monitors the repository. You can create a GitHub issue to submit feedback or report bugs. Or make a pull request to submit project enhancements or compatibility changes that support new Sitefinity CMS versions.

## Overview
This is a custom notification sender intended to be used with the Microsoft 365 email functionalities. In addition to the built-in notification profiles, it can be used to authenticate with the Microsoft 365 SMTP servers via OAuth and use them for outbound messages.

This repo contains a mail sender sample which works with the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/use-the-api).
## Prerequisites
- You must have a Sitefinity CMS license.
- Your setup must comply with the system requirements. For more information, see the [System requirements](https://www.progress.com/documentation/sitefinity-cms/system-requirements) for the respective Sitefinity CMS version.
- You must have a Microsoft 365 account with Exchange Online, an app registration in your Azure Active Directory and the respective client secret values and API permissions. 
## Installation and configuration
For a step by step installation and configuration guide, see the [Sitefinity documentation](https://www.progress.com/documentation/sitefinity-cms/microsoft-365-mail-sender).

### Required assembly binding redirect
In your Sitefinity application's `web.config`, add the following redirect under `configuration/runtime/assemblyBinding`:

```xml
<dependentAssembly>
	<assemblyIdentity name="Microsoft.Kiota.Abstractions"
										publicKeyToken="31bf3856ad364e35"
										culture="neutral" />
	<bindingRedirect oldVersion="0.0.0.0-1.1.1.0"
									 newVersion="1.1.1.0" />
</dependentAssembly>
```

## Additional resources
For more information on how the Microsoft Graph API, which is used by the notification profile, works, see [this article](https://learn.microsoft.com/en-us/graph/use-the-api).


