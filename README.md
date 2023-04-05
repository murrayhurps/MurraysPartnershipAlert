# Murray's Partnership Alert Tool

This tool allows staff to opt-in for alerts whenever a new partnership is being established with someone they have met with in the last three months.

The information is sourced from the calendars of staff who opt-in.

Please note this is completed untested code, and is intended to serve only as a proof of concept in how this might be implemented.

## Usage

To use the Murray's Partnership Alert Tool, follow these steps:

1. Go to the URL where the app is hosted (URL TBD).

2. Enter your email address in the opt-in form and click the "Opt-in" button.

3. Whenever you are exploring a new partnership, open the link to report a new partnership (URL TBD).

4. Enter the relevant partner email address, and a one-sentence summary of what the partnership is.

6. Emails will be sent to all users who have opted-in and have a matching calendar event, asking if they would like to be involved.

We could (and probably should) make this a simpler interface with a single URL, but this code is just a proof of concept.

## Installation

### Prerequisites

- Office365 account with access to the Microsoft Graph API
- Azure AD application registration with `User.ReadWrite.All` and `Mail.Send` permissions

### Step-by-step Instructions

1. Host index.js somewhere secure. Maybe restrict to computers on our own network.

2. In the Azure portal, create a new Azure AD application registration with the following settings:

   - Name: Murray's Partnership Alert Tool
   - Supported account types: Accounts in this organizational directory only
   - Redirect URI: `http://<web hosting address>/auth/callback`

3. Once the application is created, note down the following values:

   - Application (client) ID: `{client_id}`
   - Directory (tenant) ID: `{tenant_id}`
   - Client secret: `{client_secret}`

4. In the app registration settings, add the following API permissions:

   - Microsoft Graph:
     - Delegated permissions:
       - User.Read.All
       - Calendars.Read
       - Calendars.Read.Shared
       - Mail.Send
     - Application permissions:
       - User.ReadWrite.All
       - Calendars.ReadWrite
       - Calendars.ReadWrite.Shared

5. Grant admin consent for the application permissions.

6. In the root of the server, create a file named `.env` with the following contents:

```
TENANT_ID={tenant_id}
CLIENT_ID={client_id}
CLIENT_SECRET={client_secret}
SESSION_SECRET=supersecret
```
