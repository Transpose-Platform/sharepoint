# Flask App with Azure and SharePoint Integration

This project is a Flask application that integrates with Microsoft 365 SharePoint using Microsoft Graph API. It supports uploading and fetching files from a SharePoint site.

## James Tan Notes
The following is just a rough sketch of how the integration works and is provided only as an example. Generally the process of setting up an integration is:

1. Create the Azure App, provision with the appropriate permissions (which may need to be more permissive than what's listed here).
2. Create the Sharepoint Site, which may involve getting the actual sharepoint site from the customer
3. Use Microsoft Graph Explorer to get the remaining IDs for the Site and Tenant
4. Update the flask app and run the basic commands of reading / uploading.

Some tips:
- It's easier to provision your own Azure environment and Sharepoint site to test the code / integration first then provide these steps to your partner
- Sometimes the permissions changes required more permissive things, I never spent the time to actually narrow down exactly what worked so YMMV.
- Most of the difficulty is in getting the right IDs for everything and Graph Explorer was pretty key for that.

## Prerequisites

1. An Azure account
2. Access to Microsoft 365 with SharePoint
3. Python 3.8+ installed locally

---

## Setup Instructions

### 1. Configure Azure Active Directory (AAD)

1. **Log into Azure Portal**:
   Go to the [Azure Portal](https://portal.azure.com).

2. **Register Your Application**:
   - Navigate to **Azure Active Directory** > **App registrations** > **New registration**.
   - Enter a name for your app (e.g., `SharePointApp`).
   - Set the supported account types to **"Accounts in this organizational directory only"** (or others if needed).
   - Set the redirect URI to `http://localhost:5000` (adjust for production if necessary).
   - Click **Register**.

3. **Save App Credentials**:
   - From the **Overview** page, note the **Application (client) ID** and **Directory (tenant) ID**.
   - Go to **Certificates & Secrets** > **New client secret** to generate a client secret. Save it immediately.

4. **Add API Permissions**:
   - Under **API permissions**, click **Add a permission**.
   - Choose **Microsoft Graph** > **Application permissions**.
   - Add the following permissions:
     - `Sites.ReadWrite.All`
     - `Files.ReadWrite.All`
   - Click **Grant admin consent** to approve these permissions.

---

### 2. Configure SharePoint

1. **Create a SharePoint Site**:
   - Log in to Microsoft 365 Admin Center or SharePoint Online.
   - Create a site (e.g., `SharePointDemo`).
   - Note the site's URL (e.g., `https://<organization>.sharepoint.com/sites/SharePointDemo`).

2. **Determine Site ID and Tenant ID**:
   - Use the [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) to query the **Site ID** and **Tenant ID**:
     
     **Get Site ID**:
     ```
     GET https://graph.microsoft.com/v1.0/sites/<organization>.sharepoint.com:/sites/<SiteName>
     ```
     Look for the `id` field in the response.

     **Get Tenant ID**:
     - The Tenant ID can be found in the URL when accessing the Azure Portal or in the AAD overview for your organization.

---

### 3. Deploy the Flask App

1. **Prepare the App**:
   - Ensure all dependencies are listed in `requirements.txt`:
     ```
     flask
     requests
     msal
     ```
   - Add the following environment variables:
     ```
     SP_CLIENT_ID=<Your Application (client) ID>
     SP_CLIENT_SECRET=<Your Client Secret>
     SP_TENANT_ID=<Your Directory (tenant) ID>
     SP_SITE_ID=<Your SharePoint Site ID>
     ```

---