# Deployment Guide - SWA Ticket Sync

This guide covers multiple deployment options for the SWA Ticket Sync application.

## Prerequisites

Before deploying, ensure:
- ✅ Application works locally (`npm start`)
- ✅ Azure AD app registration is complete
- ✅ API permissions are granted (Mail.Read, Sites.ReadWrite.All)
- ✅ You have admin access to create Azure resources (for Azure deployments)

## Deployment Options

### Option 1: Azure Static Web Apps (Recommended)

Azure Static Web Apps is designed for SPAs and includes:
- Free tier available
- Global CDN
- Custom domains with SSL
- CI/CD from GitHub
- Staging environments

#### Steps:

1. **Build the production bundle:**
   ```bash
   cd swa-ticket-sync
   npm run build
   ```

2. **Install Azure Static Web Apps CLI (optional for testing):**
   ```bash
   npm install -g @azure/static-web-apps-cli
   ```

3. **Deploy via Azure Portal:**

   a. Go to [Azure Portal](https://portal.azure.com)

   b. Create a new **Static Web App**:
   - Click "Create a resource"
   - Search for "Static Web App"
   - Click "Create"

   c. Fill in the details:
   - **Subscription**: Choose your subscription
   - **Resource Group**: Create new or use existing
   - **Name**: `swa-ticket-sync` (or your preferred name)
   - **Plan type**: Free (or Standard for production)
   - **Region**: Choose closest to your users
   - **Source**: GitHub
   - **Organization**: Your GitHub account
   - **Repository**: `outlook-sync`
   - **Branch**: `main`

   d. Build Details:
   - **Build Presets**: React
   - **App location**: `/swa-ticket-sync`
   - **Api location**: (leave empty)
   - **Output location**: `build`

   e. Click "Review + Create" then "Create"

4. **Azure will automatically:**
   - Create a GitHub Actions workflow
   - Build your app on every push
   - Deploy to `https://YOUR-APP-NAME.azurestaticapps.net`

5. **Update Azure AD Redirect URI:**
   - Go to Azure AD → App registrations → Your app
   - Add the new URL as a redirect URI: `https://YOUR-APP-NAME.azurestaticapps.net`
   - Platform type: Single-page application (SPA)

6. **Configure Custom Domain (optional):**
   - In Static Web App → Custom domains
   - Add your domain and follow DNS instructions

#### Environment Variables (if needed):

For Static Web Apps, if you need environment variables:

1. Create `.env.production` in `swa-ticket-sync/`:
   ```
   REACT_APP_CLIENT_ID=your-client-id-here
   ```

2. Update `src/authConfig.js`:
   ```javascript
   export const msalConfig = {
     auth: {
       clientId: process.env.REACT_APP_CLIENT_ID || "5c1e64c0-76f2-4200-8ee5-b3b3d19b53da",
       authority: "https://login.microsoftonline.com/common",
       redirectUri: window.location.origin,
     },
     // ...
   };
   ```

---

### Option 2: Azure Blob Storage with CDN

Cheaper option using static website hosting.

#### Steps:

1. **Create Storage Account:**
   ```bash
   # Install Azure CLI if not already installed
   # https://docs.microsoft.com/en-us/cli/azure/install-azure-cli

   az login
   az group create --name swa-ticket-sync-rg --location eastus
   az storage account create \
     --name swaticketsyncstorage \
     --resource-group swa-ticket-sync-rg \
     --location eastus \
     --sku Standard_LRS \
     --kind StorageV2
   ```

2. **Enable Static Website:**
   ```bash
   az storage blob service-properties update \
     --account-name swaticketsyncstorage \
     --static-website \
     --404-document index.html \
     --index-document index.html
   ```

3. **Build and Upload:**
   ```bash
   cd swa-ticket-sync
   npm run build

   az storage blob upload-batch \
     --account-name swaticketsyncstorage \
     --source ./build \
     --destination '$web'
   ```

4. **Get the URL:**
   ```bash
   az storage account show \
     --name swaticketsyncstorage \
     --resource-group swa-ticket-sync-rg \
     --query "primaryEndpoints.web" \
     --output tsv
   ```

5. **Update Azure AD with the URL**

6. **Add CDN (optional for performance):**
   ```bash
   az cdn profile create \
     --name swa-ticket-sync-cdn \
     --resource-group swa-ticket-sync-rg \
     --sku Standard_Microsoft

   az cdn endpoint create \
     --name swa-ticket-sync \
     --profile-name swa-ticket-sync-cdn \
     --resource-group swa-ticket-sync-rg \
     --origin swaticketsyncstorage.z13.web.core.windows.net
   ```

---

### Option 3: Vercel (Easy GitHub Integration)

Great for quick deployments with automatic HTTPS.

#### Steps:

1. **Install Vercel CLI:**
   ```bash
   npm install -g vercel
   ```

2. **Login:**
   ```bash
   vercel login
   ```

3. **Deploy:**
   ```bash
   cd swa-ticket-sync
   vercel
   ```

4. **Follow prompts:**
   - Set up and deploy: Yes
   - Which scope: Your account
   - Link to existing project: No
   - Project name: `swa-ticket-sync`
   - Directory: `./` (current)
   - Override settings: No
   - Build command: `npm run build`
   - Output directory: `build`
   - Development command: `npm start`

5. **Your app is deployed!** Vercel gives you a URL like `https://swa-ticket-sync.vercel.app`

6. **Update Azure AD Redirect URI** with the Vercel URL

7. **For production domain:**
   ```bash
   vercel --prod
   ```

#### Vercel Configuration File (optional):

Create `vercel.json` in `swa-ticket-sync/`:
```json
{
  "version": 2,
  "name": "swa-ticket-sync",
  "builds": [
    {
      "src": "package.json",
      "use": "@vercel/static-build",
      "config": {
        "distDir": "build"
      }
    }
  ],
  "routes": [
    {
      "src": "/static/(.*)",
      "headers": {
        "cache-control": "s-maxage=31536000,immutable"
      },
      "dest": "/static/$1"
    },
    {
      "src": "/favicon.ico",
      "dest": "/favicon.ico"
    },
    {
      "src": "/manifest.json",
      "dest": "/manifest.json"
    },
    {
      "src": "/(.*)",
      "dest": "/index.html"
    }
  ]
}
```

---

### Option 4: Netlify

Similar to Vercel, great for SPAs.

#### Steps:

1. **Install Netlify CLI:**
   ```bash
   npm install -g netlify-cli
   ```

2. **Login:**
   ```bash
   netlify login
   ```

3. **Initialize:**
   ```bash
   cd swa-ticket-sync
   netlify init
   ```

4. **Configure:**
   - Create & configure a new site
   - Build command: `npm run build`
   - Directory to deploy: `build`
   - Production branch: `main`

5. **Deploy:**
   ```bash
   netlify deploy --prod
   ```

6. **Your app is live at:** `https://YOUR-SITE.netlify.app`

7. **Update Azure AD Redirect URI**

#### Netlify Configuration File:

Create `netlify.toml` in `swa-ticket-sync/`:
```toml
[build]
  command = "npm run build"
  publish = "build"

[[redirects]]
  from = "/*"
  to = "/index.html"
  status = 200

[[headers]]
  for = "/static/*"
  [headers.values]
    cache-control = "public, max-age=31536000, immutable"
```

---

### Option 5: GitHub Pages

Free hosting directly from your GitHub repository.

#### Steps:

1. **Install gh-pages:**
   ```bash
   cd swa-ticket-sync
   npm install --save-dev gh-pages
   ```

2. **Update `package.json`:**
   ```json
   {
     "name": "swa-ticket-sync",
     "version": "1.0.0",
     "homepage": "https://aaronjt12.github.io/outlook-sync",
     "scripts": {
       "predeploy": "npm run build",
       "deploy": "gh-pages -d build",
       "start": "react-scripts start",
       "build": "react-scripts build"
     }
   }
   ```

3. **Update `authConfig.js` to handle GitHub Pages path:**
   ```javascript
   export const msalConfig = {
     auth: {
       clientId: "5c1e64c0-76f2-4200-8ee5-b3b3d19b53da",
       authority: "https://login.microsoftonline.com/common",
       redirectUri: window.location.origin + window.location.pathname,
     },
     cache: {
       cacheLocation: "sessionStorage",
       storeAuthStateInCookie: false,
     }
   };
   ```

4. **Deploy:**
   ```bash
   npm run deploy
   ```

5. **Enable GitHub Pages:**
   - Go to GitHub repository → Settings → Pages
   - Source: Deploy from a branch
   - Branch: `gh-pages` → `/ (root)`
   - Save

6. **Your app will be at:** `https://aaronjt12.github.io/outlook-sync/`

7. **Update Azure AD Redirect URI**

---

## Post-Deployment Checklist

After deploying to any platform:

### 1. Update Azure AD App Registration

- [ ] Add production URL as redirect URI
- [ ] Remove development URLs (http://localhost:3000) if not needed
- [ ] Verify API permissions are granted
- [ ] Test authentication flow

### 2. Test Critical Flows

- [ ] Sign in to Outlook works
- [ ] Sign in to SharePoint works
- [ ] Can select site and list
- [ ] Field mapping saves and loads
- [ ] Can load unread emails
- [ ] Can create tickets successfully
- [ ] Emails are marked as read

### 3. Security Review

- [ ] HTTPS is enabled (should be automatic on all platforms)
- [ ] No secrets or tokens in client-side code
- [ ] sessionStorage is used (not localStorage for tokens)
- [ ] CSP headers configured (optional but recommended)

### 4. Performance Optimization

For production, consider:

**Add to `public/index.html`:**
```html
<!-- Preconnect to Microsoft services -->
<link rel="preconnect" href="https://login.microsoftonline.com">
<link rel="preconnect" href="https://graph.microsoft.com">

<!-- Security headers -->
<meta http-equiv="Content-Security-Policy"
      content="default-src 'self';
               script-src 'self' https://alcdn.msauth.net;
               connect-src 'self' https://login.microsoftonline.com https://graph.microsoft.com;
               style-src 'self' 'unsafe-inline';">
```

---

## Continuous Deployment (CI/CD)

### GitHub Actions (for Azure Static Web Apps)

This is automatically created by Azure, but you can customize:

**.github/workflows/azure-static-web-apps.yml:**
```yaml
name: Azure Static Web Apps CI/CD

on:
  push:
    branches:
      - main
  pull_request:
    types: [opened, synchronize, reopened, closed]
    branches:
      - main

jobs:
  build_and_deploy_job:
    if: github.event_name == 'push' || (github.event_name == 'pull_request' && github.event.action != 'closed')
    runs-on: ubuntu-latest
    name: Build and Deploy Job
    steps:
      - uses: actions/checkout@v3
        with:
          submodules: true

      - name: Build And Deploy
        uses: Azure/static-web-apps-deploy@v1
        with:
          azure_static_web_apps_api_token: ${{ secrets.AZURE_STATIC_WEB_APPS_API_TOKEN }}
          repo_token: ${{ secrets.GITHUB_TOKEN }}
          action: "upload"
          app_location: "/swa-ticket-sync"
          api_location: ""
          output_location: "build"
```

### For Vercel/Netlify

Just push to GitHub - they auto-deploy on push to main branch!

---

## Troubleshooting Deployment

### "Redirect URI mismatch" error
- Check Azure AD redirect URI exactly matches deployed URL
- Include trailing slash if needed
- Ensure platform type is "Single-page application"

### "Build failed" on deployment platform
- Check `npm run build` works locally first
- Ensure all dependencies are in `dependencies` not `devDependencies`
- Check Node version (use 18.x or 20.x)

### "Cannot read property of undefined" errors
- Clear browser cache and sessionStorage
- Check that all environment variables are set
- Verify API calls aren't being blocked by CORS

### MSAL authentication loops
- Clear all cookies and storage
- Verify `redirectUri` in authConfig.js matches deployment URL
- Check `cacheLocation` is set to `sessionStorage`

---

## Recommended: Azure Static Web Apps

For this application, **Azure Static Web Apps** is recommended because:
- ✅ Native Azure integration
- ✅ Free tier is generous
- ✅ Global CDN included
- ✅ Easy custom domains with SSL
- ✅ GitHub Actions CI/CD built-in
- ✅ Staging environments for pull requests
- ✅ Works perfectly with MSAL authentication

---

## Cost Comparison

| Platform | Free Tier | Bandwidth | Custom Domain | SSL |
|----------|-----------|-----------|---------------|-----|
| **Azure Static Web Apps** | 100 GB/month | Unlimited | ✅ | ✅ Auto |
| **Azure Blob Storage** | 5 GB storage | Pay per GB | ✅ | ✅ With CDN |
| **Vercel** | 100 GB/month | 100 GB | ✅ | ✅ Auto |
| **Netlify** | 100 GB/month | 100 GB | ✅ | ✅ Auto |
| **GitHub Pages** | 1 GB storage | 100 GB/month | ✅ | ✅ Auto |

---

## Need Help?

- **Azure Static Web Apps**: https://docs.microsoft.com/en-us/azure/static-web-apps/
- **Vercel**: https://vercel.com/docs
- **Netlify**: https://docs.netlify.com/
- **GitHub Pages**: https://docs.github.com/en/pages
- **MSAL Issues**: https://github.com/AzureAD/microsoft-authentication-library-for-js

---

## Quick Start (Azure Static Web Apps via CLI)

If you want the absolute fastest deployment:

```bash
# Install Azure Static Web Apps CLI
npm install -g @azure/static-web-apps-cli

# Login to Azure
az login

# Deploy (will prompt for details)
cd swa-ticket-sync
swa deploy --app-location . --output-location build

# Follow prompts to create the resource
```

Then just update your Azure AD redirect URI and you're done!
