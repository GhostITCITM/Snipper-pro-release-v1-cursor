# SnipperClone Deployment Guide

## Quick Solution: Test with a Sample Add-in First

Since you're having issues with the XML validation, let's first test with a known working manifest:

1. **Download this sample manifest**: 
   - Go to: https://raw.githubusercontent.com/OfficeDev/Office-Add-in-samples/main/Samples/hello-world/excel-hello-world/manifest.xml
   - Save it as `test-manifest.xml`

2. **Upload to Admin Center**:
   - In M365 Admin Center → Settings → Integrated Apps
   - Choose "Upload custom app"
   - Select "Office Add-in"
   - Upload the `test-manifest.xml`
   - This should work and prove your admin access is set up correctly

## Deploy Your SnipperClone Add-in

You have three options:

### Option 1: Use GitHub Pages (Easiest)

1. Create a GitHub repository
2. Upload the contents of `dist/app/` to the repository
3. Enable GitHub Pages in repository settings
4. Update `manifest-https.xml` to use your GitHub Pages URL:
   - Replace all `https://localhost:8443` with `https://[your-username].github.io/[repo-name]`
5. Upload the updated manifest to Admin Center

### Option 2: Use Office 365 Developer Tenant (Recommended)

1. Get a free Office 365 Developer subscription at: https://developer.microsoft.com/microsoft-365/dev-program
2. This gives you a test tenant where you can:
   - Use centralized deployment without restrictions
   - Test add-ins freely
   - Avoid corporate policy blocks

### Option 3: Local Development Server

1. **Install ngrok** (provides HTTPS tunnel):
   ```powershell
   # Download from https://ngrok.com/download
   # Or use chocolatey:
   choco install ngrok
   ```

2. **Start local server**:
   ```powershell
   node serve-http.js
   ```

3. **Start ngrok tunnel**:
   ```powershell
   ngrok http 3000
   ```

4. **Update manifest**:
   - Copy the HTTPS URL from ngrok (e.g., `https://abc123.ngrok-free.app`)
   - Update all URLs in `manifest-https.xml`
   - Save as `manifest-ngrok.xml`

5. **Upload to Admin Center**

## Troubleshooting XML Validation Errors

If you're still getting XML validation errors:

1. **Validate your manifest**:
   ```powershell
   # Install Office Add-in Validator
   npm install -g office-addin-manifest

   # Validate manifest
   office-addin-manifest validate manifest-https.xml
   ```

2. **Common fixes**:
   - Ensure all URLs use HTTPS (not HTTP or file://)
   - Check that IconUrl points to valid image files
   - Verify the XML schema URL is correct
   - Make sure all required elements are present

3. **Use the minimal manifest**:
   Create `manifest-minimal.xml` with just essential elements to test

## Alternative: Personal Deployment

If admin deployment is blocked, try personal deployment:

1. In Excel, go to: Insert → Add-ins → My Add-ins
2. Click "Upload My Add-in"
3. Browse to your manifest file
4. This deploys only for your account

## Next Steps

Once deployed:
1. Open Excel
2. Look for "SnipperClone" tab in the ribbon
3. Click "Open Viewer" to start using the add-in

## Files You Need

- `manifest-https.xml` - Your manifest with HTTPS URLs
- `dist/app/*` - Your built application files
- A way to serve files over HTTPS (GitHub Pages, ngrok, etc.) 