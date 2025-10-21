# Deployment Guide - HiMarine Invoicing

## Firebase Hosting Configuration

**Project ID:** himarine-invoicing  
**Project Number:** 31012344989  
**Hosting URL:** https://himarine-invoicing.web.app

---

## Deployment Steps

### 1. Build the Application

Build the production version of the Angular app:

```bash
npm run build
```

This creates optimized files in `dist/hi-marine-invoicing/browser/`

### 2. Deploy to Firebase

Deploy the built application to Firebase Hosting:

```bash
firebase deploy --only hosting
```

### 3. Verify Deployment

Open your browser and navigate to:
**https://himarine-invoicing.web.app**

---

## Quick Deploy Commands

```bash
# Build and deploy in one go
npm run build && firebase deploy --only hosting
```

---

## Configuration Files

### `.firebaserc`
Specifies which Firebase project to use:
```json
{
  "projects": {
    "default": "himarine-invoicing"
  }
}
```

### `firebase.json`
Configures Firebase Hosting settings:
```json
{
  "hosting": {
    "public": "dist/hi-marine-invoicing/browser",
    "ignore": [
      "firebase.json",
      "**/.*",
      "**/node_modules/**"
    ],
    "rewrites": [
      {
        "source": "**",
        "destination": "/index.html"
      }
    ]
  }
}
```

The `rewrites` section ensures that all routes are handled by Angular's router (for Single Page Application routing).

---

## Firebase Project Links

- **Project Console:** https://console.firebase.google.com/project/himarine-invoicing/overview
- **Live Website:** https://himarine-invoicing.web.app
- **Firebase CLI:** Already configured in this project

---

## Prerequisites

Before deploying, ensure you have:

1. **Firebase CLI installed:**
   ```bash
   npm install -g firebase-tools
   ```

2. **Logged into Firebase:**
   ```bash
   firebase login
   ```

3. **Project initialized** (already done for this project)

---

## Troubleshooting

### Build Errors
If you encounter build errors, try:
```bash
# Clean install dependencies
rm -rf node_modules package-lock.json
npm install

# Try building again
npm run build
```

### Deployment Issues
If deployment fails:
```bash
# Verify you're logged in
firebase login

# Check which project you're using
firebase projects:list

# Ensure you're using the correct project
firebase use himarine-invoicing
```

### Cache Issues
If the website doesn't update after deployment:
- Clear your browser cache (Ctrl+Shift+Delete)
- Try incognito/private browsing mode
- Wait a few minutes for CDN propagation

---

## Production Build Details

The production build includes:
- Minified JavaScript and CSS
- Optimized bundle sizes
- Tree-shaking to remove unused code
- Ahead-of-Time (AOT) compilation
- Route-based code splitting (lazy loading)

**Total Bundle Size:** ~280 KB (initial) + lazy chunks as needed

---

## Security Notes

- All file processing happens **client-side** (in the browser)
- No data is sent to any server
- Files are stored in browser memory only
- HTTPS is enforced by Firebase Hosting

---

## Maintenance

To update the website:
1. Make your code changes
2. Test locally with `npm start`
3. Build with `npm run build`
4. Deploy with `firebase deploy --only hosting`

---

**Last Deployed:** Check the Firebase Console for deployment history

