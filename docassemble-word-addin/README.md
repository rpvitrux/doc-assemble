# Enhanced DocAssemble Word Add-in

## Quick Setup (From Scratch)

### 1. Start DocAssemble
```bash
cd doc-assemble
docker compose up -d
```

### 2. Create Admin Account
- Visit http://localhost
- Register: admin@admin.com / password

### 3. Setup Enhanced Add-in
```bash
cd docassemble-word-addin
npm install
npm start
```

### 4. Install in Word
- Download manifest.xml from https://localhost:8444/manifest.xml
- Word → Insert → Add-ins → Upload My Add-in
- Select manifest.xml

## What Was Built

### Core Problems Solved
1. **Authentication Issue**: DocAssemble's iframe auth blocked by Office 365
   - **Solution**: Implemented Office Dialog API authentication
2. **HTTP/HTTPS Mismatch**: DocAssemble runs HTTP, Office requires HTTPS
   - **Solution**: Created HTTPS proxy server on port 8444
3. **Route Conflicts**: Proxy intercepting custom JavaScript
   - **Solution**: Used skip function in proxy middleware

### Enhanced Features Added
1. **Document Scanning**: Detects `{{ variable }}` patterns in Word documents
2. **Template Building**: Generates questions from found variables
3. **Legal Clause Library**: 6 pre-built clauses for attorneys
4. **Tabbed Interface**: Professional 2-tab design

## File Structure

```
docassemble-word-addin/
├── server.js              # HTTPS proxy server
├── word-fixed.js           # Enhanced add-in JavaScript
├── manifest.xml            # Office add-in manifest
└── package.json            # Dependencies
```
