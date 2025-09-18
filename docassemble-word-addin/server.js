const express = require('express');
const https = require('https');
const { createProxyMiddleware } = require('http-proxy-middleware');

const app = express();
const PORT = process.env.PORT || 8444;
const DOCASSEMBLE_URL = 'http://localhost'; // DocAssemble running on HTTP port 80

// Generate self-signed certificates for HTTPS (same as your working example)
let httpsOptions = null;
try {
    const selfsigned = require('selfsigned');
    const attrs = [{ name: 'commonName', value: 'localhost' }];
    const pems = selfsigned.generate(attrs, {
        keySize: 2048,
        days: 365,
        algorithm: 'sha256',
        extensions: [{
            name: 'subjectAltName',
            altNames: [
                { type: 2, value: 'localhost' },
                { type: 7, ip: '127.0.0.1' }
            ]
        }]
    });

    httpsOptions = {
        key: pems.private,
        cert: pems.cert
    };
    console.log('🔒 Generated self-signed HTTPS certificates for localhost');
} catch (error) {
    console.log('⚠️  HTTPS setup failed:', error.message);
    process.exit(1);
}

// Enable CORS and CSP for Office Add-ins
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');

    // Content Security Policy headers for Office iframe embedding
    res.header('X-Frame-Options', 'ALLOWALL');
    res.header('Content-Security-Policy', 'frame-ancestors https://*.officeapps.live.com https://*.office.com https://*.sharepoint.com https://*.microsoftonline.com https://*.office365.com *');

    // Additional headers for Office compatibility
    res.header('X-Content-Type-Options', 'nosniff');
    res.header('X-XSS-Protection', '1; mode=block');

    // Handle preflight OPTIONS requests
    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    next();
});

// Serve our custom manifest.xml
app.get('/manifest.xml', (req, res) => {
    res.setHeader('Content-Type', 'application/xml');
    res.sendFile(__dirname + '/manifest.xml');
});

// Serve our fixed Word JavaScript (BEFORE proxy routes) - matches with or without query params
app.get('/static/office/word.js', (req, res) => {
    console.log('🔧 Serving custom word.js from word-fixed.js (query params:', req.query, ')');
    res.setHeader('Content-Type', 'application/javascript');
    res.sendFile(__dirname + '/word-fixed.js');
});

// Handle Office Dialog API authentication
app.get('/office-auth-dialog', (req, res) => {
    // Create a simple authentication dialog page with login form
    res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>DocAssemble Authentication</title>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
            <style>
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    padding: 20px;
                    margin: 0;
                    background-color: #f8f9fa;
                }
                .container {
                    max-width: 400px;
                    margin: 0 auto;
                    background: white;
                    padding: 30px;
                    border-radius: 8px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }
                .form-group { margin-bottom: 20px; }
                .form-label {
                    display: block;
                    margin-bottom: 5px;
                    font-weight: 500;
                    color: #333;
                }
                .form-control {
                    width: 100%;
                    padding: 10px;
                    border: 1px solid #ddd;
                    border-radius: 4px;
                    font-size: 14px;
                    box-sizing: border-box;
                }
                .btn {
                    background-color: #0078d4;
                    color: white;
                    padding: 12px 20px;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    width: 100%;
                    font-size: 14px;
                    font-weight: 500;
                }
                .btn:hover { background-color: #106ebe; }
                .btn:disabled { background-color: #ccc; cursor: not-allowed; }
                .error { color: #d13438; margin: 10px 0; padding: 10px; background: #fdf2f2; border-radius: 4px; }
                .loading { text-align: center; color: #666; }
                h3 { color: #333; text-align: center; margin-bottom: 30px; }
                .spinner {
                    display: inline-block;
                    width: 16px;
                    height: 16px;
                    border: 2px solid #f3f3f3;
                    border-top: 2px solid #333;
                    border-radius: 50%;
                    animation: spin 1s linear infinite;
                }
                @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
                .default-credentials {
                    background: #e3f2fd;
                    padding: 15px;
                    border-radius: 4px;
                    margin-bottom: 20px;
                    font-size: 12px;
                    color: #1565c0;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h3>Sign in to DocAssemble</h3>

                <div id="loginForm">
                    <div class="form-group">
                        <label class="form-label" for="email">Email</label>
                        <input id="email" class="form-control" type="email" placeholder="Enter your email" required>
                    </div>
                    <div class="form-group">
                        <label class="form-label" for="password">Password</label>
                        <input id="password" class="form-control" type="password" placeholder="Enter your password" required>
                    </div>
                    <div class="form-group">
                        <button id="signInBtn" class="btn" onclick="authenticateUser()">
                            Sign In
                        </button>
                    </div>
                    <div id="errorMessage" style="display: none;"></div>
                </div>

                <div id="loadingDiv" style="display: none;">
                    <div class="loading">
                        <div class="spinner"></div>
                        <p>Authenticating with DocAssemble...</p>
                    </div>
                </div>
            </div>

            <script>
                Office.onReady(() => {
                    console.log('Office Dialog ready');
                });

                async function authenticateUser() {
                    const email = document.getElementById('email').value;
                    const password = document.getElementById('password').value;
                    const signInBtn = document.getElementById('signInBtn');
                    const errorDiv = document.getElementById('errorMessage');

                    if (!email || !password) {
                        showError('Please enter both email and password');
                        return;
                    }

                    // Show loading state
                    document.getElementById('loginForm').style.display = 'none';
                    document.getElementById('loadingDiv').style.display = 'block';

                    try {
                        const response = await fetch('/user/sign-in', {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/x-www-form-urlencoded',
                            },
                            body: 'email=' + encodeURIComponent(email) + '&password=' + encodeURIComponent(password)
                        });

                        if (response.ok) {
                            // Authentication successful
                            console.log('Authentication successful!');
                            Office.context.ui.messageParent(JSON.stringify({
                                success: true,
                                server: 'https://localhost:8444',
                                sessionInfo: { authenticated: true, timestamp: Date.now() }
                            }));
                        } else {
                            // Authentication failed
                            const errorText = response.status === 401 ? 'Invalid credentials' : 'Authentication failed';
                            showError(errorText);
                            resetForm();
                        }
                    } catch (error) {
                        console.error('Authentication error:', error);
                        showError('Network error: ' + error.message);
                        resetForm();
                    }
                }

                function showError(message) {
                    const errorDiv = document.getElementById('errorMessage');
                    errorDiv.innerHTML = '<div class="error">' + message + '</div>';
                    errorDiv.style.display = 'block';
                }

                function resetForm() {
                    document.getElementById('loginForm').style.display = 'block';
                    document.getElementById('loadingDiv').style.display = 'none';
                }

                // Allow Enter key to submit
                document.addEventListener('keypress', function(e) {
                    if (e.key === 'Enter') {
                        authenticateUser();
                    }
                });
            </script>
        </body>
        </html>
    `);
});

// Proxy all DocAssemble Office add-in endpoints
const proxyOptions = {
    target: DOCASSEMBLE_URL,
    changeOrigin: true,
    secure: false, // DocAssemble uses HTTP
    logLevel: 'info',
    onError: (err, req, res) => {
        console.error('Proxy error:', err.message);
        res.status(500).send('Proxy Error: ' + err.message);
    },
    onProxyReq: (proxyReq, req, res) => {
        console.log(`📡 Proxying: ${req.method} https://localhost:${PORT}${req.url} → ${DOCASSEMBLE_URL}${req.url}`);
    },
    onProxyRes: (proxyRes, req, res) => {
        // Override DocAssemble's CSP headers with Office-compatible ones
        proxyRes.headers['x-frame-options'] = 'ALLOWALL';
        proxyRes.headers['content-security-policy'] = 'frame-ancestors https://*.officeapps.live.com https://*.office.com https://*.sharepoint.com https://*.microsoftonline.com https://*.office365.com *';

        // Remove conflicting headers that might interfere
        delete proxyRes.headers['x-content-security-policy'];
        delete proxyRes.headers['x-webkit-csp'];

        console.log(`🔒 Updated CSP headers for: ${req.url}`);
    }
};

// Proxy specific DocAssemble Office endpoints
app.use('/officetaskpane', createProxyMiddleware(proxyOptions));
app.use('/officefunctionfile', createProxyMiddleware(proxyOptions));

// Proxy /static paths but EXCLUDE /static/office/word.js (already handled above)
app.use('/static', createProxyMiddleware({
    ...proxyOptions,
    skip: (req, res) => {
        // Skip proxying for our custom word.js - let Express handle it
        const shouldSkip = req.path === '/static/office/word.js';
        if (shouldSkip) {
            console.log(`⚠️  Skipping proxy for custom route: ${req.path} (query: ${req.url})`);
        }
        return shouldSkip;
    }
}));

app.use('/favicon.ico', createProxyMiddleware(proxyOptions));

// Proxy any other DocAssemble endpoints that might be needed
app.use('/user', createProxyMiddleware(proxyOptions));
app.use('/config', createProxyMiddleware(proxyOptions));
app.use('/utilities', createProxyMiddleware(proxyOptions));

// Default route - show proxy status
app.get('/', (req, res) => {
    res.send(`
        <h1>DocAssemble Word Add-in HTTPS Proxy</h1>
        <p>🔒 Proxy server running on port ${PORT}</p>
        <p>📡 Forwarding requests to: ${DOCASSEMBLE_URL}</p>

        <h2>Add-in Files:</h2>
        <ul>
            <li><a href="/manifest.xml">Manifest (manifest.xml)</a></li>
            <li><a href="/officetaskpane">Task Pane (proxied from DocAssemble)</a></li>
            <li><a href="/officefunctionfile">Function File (proxied from DocAssemble)</a></li>
            <li><a href="/favicon.ico">Icon (proxied from DocAssemble)</a></li>
        </ul>

        <h2>Status:</h2>
        <ul>
            <li>✅ HTTPS enabled with self-signed certificate</li>
            <li>✅ CORS headers configured for Office Add-ins</li>
            <li>✅ Proxying to DocAssemble: ${DOCASSEMBLE_URL}</li>
        </ul>

        <h2>Installation Instructions:</h2>
        <ol>
            <li>Open Word Online or Desktop Word</li>
            <li>Go to Insert → Add-ins → Upload My Add-in</li>
            <li>Select the manifest.xml file from this server</li>
            <li>The DocAssemble add-in should appear in the Home ribbon</li>
        </ol>

        <p><strong>Manifest URL:</strong> <code>https://localhost:${PORT}/manifest.xml</code></p>
    `);
});

// Error handling
app.use((err, req, res, next) => {
    console.error('Server error:', err);
    res.status(500).send('Internal Server Error: ' + err.message);
});

// Start HTTPS server
https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`🚀 DocAssemble Word Add-in Proxy running on https://localhost:${PORT}`);
    console.log(`📄 Manifest available at: https://localhost:${PORT}/manifest.xml`);
    console.log(`🎯 Task pane proxied at: https://localhost:${PORT}/officetaskpane`);
    console.log(`📡 Forwarding to DocAssemble: ${DOCASSEMBLE_URL}`);
    console.log('🔒 HTTPS enabled - compatible with Word Online!');
    console.log('');
    console.log('To install the add-in:');
    console.log('1. Download manifest.xml from https://localhost:${PORT}/manifest.xml');
    console.log('2. Open Word Online or Desktop Word');
    console.log('3. Go to Insert → Add-ins → Upload My Add-in');
    console.log('4. Select the downloaded manifest.xml file');
    console.log('5. The DocAssemble add-in will appear in the ribbon');
});

// Handle graceful shutdown
process.on('SIGINT', () => {
    console.log('\n👋 Shutting down DocAssemble proxy server...');
    process.exit(0);
});

process.on('SIGTERM', () => {
    console.log('\n👋 Shutting down DocAssemble proxy server...');
    process.exit(0);
});