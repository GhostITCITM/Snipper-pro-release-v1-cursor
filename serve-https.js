const https = require('https');
const fs = require('fs');
const path = require('path');
const express = require('express');

const app = express();
const PORT = 8443;

// Serve static files from dist/app directory
app.use(express.static(path.join(__dirname, 'dist', 'app'), {
  setHeaders: (res, path) => {
    // Allow CORS for Office Add-ins
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
}));

// Create self-signed certificate if it doesn't exist
const certPath = path.join(__dirname, 'cert.pem');
const keyPath = path.join(__dirname, 'key.pem');

if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
  console.log('Generating self-signed certificate...');
  const { execSync } = require('child_process');
  
  try {
    // Generate self-signed certificate using OpenSSL
    execSync(`openssl req -x509 -newkey rsa:2048 -keyout ${keyPath} -out ${certPath} -days 365 -nodes -subj "/CN=localhost"`);
    console.log('Certificate generated successfully');
  } catch (error) {
    console.error('Error generating certificate. Make sure OpenSSL is installed.');
    console.error('You can also create the certificate manually.');
    process.exit(1);
  }
}

// Create HTTPS server
const httpsOptions = {
  key: fs.readFileSync(keyPath),
  cert: fs.readFileSync(certPath)
};

https.createServer(httpsOptions, app).listen(PORT, () => {
  console.log(`HTTPS Server running at https://localhost:${PORT}`);
  console.log('\nIMPORTANT: You need to trust the self-signed certificate:');
  console.log('1. Navigate to https://localhost:8443 in your browser');
  console.log('2. Accept the security warning and proceed to the site');
  console.log('3. This allows Office to trust the certificate\n');
}); 