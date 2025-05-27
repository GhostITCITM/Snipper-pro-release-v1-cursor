const http = require('http');
const express = require('express');
const path = require('path');

const app = express();
const PORT = 3000;

// Serve static files from dist/app directory
app.use(express.static(path.join(__dirname, 'dist', 'app'), {
  setHeaders: (res, path) => {
    // Allow CORS for Office Add-ins
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
}));

http.createServer(app).listen(PORT, () => {
  console.log(`HTTP Server running at http://localhost:${PORT}`);
  console.log('\nNEXT STEPS:');
  console.log('1. Download ngrok from https://ngrok.com/download');
  console.log('2. Run: ngrok http 3000');
  console.log('3. Copy the HTTPS URL from ngrok (e.g., https://abc123.ngrok.io)');
  console.log('4. Update manifest-https.xml with the ngrok URL');
  console.log('5. Upload the updated manifest to M365 Admin Center\n');
}); 