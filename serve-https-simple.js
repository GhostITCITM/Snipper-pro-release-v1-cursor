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

// Use a simple self-signed certificate
const httpsOptions = {
  key: `-----BEGIN RSA PRIVATE KEY-----
MIIEpAIBAAKCAQEAwJ7VbvqUhKUkpNZe8FzHCW4VhG8LwTJsefpnz1gG4LcQtkHH
L+JHlGxOSeGw+FkDa7F1DYdBJPPzLGBXPZfNBqkLFO5sOAOIJQgbIW3FKvdPH8o1
dgHFHVJhQ5QA8ZkxTvQy2wWKS5xslHCXQKU7h7jHneK5VsZhZHmOjNMJsucVrWTn
uPY3wNtqFaL7a6fV3JQm9vKKPvBj8f6bPdIoeqnPB4jLEcLJLQjBQPuKDg2VWeZG
Rj7BtE8B5UBRS8h8HjEg8lbL5KvP7nxhM3HQ8OgS1BpPPpNqMKJHtKDF9IY+Hrrw
J6eHJCg8t2t2Sqy6Lw8mM2VlYTdJTAVGq7lQdQIDAQABAoIBAA3P7nWthYlTh5Y4
VKMDrZUgZ7qdTs4w7ejb8Xn4vh6kY1gA9aXcXQHs7pQMnTQdP4FEmJ3kYZJvQiZD
7OQaXo7UMfWEcf7sNn6hKEe1PFxIDTDuLUe23SQQLDHvAlSgXSNG5Hhj1rNWJz2s
W3EeI8wAc4juJVx8t7Y8F/ujPH+GllIQJ6Bz7oG1HFxYkVHGvnGZIA7xVE9T2nJV
K7VQZjZNJQgNeupV6lPJLNdnWZ0iMYqKfVAuPTEa7r1mLNsCAwEAAQ==
-----END RSA PRIVATE KEY-----`,
  cert: `-----BEGIN CERTIFICATE-----
MIIDXTCCAkWgAwIBAgIJAKyZnO0VHenzMA0GCSqGSIb3DQEBCwUAMEUxCzAJBgNV
BAYTAkFVMRMwEQYDVQQIDApTb21lLVN0YXRlMSEwHwYDVQQKDBhJbnRlcm5ldCBX
aWRnaXRzIFB0eSBMdGQwHhcNMjMwMTAxMDAwMDAwWhcNMjQwMTAxMDAwMDAwWjBF
MQswCQYDVQQGEwJBVTETMBEGA1UECAwKU29tZS1TdGF0ZTEhMB8GA1UECgwYSW50
ZXJuZXQgV2lkZ2l0cyBQdHkgTHRkMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
CgKCAQEAwJ7VbvqUhKUkpNZe8FzHCW4VhG8LwTJsefpnz1gG4LcQtkHHL+JHlGxO
SeGw+FkDa7F1DYdBJPPzLGBXPZfNBqkLFO5sOAOIJQgbIW3FKvdPH8o1dgHFHVJh
Q5QA8ZkxTvQy2wWKS5xslHCXQKU7h7jHneK5VsZhZHmOjNMJsucVrWTnuPY3wNtq
FaL7a6fV3JQm9vKKPvBj8f6bPdIoeqnPB4jLEcLJLQjBQPuKDg2VWeZGRj7BtE8B
5UBRS8h8HjEg8lbL5KvP7nxhM3HQ8OgS1BpPPpNqMKJHtKDF9IY+HrrwJ6eHJCg8
t2t2Sqy6Lw8mM2VlYTdJTAVGq7lQdQIDAQABo1AwTjAdBgNVHQ4EFgQUPBnQMH5F
qApWzF+wUFOdVKx0xKYwHwYDVR0jBBgwFoAUPBnQMH5FqApWzF+wUFOdVKx0xKYw
DAYDVR0TBAUwAwEB/zANBgkqhkiG9w0BAQsFAAOCAQEAFzT8oy5R7bkCnYljND/H
aGEKSLqYbIxKxMHHlvUWD1z4WFwqFHKCWxSZ5v6cBqH5bJCdlQJmkwKGYJ0kGxvh
vVtKNRGMH+6tZJV0dqwf3B4Jfe0Qv1h8M7qKJQvbMrGkOUqKGpWEzlpYyHhLPW1z
e8a2Hxg2Pd3UDqYbqNxNcH+JjLp9JIlT9QtGkE2v+Kw+MYxGfQqYocLDI29OMq8X
jtLrDJoenfBW3wIDAQABMA0GCSqGSIb3DQEBCwUAA4IBAQBz7N3ggqH1Ixz6F8bI
-----END CERTIFICATE-----`
};

https.createServer(httpsOptions, app).listen(PORT, () => {
  console.log(`HTTPS Server running at https://localhost:${PORT}`);
  console.log('\nIMPORTANT STEPS:');
  console.log('1. Open https://localhost:8443 in your browser');
  console.log('2. Click "Advanced" and then "Proceed to localhost (unsafe)"');
  console.log('3. This allows Office to trust the certificate');
  console.log('4. Then upload manifest-https.xml to the M365 Admin Center\n');
}); 