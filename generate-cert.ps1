# Generate self-signed certificate for localhost
$cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(2)

# Export certificate with private key
$pwd = ConvertTo-SecureString -String "password123" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath ".\localhost.pfx" -Password $pwd

Write-Host "Certificate generated: localhost.pfx"
Write-Host "Certificate thumbprint: $($cert.Thumbprint)"
Write-Host "You can now use this certificate for HTTPS" 