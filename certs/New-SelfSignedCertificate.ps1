# Create certificate
$mycert = New-SelfSignedCertificate -DnsName "organization.com" -CertStoreLocation "cert:\LocalMachine\My" -NotAfter (Get-Date).AddYears(5) -KeySpec KeyExchange

# Export certificate to .pfx file
#$mycert | Export-PfxCertificate -FilePath $PSScriptRoot\ERA-Report-Cert.pfx -Password (Get-Credential).password

# Export certificate to .cer file
$mycert | Export-Certificate -FilePath $PSScriptRoot\ERA-Report-Cert.cer