$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$certDir = Join-Path $root "certs"
$null = New-Item -ItemType Directory -Force -Path $certDir

$pfxPath = Join-Path $certDir "localhost.pfx"
$cerPath = Join-Path $certDir "localhost.cer"
$certName = "LLM4Proposals Dev Cert"
$dnsName = "localhost"
$passphrase = "llm4p-dev"

$cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {
  $_.Subject -eq "CN=$dnsName" -and $_.FriendlyName -eq $certName
} | Select-Object -First 1

if (-not $cert) {
  $cert = New-SelfSignedCertificate -DnsName $dnsName -CertStoreLocation "Cert:\CurrentUser\My" -FriendlyName $certName -NotAfter (Get-Date).AddYears(1)
}

$securePass = ConvertTo-SecureString -String $passphrase -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePass | Out-Null
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

$trusted = Get-ChildItem -Path Cert:\CurrentUser\Root | Where-Object {
  $_.Thumbprint -eq $cert.Thumbprint
} | Select-Object -First 1

if (-not $trusted) {
  Import-Certificate -FilePath $cerPath -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null
}

Write-Host "Created/updated dev cert at $pfxPath (passphrase: $passphrase)."
Write-Host "Trusted in CurrentUser Root store."
