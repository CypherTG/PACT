param(
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$Args
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$logPath = Join-Path $projectRoot 'temp\pact-gulp-serve.log'
$port = 4321
$tenantWorkbenchUrl = 'https://netorgft13110820.sharepoint.com/sites/KONSTRUCTUM/_layouts/15/workbench.aspx'
$debugManifestsUrl = "http://localhost:$port/temp/build/manifests.js"
$launchWorkbenchUrl = "${tenantWorkbenchUrl}?loadSPFX=true&debugManifestsFile=${debugManifestsUrl}"
$certFriendlyName = 'PACT SPFx Dev Cert'

Add-Type -TypeDefinition @"
using System;
using System.IO;

public sealed class PrefixedStream : Stream
{
    private readonly Stream _inner;
    private readonly byte[] _prefix;
    private int _prefixOffset;

    public PrefixedStream(byte[] prefix, Stream inner)
    {
        _prefix = prefix ?? new byte[0];
        if (inner == null)
        {
            throw new ArgumentNullException("inner");
        }

        _inner = inner;
    }

    public override bool CanRead { get { return true; } }
    public override bool CanSeek { get { return false; } }
    public override bool CanWrite { get { return _inner.CanWrite; } }
    public override long Length { get { throw new NotSupportedException(); } }
    public override long Position
    {
        get { throw new NotSupportedException(); }
        set { throw new NotSupportedException(); }
    }

    public override void Flush()
    {
        _inner.Flush();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        var read = 0;

        if (_prefixOffset < _prefix.Length)
        {
            var prefixCount = Math.Min(count, _prefix.Length - _prefixOffset);
            Array.Copy(_prefix, _prefixOffset, buffer, offset, prefixCount);
            _prefixOffset += prefixCount;
            offset += prefixCount;
            count -= prefixCount;
            read += prefixCount;

            if (count == 0)
            {
                return read;
            }
        }

        read += _inner.Read(buffer, offset, count);
        return read;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        throw new NotSupportedException();
    }

    public override void SetLength(long value)
    {
        throw new NotSupportedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        _inner.Write(buffer, offset, count);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _inner.Dispose();
        }

        base.Dispose(disposing);
    }
}
"@ -ErrorAction SilentlyContinue

function Write-Log([string]$message) {
  $line = "[$((Get-Date).ToString('o'))] $message"
  Add-Content -Path $logPath -Value $line
}

function Get-ContentType([string]$path) {
  switch ([IO.Path]::GetExtension($path).ToLowerInvariant()) {
    '.html' { 'text/html; charset=utf-8' }
    '.js' { 'application/javascript; charset=utf-8' }
    '.json' { 'application/json; charset=utf-8' }
    '.css' { 'text/css; charset=utf-8' }
    '.map' { 'application/json; charset=utf-8' }
    '.png' { 'image/png' }
    '.jpg' { 'image/jpeg' }
    '.jpeg' { 'image/jpeg' }
    '.svg' { 'image/svg+xml' }
    '.txt' { 'text/plain; charset=utf-8' }
    default { 'application/octet-stream' }
  }
}

function Escape-Html([string]$text) {
  [System.Security.SecurityElement]::Escape($text)
}

function New-LocalDevCert {
  $rsa = [System.Security.Cryptography.RSA]::Create(2048)
  $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    'CN=localhost',
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
  )

  $request.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false, $false, 0, $false))
  $request.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
    [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature -bor [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::KeyEncipherment,
    $true
  ))

  $eku = New-Object System.Security.Cryptography.OidCollection
  [void]$eku.Add([System.Security.Cryptography.Oid]::new('1.3.6.1.5.5.7.3.1'))
  $request.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($eku, $false))

  # SAN: DNS=localhost. ASN.1 encoding for a single dNSName entry.
  $sanBytes = [byte[]](0x30, 0x0D, 0x82, 0x0B, 0x6C, 0x6F, 0x63, 0x61, 0x6C, 0x68, 0x6F, 0x73, 0x74)
  $request.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509Extension]::new('2.5.29.17', $sanBytes, $false))

  $cert = $request.CreateSelfSigned((Get-Date).AddMinutes(-5), (Get-Date).AddYears(2))
  $exportPath = Join-Path $env:TEMP 'PACT-SPFx-DevCert.cer'
  $password = ConvertTo-SecureString -String 'PACT-SPFx-DevCert!' -AsPlainText -Force

  try {
    [IO.File]::WriteAllBytes($exportPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    Import-Certificate -FilePath $exportPath -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
  } catch {
    Write-Log "Certificate trust import warning: $($_.Exception.Message)"
  }

  # Re-open the certificate from a PFX blob so the HTTPS listener gets a private key
  # handle that SslStream can use for the TLS handshake.
  $pfxBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $password)
  return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
    $pfxBytes,
    $password,
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet -bor
      [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
  )
}

function Send-Response {
  param(
    [System.IO.Stream]$Stream,
    [int]$StatusCode,
    [string]$StatusText,
    [string]$ContentType,
    [byte[]]$BodyBytes
  )

  $writer = New-Object System.IO.StreamWriter($Stream, [System.Text.Encoding]::ASCII, 1024, $true)
  $writer.NewLine = "`r`n"
  $writer.Write("HTTP/1.1 $StatusCode $StatusText`r`n")
  $writer.Write("Content-Type: $ContentType`r`n")
  $writer.Write("Content-Length: $($BodyBytes.Length)`r`n")
  $writer.Write("Access-Control-Allow-Origin: *`r`n")
  $writer.Write("Cache-Control: no-store`r`n")
  $writer.Write("Connection: close`r`n`r`n")
  $writer.Flush()
  $Stream.Write($BodyBytes, 0, $BodyBytes.Length)
  $Stream.Flush()
}

function Get-StaticFile {
  param([string]$RelativePath)

  $cleanPath = $RelativePath.Split('?', 2)[0].Split('#', 2)[0]
  $safePath = $cleanPath.TrimStart('/').Replace('/', [IO.Path]::DirectorySeparatorChar)
  $candidate = Join-Path $projectRoot $safePath
  if (Test-Path $candidate -PathType Leaf) {
    return $candidate
  }

  if ($cleanPath.StartsWith('/dist/', [StringComparison]::OrdinalIgnoreCase)) {
    $fileName = [IO.Path]::GetFileName($cleanPath)
    $assetRoot = Join-Path $projectRoot 'sharepoint\solution\debug\ClientSideAssets'
    if (Test-Path $assetRoot) {
      $baseName = [IO.Path]::GetFileNameWithoutExtension($fileName)
      $extension = [IO.Path]::GetExtension($fileName)
      $match = Get-ChildItem -Path $assetRoot -File |
        Where-Object {
          $_.Name -like "$baseName*"
        } |
        Sort-Object Length -Descending |
        Select-Object -First 1
      if ($match) {
        return $match.FullName
      }
    }
  }

  return $null
}

function Sync-DebugAssets {
  $assetRoot = Join-Path $projectRoot 'sharepoint\solution\debug\ClientSideAssets'
  $distRoot = Join-Path $projectRoot 'dist'
  if (-not (Test-Path $assetRoot)) {
    return
  }

  New-Item -ItemType Directory -Path $distRoot -Force | Out-Null
  Get-ChildItem -Path $assetRoot -File | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $distRoot $_.Name) -Force
  }
}

function Rewrite-DebugManifests {
  $manifestPath = Join-Path $projectRoot 'temp\build\manifests.js'
  if (-not (Test-Path $manifestPath)) {
    return
  }

  $content = [IO.File]::ReadAllText($manifestPath)
  $updated = $content.Replace('https://localhost:4321/', 'http://localhost:4321/')
  $updated = $updated.Replace('https://localhost:54321/', 'http://localhost:54321/')

  if ($updated -ne $content) {
    [IO.File]::WriteAllText($manifestPath, $updated, [System.Text.Encoding]::UTF8)
    Write-Log "Rewrote manifest URLs to use http for local assets."
  }
}

function Rewrite-PactBundle {
  param([string]$Content)

  if (-not $Content) {
    return $Content
  }

  $updated = $Content

  $updated = $updated.Replace(
    'this.isLocal=e.isServedFromLocalhost',
    'this.isLocal=e.isServedFromLocalhost&&!/(sharepoint\.com|sharepoint-df\.com)/i.test(this.siteUrl)'
  )
  $updated = $updated.Replace('PACT Disciplinary Actions', 'Disciplinary Actions')
  $updated = $updated.Replace('PACT Escalation Log', 'Escalation Log')
  $updated = $updated.Replace('CASES:{TITLE:"Title",CHARGED_PERSON:"ChargedPerson",STAFF_EMAIL:"StaffEmail",DEPARTMENT:"Department",OFFENCE_CATEGORY:"OffenceCategory"', 'CASES:{TITLE:"Case ID",CHARGED_PERSON:"ChargedPerson",STAFF_EMAIL:"StaffEmail",DEPARTMENT:"Department",OFFENCE_CATEGORY:"Offence Category"')
  $updated = $updated.Replace('STAFF:{TITLE:"Title",EMAIL:"EmailAddress",DEPARTMENT:"Department",ROLE:"Role",LINE_MANAGER:"LineManager",COMPANY:"Company",EMPLOYEE_TYPE:"EmployeeType",STATUS:"Status"}', 'STAFF:{TITLE:"Full Name",EMAIL:"EmailAddress",DEPARTMENT:"Department",ROLE:"Role",LINE_MANAGER:"LineManager",COMPANY:"Company",EMPLOYEE_TYPE:"EmployeeType",STATUS:"Status"}')
  $updated = $updated.Replace('MAIL:{TITLE:"Title",TO:"RecipientEmail",BODY:"MailBody",STATUS:"Status"}', 'MAIL:{TITLE:"Subject",TO:"RecipientEmail",BODY:"MailBody",STATUS:"Status"}')
  $updated = $updated.Replace('actionType:x.split(" ")[0]', 'actionType:x')
  $updated = $updated.Replace(
    'case 7:return r.sent(),[4,this.createDisciplinaryAction({caseReference:c.title,actionType:x.split(" ")[0],penaltyAmount:c.penaltyAmount,notes:"Action Classification: ".concat(I,". Recommended by PACT Engine."),status:"Pending"})];case 8:return r.sent(),[2,c]',
    'case 7:return r.sent(),[4,this.createDisciplinaryAction({caseReference:c.title,actionType:x,penaltyAmount:c.penaltyAmount,notes:"Action Classification: ".concat(I,". Recommended by PACT Engine."),status:"Pending"})];case 8:return r.sent(),[4,this.sendEmailNotification([c.staffEmail],"PACT NOTICE: ".concat(c.title," - ").concat(c.offenceCategoryName),"<p>An offence has been logged against <b>".concat(c.chargedPersonName,"</b>.</p><p><b>Offence:</b> ").concat(c.offenceCategoryName,"</p><p><b>Sanction:</b> ").concat(x,"</p>"))];case 9:return r.sent(),[2,c]'
  )

  return $updated
}

function Handle-Client {
  param(
    [System.Net.Sockets.TcpClient]$Client
  )

  $stream = $null
  try {
    $networkStream = $Client.GetStream()
    $stream = $networkStream

    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII)
    $requestLine = $reader.ReadLine()
    if (-not $requestLine) {
      return
    }

    $requestParts = $requestLine.Split(' ')
    if ($requestParts.Count -lt 2) {
      return
    }

    $method = $requestParts[0]
      $path = $requestParts[1]
      $path = $path.Split('?', 2)[0].Split('#', 2)[0]

    while ($true) {
      $header = $reader.ReadLine()
      if ([string]::IsNullOrWhiteSpace($header)) { break }
    }

    if ($method -ne 'GET') {
      $body = [System.Text.Encoding]::UTF8.GetBytes('Method Not Allowed')
      Send-Response -Stream $stream -StatusCode 405 -StatusText 'Method Not Allowed' -ContentType 'text/plain; charset=utf-8' -BodyBytes $body
      return
    }

    if ($path -eq '/' -or $path -eq '/index.html') {
      $escapedWorkbenchUrl = Escape-Html $launchWorkbenchUrl
      $escapedManifestsUrl = Escape-Html $debugManifestsUrl
      $escapedProjectRoot = Escape-Html $projectRoot
      $html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="refresh" content="0;url=$escapedWorkbenchUrl" />
  <title>PACT SPFx SharePoint Workbench Launcher</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 32px; background: linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%); color: #111827; }
    .card { max-width: 840px; background: white; border: 1px solid #e5e7eb; border-radius: 18px; padding: 28px; box-shadow: 0 18px 40px rgba(15, 23, 42, .10); }
    code, pre { background: #f3f4f6; padding: 2px 6px; border-radius: 6px; }
    a { color: #0f766e; }
    .small { color: #4b5563; font-size: 0.95rem; }
  </style>
</head>
<body>
  <div class="card">
    <h1>PACT SharePoint Workbench Launcher</h1>
    <p>Redirecting to the SharePoint workbench with local debug manifests now.</p>
    <p class="small">If the redirect does not happen, open this URL:</p>
    <p><a href="$escapedWorkbenchUrl">$escapedWorkbenchUrl</a></p>
    <p class="small">Local manifests: <code>$escapedManifestsUrl</code></p>
    <p class="small">Workspace: <code>$escapedProjectRoot</code></p>
  </div>
</body>
</html>
"@
      $body = [System.Text.Encoding]::UTF8.GetBytes($html)
      Send-Response -Stream $stream -StatusCode 200 -StatusText 'OK' -ContentType 'text/html; charset=utf-8' -BodyBytes $body
      return
    }

    $filePath = Get-StaticFile -RelativePath $path
    if (-not $filePath) {
      $body = [System.Text.Encoding]::UTF8.GetBytes('Not Found')
      Send-Response -Stream $stream -StatusCode 404 -StatusText 'Not Found' -ContentType 'text/plain; charset=utf-8' -BodyBytes $body
      return
    }

    $fileName = [IO.Path]::GetFileName($filePath)
    if ($fileName -like 'pact-app-web-part*.js') {
      $bundleText = [IO.File]::ReadAllText($filePath)
      $bundleText = Rewrite-PactBundle -Content $bundleText
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($bundleText)
    } else {
      $bytes = [IO.File]::ReadAllBytes($filePath)
    }
    Send-Response -Stream $stream -StatusCode 200 -StatusText 'OK' -ContentType (Get-ContentType $filePath) -BodyBytes $bytes
  } catch {
    Write-Log "Request error: $($_.Exception.Message)"
    try {
      if ($stream) {
        $body = [System.Text.Encoding]::UTF8.GetBytes($_.Exception.Message)
        Send-Response -Stream $stream -StatusCode 500 -StatusText 'Internal Server Error' -ContentType 'text/plain; charset=utf-8' -BodyBytes $body
      }
    } catch {
      # ignore nested failures
    }
  } finally {
    try { if ($stream) { $stream.Close() } } catch { }
    try { $Client.Close() } catch { }
  }
}

try {
  Sync-DebugAssets
  Rewrite-DebugManifests
  $listener = [System.Net.Sockets.TcpListener]::Create($port)
  $listener.Start()
  Write-Log "Listening on http://127.0.0.1:$port/"
  Write-Host "PACT SPFx server listening on http://127.0.0.1:$port/"
  Write-Host "Redirecting root requests to SharePoint workbench."
  Write-Host "Workbench URL: $launchWorkbenchUrl"

  while ($true) {
    $client = $listener.AcceptTcpClient()
    Handle-Client -Client $client
  }
} catch {
  Write-Log "Startup failed: $($_.Exception.Message)"
  throw
}
