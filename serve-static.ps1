param(
  [Parameter(Mandatory = $true)]
  [string]$Root,

  [Parameter(Mandatory = $true)]
  [int]$Port,

  [Parameter(Mandatory = $true)]
  [string]$LogPath
)

$ErrorActionPreference = 'Stop'

function Get-MimeType {
  param([string]$Path)

  switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
    '.html' { 'text/html; charset=utf-8' }
    '.htm' { 'text/html; charset=utf-8' }
    '.css' { 'text/css; charset=utf-8' }
    '.js' { 'application/javascript; charset=utf-8' }
    '.mjs' { 'application/javascript; charset=utf-8' }
    '.json' { 'application/json; charset=utf-8' }
    '.svg' { 'image/svg+xml' }
    '.png' { 'image/png' }
    '.jpg' { 'image/jpeg' }
    '.jpeg' { 'image/jpeg' }
    '.ico' { 'image/x-icon' }
    '.txt' { 'text/plain; charset=utf-8' }
    default { 'application/octet-stream' }
  }
}

function Resolve-ResponseFile {
  param(
    [string]$RequestPath,
    [string]$RootPath
  )

  $relative = [Uri]::UnescapeDataString($RequestPath.TrimStart('/'))
  if ([string]::IsNullOrWhiteSpace($relative)) {
    return (Join-Path $RootPath 'index.html')
  }

  $candidate = Join-Path $RootPath $relative
  if (Test-Path $candidate -PathType Leaf) {
    return $candidate
  }

  if (Test-Path $candidate -PathType Container) {
    $indexCandidate = Join-Path $candidate 'index.html'
    if (Test-Path $indexCandidate -PathType Leaf) {
      return $indexCandidate
    }
  }

  return (Join-Path $RootPath 'index.html')
}

Add-Content -Path $LogPath -Value "Starting static server on http://127.0.0.1:$Port/ serving $Root"

$listener = [System.Net.Sockets.TcpListener]::Create($Port)
$listener.Start()

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $stream = $client.GetStream()
      $buffer = New-Object byte[] 8192
      $requestText = ''

      while ($true) {
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { break }
        $requestText += [System.Text.Encoding]::ASCII.GetString($buffer, 0, $read)
        if ($requestText.Contains("`r`n`r`n")) { break }
      }

      $requestLine = $requestText.Split("`r`n")[0]
      $parts = $requestLine.Split(' ')
      $path = if ($parts.Length -ge 2) { $parts[1] } else { '/' }
      $fullPath = Resolve-ResponseFile -RequestPath $path -RootPath $Root
      $bytes = [System.IO.File]::ReadAllBytes($fullPath)
      $contentType = Get-MimeType -Path $fullPath
      $statusLine = "HTTP/1.1 200 OK`r`n"
      $headers = @(
        "Content-Type: $contentType"
        "Content-Length: $($bytes.Length)"
        "Connection: close"
      ) -join "`r`n"
      $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($statusLine + $headers + "`r`n`r`n")
      $stream.Write($responseBytes, 0, $responseBytes.Length)
      $stream.Write($bytes, 0, $bytes.Length)
      $stream.Flush()
    }
    catch {
      Add-Content -Path $LogPath -Value ("Request error: " + $_.Exception.Message)
    }
    finally {
      if ($client) { $client.Close() }
    }
  }
}
finally {
  $listener.Stop()
}
