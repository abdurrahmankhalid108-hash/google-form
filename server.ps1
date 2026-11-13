# Local landing page. Serves public/ and emails through Gmail SMTP:
#   1) a copy of each note to you (NOTIFY_TO, or GMAIL_USER if empty), Reply-To = the visitor
#   2) a confirmation to the visitor
# Copy .env.example to .env and set GMAIL_USER plus a Google App Password first.
# Changes to .env are picked up on the next request; only PORT needs a restart.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$public = Join-Path $root "public"
$script:Hits = @{}
$script:AllowedOrigins = @()

function Read-DotEnv {
  $path = Join-Path $root ".env"
  $map = @{}
  if (-not (Test-Path -LiteralPath $path)) { return $map }
  foreach ($line in (Get-Content -LiteralPath $path -Encoding UTF8)) {
    $trimmed = $line.Trim()
    if ($trimmed -eq "" -or $trimmed.StartsWith("#")) { continue }
    $splitAt = $trimmed.IndexOf("=")
    if ($splitAt -lt 1) { continue }
    $key = $trimmed.Substring(0, $splitAt).Trim()
    $value = $trimmed.Substring($splitAt + 1).Trim().Trim('"').Trim("'")
    $map[$key] = $value
  }
  return $map
}

function Write-Bytes($response, $status, $contentType, [byte[]]$bytes) {
  $response.StatusCode = $status
  $response.ContentType = $contentType
  $response.Headers.Add("X-Content-Type-Options", "nosniff")
  $response.ContentLength64 = $bytes.Length
  $response.OutputStream.Write($bytes, 0, $bytes.Length)
  $response.OutputStream.Close()
  $response.Close()
}

function Write-Text($response, $status, $contentType, [string]$text) {
  $bytes = [Text.Encoding]::UTF8.GetBytes($text)
  Write-Bytes $response $status $contentType $bytes
}

function Write-Json($response, $status, $obj) {
  $json = $obj | ConvertTo-Json -Compress
  Write-Text $response $status "application/json; charset=utf-8" $json
}

function Get-Field($obj, [string]$name) {
  if ($null -eq $obj) { return "" }
  $prop = $obj.PSObject.Properties[$name]
  if ($null -eq $prop -or $null -eq $prop.Value) { return "" }
  return ([string]$prop.Value).Replace("`r", " ").Replace("`n", " ").Trim()
}

function Test-MailConfigured($envMap) {
  $user = [string]$envMap["GMAIL_USER"]
  $pass = ([string]$envMap["GMAIL_APP_PASSWORD"]) -replace "\s", ""
  return ($user -match '@') -and ($pass.Length -ge 8) -and ($user -notmatch 'your\.name@gmail\.com')
}

function Get-NotifyAddress($envMap) {
  $to = ([string]$envMap["NOTIFY_TO"]).Trim()
  if ($to -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $to }
  return [string]$envMap["GMAIL_USER"]
}

function Hide-Email([string]$email) {
  $at = $email.IndexOf("@")
  if ($at -lt 1) { return "***" }
  return $email.Substring(0, 1) + "***" + $email.Substring($at)
}

# --- Rate limiting: check first, record only when a send is actually attempted ---
function Get-RecentHits([string]$key, [int]$minutes) {
  if (-not $script:Hits.ContainsKey($key)) { return @() }
  $window = [DateTime]::UtcNow.AddMinutes(-$minutes)
  $recent = @($script:Hits[$key] | Where-Object { $_ -gt $window })
  $script:Hits[$key] = $recent
  return $recent
}

function Test-RateLimit([string]$key, [int]$max, [int]$minutes) {
  return (@(Get-RecentHits $key $minutes).Count -lt $max)
}

function Add-Hit([string]$key) {
  if (-not $script:Hits.ContainsKey($key)) { $script:Hits[$key] = @() }
  $script:Hits[$key] = @($script:Hits[$key]) + @([DateTime]::UtcNow)
}

function Send-Mail($envMap, [string]$to, [string]$subject, [string]$body, [string]$replyTo) {
  $fromAddress = [string]$envMap["GMAIL_USER"]
  $password = ([string]$envMap["GMAIL_APP_PASSWORD"]) -replace "\s", ""
  $mail = New-Object System.Net.Mail.MailMessage
  $smtp = New-Object System.Net.Mail.SmtpClient("smtp.gmail.com", 587)
  try {
    $mail.From = New-Object System.Net.Mail.MailAddress($fromAddress, "Lumen Post")
    $mail.To.Add($to)
    if ($replyTo -ne "") { $mail.ReplyToList.Add($replyTo) }
    $mail.Subject = $subject
    $mail.Body = $body
    $mail.BodyEncoding = [Text.Encoding]::UTF8
    $mail.SubjectEncoding = [Text.Encoding]::UTF8
    $smtp.EnableSsl = $true
    $smtp.Timeout = 20000
    $smtp.Credentials = New-Object System.Net.NetworkCredential($fromAddress, $password)
    $smtp.Send($mail)
  } finally {
    $mail.Dispose()
    $smtp.Dispose()
  }
}

function Send-OwnerCopy($envMap, $name, $email, $topic, $message) {
  $nl = "`r`n"
  $body = "New note from the Lumen Post form." + $nl + $nl +
    "Name: " + $name + $nl +
    "Email: " + $email + $nl +
    "Topic: " + $topic + $nl + $nl +
    "Message:" + $nl + $message + $nl + $nl +
    "Reply to this email to answer them directly." + $nl
  Send-Mail $envMap (Get-NotifyAddress $envMap) ("New note from " + $name + ": " + $topic) $body $email
}

function Send-Confirmation($envMap, $name, $email, $topic, $message) {
  $nl = "`r`n"
  $body = "Hi " + $name + "," + $nl + $nl +
    "This confirms we received your form on Lumen Post." + $nl + $nl +
    "Topic: " + $topic + $nl + $nl +
    "Your message:" + $nl + $message + $nl + $nl +
    "We will reply to this Gmail address. If you did not submit the form, you can ignore this note." + $nl + $nl +
    "Lumen Post" + $nl
  Send-Mail $envMap $email "Confirmation: we received your note" $body ""
}

function Resolve-PublicFile([string]$rawPath) {
  try {
    $relative = [Uri]::UnescapeDataString(($rawPath -split "\?")[0]).TrimStart("/")
    if ([string]::IsNullOrWhiteSpace($relative)) { $relative = "index.html" }
    $relative = $relative -replace "/", [IO.Path]::DirectorySeparatorChar
    $full = [IO.Path]::GetFullPath((Join-Path $public $relative))
    $sep = [string][IO.Path]::DirectorySeparatorChar
    $publicRoot = [IO.Path]::GetFullPath($public)
    if (-not $publicRoot.EndsWith($sep)) { $publicRoot = $publicRoot + $sep }
    $check = $full
    if (-not $check.EndsWith($sep)) { $check = $check + $sep }
    if (-not $check.StartsWith($publicRoot, [StringComparison]::OrdinalIgnoreCase)) { return $null }
    if (Test-Path -LiteralPath $full -PathType Container) { $full = Join-Path $full "index.html" }
    return $full
  } catch {
    return $null
  }
}

function Get-ContentType([string]$path) {
  switch ([IO.Path]::GetExtension($path).ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8" }
    ".css"  { "text/css; charset=utf-8" }
    ".js"   { "text/javascript; charset=utf-8" }
    ".json" { "application/json; charset=utf-8" }
    ".txt"  { "text/plain; charset=utf-8" }
    ".svg"  { "image/svg+xml" }
    ".png"  { "image/png" }
    ".jpg"  { "image/jpeg" }
    ".jpeg" { "image/jpeg" }
    ".gif"  { "image/gif" }
    ".webp" { "image/webp" }
    ".ico"  { "image/x-icon" }
    ".woff" { "font/woff" }
    ".woff2" { "font/woff2" }
    default { "application/octet-stream" }
  }
}

function Handle-Submit($response, $request) {
  # Only our own page (or a non-browser tool such as curl) may post here.
  $origin = [string]$request.Headers["Origin"]
  if ($origin -ne "" -and $script:AllowedOrigins -notcontains $origin.ToLowerInvariant()) {
    Write-Json $response 403 @{ ok = $false; error = "This request was not allowed." }
    return
  }
  if (([string]$request.ContentType) -notmatch '^application/json\b') {
    Write-Json $response 415 @{ ok = $false; error = "The form could not be read." }
    return
  }

  $length = [int64]$request.ContentLength64
  if ($length -lt 0) {
    Write-Json $response 411 @{ ok = $false; error = "The form could not be read." }
    return
  }
  if ($length -gt 20000) {
    Write-Json $response 413 @{ ok = $false; error = "That note is too long." }
    return
  }
  $buffer = New-Object byte[] $length
  $read = 0
  while ($read -lt $length) {
    $n = $request.InputStream.Read($buffer, $read, [int]($length - $read))
    if ($n -le 0) { break }
    $read += $n
  }
  $raw = [Text.Encoding]::UTF8.GetString($buffer, 0, $read)
  try { $data = $raw | ConvertFrom-Json } catch {
    Write-Json $response 400 @{ ok = $false; error = "The form could not be read." }
    return
  }
  if ($data -isnot [pscustomobject]) {
    Write-Json $response 400 @{ ok = $false; error = "The form could not be read." }
    return
  }

  # Honeypot: real visitors never see this field. If it is filled (bot, or a stray autofill),
  # send nothing and say so, so a real person is never told a confirmation is on its way.
  $trap = Get-Field $data "leave_blank"
  if ($trap -ne "") {
    Write-Host "Honeypot tripped; submission ignored."
    Write-Json $response 400 @{ ok = $false; error = "The form could not be sent. Reload the page and try again." }
    return
  }

  $name = Get-Field $data "name"
  $email = (Get-Field $data "email").ToLowerInvariant()
  $topic = Get-Field $data "topic"
  $message = ""
  $messageProp = $data.PSObject.Properties["message"]
  if ($null -ne $messageProp -and $null -ne $messageProp.Value) {
    $message = ([string]$messageProp.Value).Trim()
  }

  $allowedTopics = @("Starting a project", "A question", "Working together", "Something else")
  if ($name.Length -lt 2) {
    Write-Json $response 400 @{ ok = $false; error = "Add your name." }
    return
  }
  if ($name.Length -gt 80) {
    Write-Json $response 400 @{ ok = $false; error = "That name is too long (80 characters max)." }
    return
  }
  if ($email.Length -gt 254 -or $email -notmatch '^[a-z0-9._%+-]+@(gmail|googlemail)\.com$') {
    Write-Json $response 400 @{ ok = $false; error = "Enter a Gmail address, like name@gmail.com." }
    return
  }
  if ($allowedTopics -notcontains $topic) {
    Write-Json $response 400 @{ ok = $false; error = "Choose a topic." }
    return
  }
  if ($message.Length -lt 8) {
    Write-Json $response 400 @{ ok = $false; error = "Add a short message (at least 8 characters)." }
    return
  }
  if ($message.Length -gt 1000) {
    Write-Json $response 400 @{ ok = $false; error = "That message is too long (1000 characters max)." }
    return
  }

  $envMap = Read-DotEnv
  if (-not (Test-MailConfigured $envMap)) {
    Write-Json $response 503 @{ ok = $false; error = "Gmail sending is not set up yet. Add GMAIL_USER and GMAIL_APP_PASSWORD to the .env file and save it." }
    return
  }

  # Rate limits apply only to valid submissions. Note: on a localhost listener every
  # visitor shares one address, so the first limit is effectively site-wide.
  $ipKey = "ip:" + $request.RemoteEndPoint.Address.ToString()
  $toKey = "to:" + $email
  if (-not (Test-RateLimit $ipKey 5 10)) {
    Write-Json $response 429 @{ ok = $false; error = "Too many submissions right now. Wait a few minutes and try again." }
    return
  }
  if (-not (Test-RateLimit $toKey 2 10)) {
    Write-Json $response 429 @{ ok = $false; error = "A confirmation was already sent to that address. Check the inbox, or wait a few minutes." }
    return
  }
  Add-Hit $ipKey
  Add-Hit $toKey

  # 1) Deliver the note to the owner first, so a submission is never lost.
  try {
    Send-OwnerCopy $envMap $name $email $topic $message
  } catch {
    Write-Host ("Owner copy failed: " + $_.Exception.Message)
    Write-Json $response 502 @{ ok = $false; error = "The note could not be sent. Check the Gmail app password and try again." }
    return
  }

  # 2) Confirm to the visitor. If this fails the owner still has the note.
  $confirmed = $true
  try {
    Send-Confirmation $envMap $name $email $topic $message
  } catch {
    $confirmed = $false
    Write-Host ("Confirmation failed: " + $_.Exception.Message)
  }

  Write-Host ("Note received from " + (Hide-Email $email) + "; confirmation sent: " + $confirmed)
  Write-Json $response 200 @{ ok = $true; confirmed = $confirmed }
}

$envMap = Read-DotEnv
$port = 8787
if ($envMap.ContainsKey("PORT")) {
  if ($envMap["PORT"] -match '^\d{1,5}$' -and [int]$envMap["PORT"] -ge 1 -and [int]$envMap["PORT"] -le 65535) {
    $port = [int]$envMap["PORT"]
  } else {
    Write-Host "PORT in .env must be a number from 1 to 65535. Using 8787."
  }
}
$script:AllowedOrigins = @("http://localhost:$port", "http://127.0.0.1:$port")

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
try {
  # Drop clients that stall mid-request so one slow connection cannot freeze the server.
  $listener.TimeoutManager.HeaderWait = [TimeSpan]::FromSeconds(15)
  $listener.TimeoutManager.EntityBody = [TimeSpan]::FromSeconds(15)
} catch {}
$listener.Start()
Write-Host "Lumen Post is running at http://localhost:$port/"
if (Test-MailConfigured $envMap) {
  Write-Host "Gmail sending is configured."
} else {
  Write-Host "Gmail sending is not configured. Copy .env.example to .env and add an app password."
}

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response
    try {
      $path = $request.Url.AbsolutePath
      Write-Host ($request.HttpMethod + " " + $path)
      if ($request.HttpMethod -eq "GET" -and $path -eq "/api/health") {
        $live = Read-DotEnv
        Write-Json $response 200 @{ ok = $true; mailConfigured = [bool](Test-MailConfigured $live) }
        continue
      }
      if ($request.HttpMethod -eq "POST" -and $path -eq "/api/submit") {
        Handle-Submit $response $request
        continue
      }
      if ($request.HttpMethod -ne "GET") {
        Write-Text $response 405 "text/plain; charset=utf-8" "Method not allowed."
        continue
      }
      $file = Resolve-PublicFile $path
      if ($null -eq $file -or -not (Test-Path -LiteralPath $file -PathType Leaf)) {
        Write-Text $response 404 "text/html; charset=utf-8" "<!DOCTYPE html><title>Not found</title><p>That page is not here.</p>"
        continue
      }
      Write-Bytes $response 200 (Get-ContentType $file) ([IO.File]::ReadAllBytes($file))
    } catch {
      Write-Host ("Request failed: " + $_.Exception.Message)
      try { Write-Text $response 500 "application/json; charset=utf-8" '{"ok":false,"error":"Something went wrong."}' } catch {}
    }
  }
} finally {
  $listener.Stop()
  $listener.Close()
}
