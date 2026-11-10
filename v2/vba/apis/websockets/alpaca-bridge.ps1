param(
    [string]   $Feed     = "test",
    [string[]] $Symbols  = @("FAKEPACA"),
    [string]   $OutFile,
    [string]   $StopFile
)

$Symbols = @(
    $Symbols |
        ForEach-Object { $_ -split ',' } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne '' }
)

$here = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if (-not $OutFile)  { $OutFile  = Join-Path $here "alpaca-stream.txt" }
if (-not $StopFile) { $StopFile = Join-Path $here "alpaca-stream.stop" }

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$keyId  = $env:ALPACA_KEY_ID
$secret = $env:ALPACA_SECRET_KEY

Write-Host "Alpaca bridge" -ForegroundColor Cyan
Write-Host "  feed    : $Feed"
Write-Host "  symbols : $($Symbols -join ', ')"
Write-Host "  writing : $OutFile"
Write-Host "  stop    : Ctrl+C here, or the STOP button in the workbook"
Write-Host ""

if (Test-Path $StopFile) { Remove-Item $StopFile -Force }
Set-Content -Path $OutFile -Value $null

$socket = [System.Net.WebSockets.ClientWebSocket]::new()
$token  = [Threading.CancellationToken]::None

function Send-Text([string] $text) {
    $bytes   = [Text.Encoding]::UTF8.GetBytes($text)
    $segment = [ArraySegment[byte]]::new($bytes)
    $socket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $token).Wait()
}

function Write-Line([string] $line) {
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            Add-Content -LiteralPath $OutFile -Value $line -ErrorAction Stop
            return
        }
        catch {
            if ($attempt -eq 5) { throw }
            Start-Sleep -Milliseconds 100
        }
    }
}

try {
    if (-not $keyId -or -not $secret) {
        throw "ALPACA_KEY_ID / ALPACA_SECRET_KEY are not set. Set them, then open a new window."
    }

    $socket.ConnectAsync([Uri]"wss://stream.data.alpaca.markets/v2/$Feed", $token).Wait()
    Write-Host "connected" -ForegroundColor Green

    Send-Text "{""action"":""auth"",""key"":""$keyId"",""secret"":""$secret""}"

    $symbolList = ($Symbols | ForEach-Object { '"' + $_ + '"' }) -join ','
    Send-Text "{""action"":""subscribe"",""quotes"":[$symbolList]}"

    $buffer       = [byte[]]::new(8192)
    $messageCount = 0

    while ($socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {

        if (Test-Path $StopFile) {
            Write-Host "stop requested" -ForegroundColor Yellow
            $socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "bye", $token).Wait()
            Remove-Item $StopFile -Force
            break
        }

        $text = [Text.StringBuilder]::new()
        do {
            $segment = [ArraySegment[byte]]::new($buffer)
            $result  = $socket.ReceiveAsync($segment, $token).GetAwaiter().GetResult()
            if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) { break }
            [void]$text.Append([Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count))
        } until ($result.EndOfMessage)

        if ($text.Length -gt 0) {
            Write-Line ($text.ToString())
            $messageCount++
            Write-Host "`rmessages received: $messageCount" -NoNewline
        }
    }
}
catch {
    $errorText = $_.Exception.Message
    Write-Warning "Bridge failed: $errorText"
    $errorEvent = ConvertTo-Json -InputObject @(@{
        T = 'error'; code = 'BRIDGE'; msg = $errorText
    }) -Compress
    try { Write-Line $errorEvent } catch { Write-Warning "Unable to write error event: $_" }
}
finally {
    try { Write-Line '[{"T":"bridge","msg":"stopped"}]' }
    catch { Write-Warning "Unable to write bridge-stopped event: $_" }
    if (Test-Path $StopFile) { Remove-Item -LiteralPath $StopFile -Force -ErrorAction SilentlyContinue }
    Write-Host ""
    Write-Host "bridge closed"
    $socket.Dispose()
}
