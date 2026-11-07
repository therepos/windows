# irm "https://github.com/therepos/windows/raw/main/apps/agentprofiler.ps1?$(Get-Random)" | iex
# Purpose: Test what corporate monitoring agents can see during local vs internet file transfers
# =============================================================================
# Note:
#   Runs once, no menu: which agents react, and what each can capture (IO spike != content inspection)
#   Output: monitoring_test_log_<timestamp>.txt on the Desktop
# =============================================================================

$ScriptUrl = 'https://github.com/therepos/windows/raw/main/apps/agentprofiler.ps1'

# Admin is needed to read IO counters and owning PIDs of security agents
# >>> elevate
function Test-Admin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Start-Elevated([string]$ScriptUrl, [switch]$NoExit) {
    Write-Host "  Requesting administrator rights..." -ForegroundColor Yellow
    $run = if ($PSCommandPath) { "& '$($PSCommandPath -replace "'", "''")'" } else { "irm '${ScriptUrl}?$(Get-Random)' | iex" }
    $cmd = "[Net.ServicePointManager]::SecurityProtocol = 'Tls12'; try { $run } catch { Write-Host `$_ -ForegroundColor Red; Read-Host 'Press Enter to exit' }"
    # Encoded, so Windows Terminal does not split the command at ';' or mangle quotes
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass')
    if ($NoExit) { $psArgs += '-NoExit' }
    $psArgs += @('-EncodedCommand', $enc)
    $opened = $false
    if (Get-Command wt.exe -ErrorAction SilentlyContinue) {
        try {
            Start-Process -FilePath 'wt.exe' -Verb RunAs -ArgumentList (@('powershell.exe') + $psArgs) -ErrorAction Stop
            $opened = $true
        } catch {
            # 'No' on the admin prompt: do not ask again for the classic window
            $e = $_.Exception
            while ($e.InnerException) { $e = $e.InnerException }
            if ($e -is [ComponentModel.Win32Exception] -and $e.NativeErrorCode -eq 1223) {
                Write-Host "  Cancelled." -ForegroundColor Yellow
                return
            }
        }
    }
    if (-not $opened) { Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $psArgs }
    Write-Host "  Continue in the new (elevated) window. You can close this one." -ForegroundColor Gray
}
# <<< elevate
$IsAdmin = Test-Admin
if (-not $IsAdmin) { Start-Elevated $ScriptUrl -NoExit; return }

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

# =============================================================================
#  AGENT KNOWLEDGE BASE
#  What each tool actually does vs what people assume it does.
# =============================================================================
$AgentProfiles = [ordered]@{
    'zsatunnel.exe' = @{
        Label = 'Zscaler Tunnel'; Category = 'Network Security'
        Purpose = "Cloud proxy - routes internet traffic through Zscaler's cloud for inspection"
        CanSee = @(
            'URLs and domains you visit',
            'File uploads/downloads over HTTP/HTTPS (via SSL interception)',
            'File names, sizes, and destinations',
            'File CONTENTS if SSL interception is active',
            'Which cloud apps you use and for how long')
        CannotSee = @(
            'Files copied to USB or local drives',
            'What you type (no keystroke logging)',
            'Your screen (no screen capture)',
            'Clipboard content',
            'Files that never touch the network')
        LocalTrafficNote = "Zscaler's agent hooks into the OS network stack via a TUN/TAP driver. It MAY see local network packets if its filter driver intercepts all traffic before routing decisions. However, local traffic is typically NOT forwarded to Zscaler's cloud for inspection - the agent may log the connection metadata (source, dest, port) locally but is unlikely to perform content inspection on LAN-to-LAN transfers."
        IoSpikeMeaning = "An IO spike from ZSATunnel during a local transfer likely means the agent's network filter SAW the packets pass through the stack and logged metadata (IP, port, bytes). This does NOT necessarily mean it inspected the content or forwarded it to the cloud. For internet transfers, it DOES inspect and forward content."
    }
    'zscaler.exe' = @{
        Label = 'Zscaler App'; Category = 'Network Security'
        Purpose = 'Zscaler client UI and management process'
        CanSee = @('Same as ZSATunnel - this is the parent process')
        CannotSee = @('Same as ZSATunnel')
        LocalTrafficNote = 'See ZSATunnel notes.'
        IoSpikeMeaning = 'UI/management process - IO often unrelated to traffic inspection.'
    }
    'zsatray.exe' = @{
        Label = 'Zscaler Tray'; Category = 'Network Security'
        Purpose = 'System tray icon for Zscaler'
        CanSee = @('Minimal - mostly UI status')
        CannotSee = @('Content of any transfers')
        LocalTrafficNote = 'UI only.'
        IoSpikeMeaning = 'Likely just refreshing status display.'
    }
    'mssense.exe' = @{
        Label = 'Defender for Endpoint (EDR)'; Category = 'Endpoint Security (EDR)'
        Purpose = 'Advanced threat detection - monitors process behavior, file access, network connections'
        CanSee = @(
            'Which processes access which files (file path, name, size)',
            'Process creation chains (what launched what)',
            'Network connections made by each process (IP, port, protocol)',
            'Suspicious behavior patterns (e.g., mass file reads, encryption)',
            'Registry changes, privilege escalation attempts',
            'DLL loads and code injection attempts')
        CannotSee = @(
            'File CONTENTS (it detects behavior patterns, not reads file text)',
            'Clipboard content',
            'Keystrokes',
            'Screen content',
            'The actual data bytes transferred over the network')
        LocalTrafficNote = "MsSense monitors at the endpoint level - it sees that a process opened a network socket and transferred X bytes to Y IP. It logs this regardless of whether the destination is local or internet. However, it does NOT inspect the content of the transfer. It looks for behavioral anomalies (e.g., 'python.exe suddenly reading 500 files and opening a network connection')."
        IoSpikeMeaning = "An IO spike means MsSense recorded telemetry about what processes did - file opens, network connections, process activity. This is metadata/behavioral logging, NOT content inspection. It knows 'python.exe sent 50KB to 192.168.1.5' but NOT what was in those 50KB."
    }
    'msmpeng.exe' = @{
        Label = 'Windows Defender Antimalware'; Category = 'Antivirus'
        Purpose = 'Real-time antimalware scanning - checks files for malware signatures'
        CanSee = @(
            'File contents (scans for malware patterns/signatures)',
            'Files as they are created, modified, or accessed',
            'Downloaded files',
            'Email attachments opened locally')
        CannotSee = @(
            'Network traffic content (not a network inspector)',
            'Clipboard content',
            'Keystrokes or screen content',
            "What you do in your browser (that's Zscaler's domain)")
        LocalTrafficNote = "MsMpEng does NOT monitor network traffic. It scans FILES on disk. If it spiked during a transfer test, it's because the test file was written to disk and Defender scanned it for malware - not because it inspected the network transfer itself."
        IoSpikeMeaning = 'An IO spike means Defender scanned a file that was created or accessed. It reads the file to check for malware signatures. It is NOT inspecting network traffic or logging your transfer activity. It would spike the same way if you simply opened the file in Notepad.'
    }
    'defendpointservice.exe' = @{
        Label = 'BeyondTrust / Defendpoint'; Category = 'Privilege Management'
        Purpose = 'Controls which apps can run with elevated privileges, application whitelisting'
        CanSee = @(
            'Which applications are running',
            'Whether an app requested admin/elevated privileges',
            'Application install/uninstall events')
        CannotSee = @(
            'File contents',
            'Network traffic or transfer content',
            'Clipboard, keystrokes, or screen content',
            'What data you send or receive')
        LocalTrafficNote = 'BeyondTrust does not monitor network traffic at all. It manages application privileges and policies.'
        IoSpikeMeaning = 'IO spike likely means it checked whether the running process (python.exe) is allowed to run or has the correct privilege level. This is routine application control, not surveillance.'
    }
    'nxtcoordinator.exe' = @{
        Label = 'Nexthink Coordinator'; Category = 'Digital Experience Management (DEX)'
        Purpose = 'IT operations - measures app performance, device health, user experience'
        CanSee = @(
            'Application names and usage duration (how long you used each app)',
            'Device health: CPU, RAM, disk, battery',
            'Network performance: latency, throughput, connection quality',
            'Software inventory and versions installed',
            'App crash reports and error counts',
            'Boot/login times')
        CannotSee = @(
            'File contents or file names',
            'Network traffic content',
            'Clipboard content',
            'Keystrokes or screen content',
            'What you upload or download (only that an app used network)',
            'Emails, chat messages, or documents')
        LocalTrafficNote = 'Nexthink does NOT inspect network traffic content. It may log that an application used X MB of network bandwidth, but it does not know what the data was. It is an IT experience tool, not a security tool.'
        IoSpikeMeaning = "IO spike means Nexthink collected its routine telemetry - app usage, performance metrics, maybe network throughput stats. This is NOT content inspection. It would record 'python.exe used 50KB of network' but has no visibility into what that 50KB contained."
    }
    'ccmexec.exe' = @{
        Label = 'SCCM / ConfigMgr Client'; Category = 'IT Management'
        Purpose = 'Software deployment, patch management, hardware/software inventory, compliance'
        CanSee = @(
            'Software installed on the machine',
            'Patch/update compliance status',
            'Hardware inventory (CPU, RAM, disk model)',
            'Whether required software is present')
        CannotSee = @(
            'File contents or user documents',
            'Network traffic content',
            'Clipboard, keystrokes, screen',
            'What you do day-to-day (not a monitoring tool)')
        LocalTrafficNote = 'SCCM does not monitor network traffic or file transfers.'
        IoSpikeMeaning = 'IO spike likely means SCCM ran a scheduled inventory scan or checked for pending software updates. Completely unrelated to your transfer.'
    }
    'smartscreen.exe' = @{
        Label = 'Windows SmartScreen'; Category = 'Application Reputation'
        Purpose = "Checks downloaded files and apps against Microsoft's reputation database"
        CanSee = @(
            'Executable files you download or run for the first time',
            'URLs visited in Edge browser')
        CannotSee = @(
            'Document contents (only checks executables)',
            'Network traffic content',
            'Clipboard, keystrokes, screen',
            'File transfers between devices')
        LocalTrafficNote = 'SmartScreen does not monitor file transfers.'
        IoSpikeMeaning = 'Likely checked whether python.exe is a known/safe application.'
    }
}

$DefaultProfile = @{
    Label = 'Unknown Agent'; Category = 'Unknown'
    Purpose = 'Not in our knowledge base'
    CanSee = @('Unknown - research this process name')
    CannotSee = @('Unknown')
    LocalTrafficNote = 'Unknown - investigate manually.'
    IoSpikeMeaning = 'Cannot determine without identifying the process.'
}

function Get-Profile($name) {
    if ($AgentProfiles.Contains($name)) { return $AgentProfiles[$name] }
    return $DefaultProfile
}

# =============================================================================
#  SNAPSHOT ENGINE
# =============================================================================
function Get-AgentIO {
    $snapshot = @{}
    Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | ForEach-Object {
        $pname = $_.Name.ToLower()
        if ($AgentProfiles.Contains($pname)) {
            $snapshot[$pname] = @{
                Pid        = $_.ProcessId
                ReadBytes  = [double]$_.ReadTransferCount
                WriteBytes = [double]$_.WriteTransferCount
                ReadCount  = [double]$_.ReadOperationCount
                WriteCount = [double]$_.WriteOperationCount
            }
        }
    }
    return $snapshot
}

function Get-Connections {
    $conns = @{}
    Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.RemoteAddress -and $_.RemoteAddress -notin @('0.0.0.0', '::')) {
            $proc = (Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue).Name
            if ($proc) {
                $key = "$proc|$($_.RemoteAddress):$($_.RemotePort)"
                $conns[$key] = @{ Process = $proc; RemoteIp = $_.RemoteAddress; RemotePort = $_.RemotePort }
            }
        }
    }
    return $conns
}

function Get-IODiff($before, $after) {
    $deltas = @{}
    foreach ($pname in $after.Keys) {
        if ($before.ContainsKey($pname)) {
            $rd = $after[$pname].ReadBytes  - $before[$pname].ReadBytes
            $wd = $after[$pname].WriteBytes - $before[$pname].WriteBytes
            $rc = $after[$pname].ReadCount  - $before[$pname].ReadCount
            $wc = $after[$pname].WriteCount - $before[$pname].WriteCount
            $p = Get-Profile $pname
            $deltas[$pname] = @{
                Label   = $p.Label
                Category = $p.Category
                ReadKb  = [math]::Round($rd / 1024, 1)
                WriteKb = [math]::Round($wd / 1024, 1)
                ReadOps = $rc
                WriteOps = $wc
                TotalKb = [math]::Round(($rd + $wd) / 1024, 1)
            }
        }
    }
    return $deltas
}

function Get-ConnDiff($before, $after) {
    $new = @{}
    foreach ($k in $after.Keys) { if (-not $before.ContainsKey($k)) { $new[$k] = $after[$k] } }
    return $new
}

function Get-LocalIP {
    try {
        $s = New-Object System.Net.Sockets.Socket('InterNetwork', 'Dgram', 'Udp')
        $s.Connect('8.8.8.8', 80)
        $ip = ([System.Net.IPEndPoint]$s.LocalEndPoint).Address.ToString()
        $s.Close()
        return $ip
    } catch { return '127.0.0.1' }
}

# =============================================================================
#  LOGGER - writes to both console and log file
# =============================================================================
$script:LogLines = New-Object System.Collections.Generic.List[string]
function Log($text = '') {
    Write-Host $text
    $script:LogLines.Add($text)
}
function LogHeader($title) {
    $w = 64
    Log ''
    Log ('=' * $w)
    Log "  $title"
    Log ('=' * $w)
}
function LogSeparator { Log ''; Log ('-' * 64) }

# =============================================================================
#  LOCAL NETWORK SERVER / CLIENT
# =============================================================================
function New-TestFile($dir) {
    $path = Join-Path $dir 'CONFIDENTIAL_Client_Data_Test.txt'
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('PRIVILEGED & CONFIDENTIAL')
    [void]$sb.AppendLine('Client: Acme Holdings Pte Ltd')
    [void]$sb.AppendLine('Engagement: SG-2026-0045')
    [void]$sb.AppendLine('SSN: 078-05-1120')
    [void]$sb.AppendLine('Credit Card: 4111-1111-1111-1111')
    [void]$sb.AppendLine('Wire: ABA 021000021 / Acct 123456789')
    [void]$sb.AppendLine('MNPI - Material Non-Public Information')
    [void]$sb.AppendLine('')
    for ($i = 0; $i -lt 500; $i++) {
        $amt = [math]::Round(($i * 347.89) % 45000, 2)
        [void]$sb.AppendLine(("Row {0}: Revenue `${1:N0} | AR `${2:N0} | Vendor V{3} | INV-2025-{4} | `${5:N2}" -f `
            $i, ($i * 1000), ($i * 500), (10000 + $i), (8000 + $i), $amt))
    }
    [System.IO.File]::WriteAllText($path, $sb.ToString())
    return $path
}

# Runs a TCP receiver in a background runspace; shares state via a synchronized hashtable.
function Start-LocalServer($bindAddr, $port, $sync) {
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('bindAddr', $bindAddr)
    $rs.SessionStateProxy.SetVariable('port', $port)
    $rs.SessionStateProxy.SetVariable('sync', $sync)
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        try {
            $addr = if ($bindAddr -eq '0.0.0.0') { [System.Net.IPAddress]::Any } else { [System.Net.IPAddress]::Parse($bindAddr) }
            $listener = New-Object System.Net.Sockets.TcpListener($addr, $port)
            $listener.Start()
            $sync['ready'] = $true
            $iar = $listener.BeginAcceptTcpClient($null, $null)
            if ($iar.AsyncWaitHandle.WaitOne(30000)) {
                $client = $listener.EndAcceptTcpClient($iar)
                $stream = $client.GetStream()
                $stream.ReadTimeout = 30000
                $buf = New-Object byte[] 65536
                $total = 0
                while (($n = $stream.Read($buf, 0, $buf.Length)) -gt 0) { $total += $n }
                $client.Close()
                $sync['received'] = $total
            } else { $sync['received'] = 0 }
            $listener.Stop()
        } catch { $sync['received'] = 0 }
        $sync['done'] = $true
    })
    $handle = $ps.BeginInvoke()
    return @{ PS = $ps; Handle = $handle; RS = $rs }
}

function Send-FileTo($filepath, $hostName, $port) {
    $client = New-Object System.Net.Sockets.TcpClient
    $client.Connect($hostName, $port)
    $stream = $client.GetStream()
    $fs = [System.IO.File]::OpenRead($filepath)
    $buf = New-Object byte[] 65536
    while (($n = $fs.Read($buf, 0, $buf.Length)) -gt 0) { $stream.Write($buf, 0, $n) }
    $stream.Flush()
    $fs.Close()
    $client.Close()
}

# =============================================================================
#  RUN A SINGLE TEST
# =============================================================================
function Invoke-TransferTest($testFunc, $waitSeconds = 8) {
    $ioBefore = Get-AgentIO
    $connBefore = Get-Connections
    & $testFunc
    if ($waitSeconds -gt 0) {
        Log "  Waiting for agent reaction ($waitSeconds sec)..."
        Start-Sleep -Seconds $waitSeconds
    }
    $ioAfter = Get-AgentIO
    $connAfter = Get-Connections
    return @{
        Deltas = (Get-IODiff $ioBefore $ioAfter)
        NewConns = (Get-ConnDiff $connBefore $connAfter)
    }
}

# =============================================================================
#  MAIN
# =============================================================================
$Desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($Desktop) -or -not (Test-Path $Desktop)) { $Desktop = $env:USERPROFILE }
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $Desktop "monitoring_test_log_$timestamp.txt"

Log ''
Log '=================================================================='
Log '     LOCAL vs INTERNET MONITORING TEST'
Log '     What exactly can each agent see?'
Log '=================================================================='
Log ''
Log ("  Timestamp: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Log ("  Computer: {0}" -f $env:COMPUTERNAME)
Log ("  User: {0}" -f $env:USERNAME)

# -- Detect agents ------------------------------------------------------------
LogHeader 'STEP 1: DETECTING MONITORING AGENTS'
$agents = Get-AgentIO
if ($agents.Count -eq 0) {
    Log '  [!] No known monitoring agents detected.'
} else {
    Log ("  Found {0} agents:" -f $agents.Count)
    Log ''
    foreach ($pname in $agents.Keys) {
        $p = Get-Profile $pname
        Log ("    - {0} ({1}, PID {2})" -f $p.Label, $pname, $agents[$pname].Pid)
        Log ("      Category: {0}" -f $p.Category)
        Log ("      Purpose: {0}" -f $p.Purpose)
    }
}

$localIp = Get-LocalIP
Log ''
Log ("  Local IP: {0}" -f $localIp)

$testDir = Join-Path ([System.IO.Path]::GetTempPath()) ("montest_" + [System.IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $testDir -Force | Out-Null
$testFile = New-TestFile $testDir
$fileSize = (Get-Item $testFile).Length
Log ("  Test file: {0:N0} bytes (with DLP trigger content)" -f $fileSize)

$results = @{}

# -- TEST A: BASELINE ---------------------------------------------------------
LogHeader 'TEST A: BASELINE - idle for 10 seconds'
Log '  Measuring background IO noise. No activity.'
Log ''
$r = Invoke-TransferTest { Start-Sleep -Seconds 10 } 0
$baseline = $r.Deltas
$results['baseline'] = $baseline
foreach ($pname in $baseline.Keys) {
    $d = $baseline[$pname]
    Log ("    {0,-30} total: {1,8:N1} KB" -f $d.Label, $d.TotalKb)
}

# -- TEST B: LOCALHOST (127.0.0.1) --------------------------------------------
LogHeader 'TEST B: LOCALHOST TRANSFER (127.0.0.1)'
Log '  Loopback transfer - never touches network hardware.'
Log ''
$r = Invoke-TransferTest {
    $port = 19876
    $sync = [hashtable]::Synchronized(@{ ready = $false; done = $false; received = 0 })
    $srv = Start-LocalServer '127.0.0.1' $port $sync
    $t = 0; while (-not $sync['ready'] -and $t -lt 50) { Start-Sleep -Milliseconds 100; $t++ }
    Send-FileTo $testFile '127.0.0.1' $port
    $t = 0; while (-not $sync['done'] -and $t -lt 100) { Start-Sleep -Milliseconds 100; $t++ }
    $srv.PS.Dispose(); $srv.RS.Dispose()
    Log ("  Transferred {0:N0} bytes over localhost" -f $sync['received'])
}
$results['localhost'] = $r.Deltas
foreach ($pname in $r.Deltas.Keys) {
    $d = $r.Deltas[$pname]
    $bl = if ($baseline.ContainsKey($pname)) { $baseline[$pname].TotalKb } else { 0 }
    $above = $d.TotalKb - $bl
    $flag = if ($above -gt 5) { '[ACTIVITY]' } else { '   quiet' }
    Log ("    {0,-30} total: {1,8:N1} KB  (+{2:N1} vs baseline) {3}" -f $d.Label, $d.TotalKb, $above, $flag)
}
Start-Sleep -Seconds 3

# -- TEST C: LAN IP -----------------------------------------------------------
LogHeader ("TEST C: LAN TRANSFER (via {0})" -f $localIp)
Log '  Simulates LocalSend - traffic goes through network stack.'
Log ''
$r = Invoke-TransferTest {
    $port = 19877
    $sync = [hashtable]::Synchronized(@{ ready = $false; done = $false; received = 0 })
    $srv = Start-LocalServer '0.0.0.0' $port $sync
    $t = 0; while (-not $sync['ready'] -and $t -lt 50) { Start-Sleep -Milliseconds 100; $t++ }
    Send-FileTo $testFile $localIp $port
    $t = 0; while (-not $sync['done'] -and $t -lt 100) { Start-Sleep -Milliseconds 100; $t++ }
    $srv.PS.Dispose(); $srv.RS.Dispose()
    Log ("  Transferred {0:N0} bytes over LAN IP" -f $sync['received'])
}
$results['lan'] = $r.Deltas
foreach ($pname in $r.Deltas.Keys) {
    $d = $r.Deltas[$pname]
    $bl = if ($baseline.ContainsKey($pname)) { $baseline[$pname].TotalKb } else { 0 }
    $above = $d.TotalKb - $bl
    $flag = if ($above -gt 5) { '[ACTIVITY]' } else { '   quiet' }
    Log ("    {0,-30} total: {1,8:N1} KB  (+{2:N1} vs baseline) {3}" -f $d.Label, $d.TotalKb, $above, $flag)
}
Start-Sleep -Seconds 3

# -- TEST D: INTERNET ---------------------------------------------------------
LogHeader 'TEST D: INTERNET UPLOAD (httpbin.org)'
Log '  Same file sent to external server over HTTPS.'
Log ''
$r = Invoke-TransferTest {
    try {
        $resp = Invoke-WebRequest -Uri 'https://httpbin.org/post' -Method Post `
            -InFile $testFile -ContentType 'application/octet-stream' -TimeoutSec 15 -UseBasicParsing
        Log ("  Upload complete: HTTP {0}" -f $resp.StatusCode)
    } catch [System.Net.WebException] {
        if ($_.Exception.Message -match 'SSL|trust|certificate') {
            Log '  SSL intercepted by Zscaler (expected)'
        } else {
            Log ("  Upload error: {0}" -f $_.Exception.Message.Substring(0, [math]::Min(80, $_.Exception.Message.Length)))
        }
    } catch {
        Log ("  Upload error: {0}" -f $_.Exception.Message.Substring(0, [math]::Min(80, $_.Exception.Message.Length)))
    }
}
$results['internet'] = $r.Deltas
foreach ($pname in $r.Deltas.Keys) {
    $d = $r.Deltas[$pname]
    $bl = if ($baseline.ContainsKey($pname)) { $baseline[$pname].TotalKb } else { 0 }
    $above = $d.TotalKb - $bl
    $flag = if ($above -gt 5) { '[ACTIVITY]' } else { '   quiet' }
    Log ("    {0,-30} total: {1,8:N1} KB  (+{2:N1} vs baseline) {3}" -f $d.Label, $d.TotalKb, $above, $flag)
}

# =============================================================================
#  COMPARISON TABLE
# =============================================================================
LogHeader 'SIDE-BY-SIDE IO COMPARISON (KB)'
$allAgents = @{}
foreach ($tn in $results.Keys) { foreach ($k in $results[$tn].Keys) { $allAgents[$k] = $true } }
$sortedAgents = $allAgents.Keys | Sort-Object

Log ''
Log ("  {0,-30} {1,10} {2,10} {3,10} {4,10}" -f 'Agent', 'Baseline', 'Localhost', 'LAN IP', 'Internet')
Log ("  {0} {1} {2} {3} {4}" -f ('-'*30), ('-'*10), ('-'*10), ('-'*10), ('-'*10))
foreach ($pname in $sortedAgents) {
    $p = Get-Profile $pname
    $bl = if ($results['baseline'].ContainsKey($pname)) { $results['baseline'][$pname].TotalKb } else { 0 }
    $lo = if ($results['localhost'].ContainsKey($pname)) { $results['localhost'][$pname].TotalKb } else { 0 }
    $la = if ($results['lan'].ContainsKey($pname)) { $results['lan'][$pname].TotalKb } else { 0 }
    $inet = if ($results['internet'].ContainsKey($pname)) { $results['internet'][$pname].TotalKb } else { 0 }
    Log ("  {0,-30} {1,8:N1}KB {2,8:N1}KB {3,8:N1}KB {4,8:N1}KB" -f $p.Label, $bl, $lo, $la, $inet)
}

# =============================================================================
#  PER-AGENT DETAILED ANALYSIS
# =============================================================================
LogHeader 'DETAILED ANALYSIS PER AGENT'
foreach ($pname in $sortedAgents) {
    $p = Get-Profile $pname
    $blVal = if ($results['baseline'].ContainsKey($pname)) { $results['baseline'][$pname].TotalKb } else { 0 }
    $loVal = if ($results['localhost'].ContainsKey($pname)) { $results['localhost'][$pname].TotalKb } else { 0 }
    $laVal = if ($results['lan'].ContainsKey($pname)) { $results['lan'][$pname].TotalKb } else { 0 }
    $inetVal = if ($results['internet'].ContainsKey($pname)) { $results['internet'][$pname].TotalKb } else { 0 }
    $loAbove = $loVal - $blVal
    $laAbove = $laVal - $blVal
    $inetAbove = $inetVal - $blVal
    $reactsLocal = ($loAbove -gt 10) -or ($laAbove -gt 10)
    $reactsInternet = $inetAbove -gt 10

    Log ''
    Log '  +-----------------------------------------------------'
    Log ("  | {0} ({1})" -f $p.Label, $pname)
    Log ("  | Category: {0}" -f $p.Category)
    Log '  +-----------------------------------------------------'
    Log ("  | Purpose: {0}" -f $p.Purpose)
    Log '  |'
    Log '  | IO above baseline:'
    Log ("  |   Localhost: +{0:N1} KB {1}" -f $loAbove, $(if ($loAbove -gt 10) { '[!]' } else { 'normal' }))
    Log ("  |   LAN IP:    +{0:N1} KB {1}" -f $laAbove, $(if ($laAbove -gt 10) { '[!]' } else { 'normal' }))
    Log ("  |   Internet:  +{0:N1} KB {1}" -f $inetAbove, $(if ($inetAbove -gt 10) { '[!]' } else { 'normal' }))
    Log '  |'
    if ($reactsLocal -or $reactsInternet) {
        Log '  | What the IO spike ACTUALLY means:'
        Log ("  | {0}" -f $p.IoSpikeMeaning)
        Log '  |'
    }
    if ($reactsLocal) {
        Log '  | Local traffic assessment:'
        Log ("  | {0}" -f $p.LocalTrafficNote)
        Log '  |'
    }
    Log '  | What this agent CAN see:'
    foreach ($item in $p.CanSee) { Log ("  |   + {0}" -f $item) }
    Log '  |'
    Log '  | What this agent CANNOT see:'
    foreach ($item in $p.CannotSee) { Log ("  |   - {0}" -f $item) }
    Log '  |'

    switch -Regex ($pname) {
        '^(zsatunnel|zscaler)\.exe$' {
            if ($reactsLocal) {
                Log '  | [~] VERDICT: Zscaler''s agent detected local network activity.'
                Log '  |    It likely logged connection metadata (IP, port, bytes).'
                Log '  |    Content inspection of local traffic is UNLIKELY - Zscaler''s'
                Log '  |    cloud inspection typically only applies to internet-bound traffic.'
                Log '  |    Your file CONTENTS over local transfer are probably NOT read.'
            } else { Log '  | [OK] VERDICT: Did not react to local traffic.' }
            break
        }
        '^mssense\.exe$' {
            if ($reactsLocal) {
                Log '  | [~] VERDICT: Defender for Endpoint logged the network activity.'
                Log '  |    It recorded METADATA (which process, destination IP, bytes).'
                Log '  |    It did NOT read the file contents being transferred.'
                Log '  |    This is behavioral telemetry, not content inspection.'
            } else { Log '  | [OK] VERDICT: Did not react to local traffic.' }
            break
        }
        '^msmpeng\.exe$' {
            if ($reactsLocal) {
                Log '  | [OK] VERDICT: Defender Antimalware scanned the TEST FILE on disk.'
                Log '  |    This is a malware scan, NOT surveillance. It would do the same'
                Log '  |    if you simply opened the file. It does not monitor transfers.'
            } else { Log '  | [OK] VERDICT: No significant reaction.' }
            break
        }
        '^nxtcoordinator\.exe$' {
            if ($reactsLocal) {
                Log '  | [OK] VERDICT: Nexthink collected routine performance telemetry.'
                Log '  |    It logged app usage metrics (python.exe was active).'
                Log '  |    It has NO visibility into file contents or transfer data.'
                Log '  |    This is an IT experience tool, not a security tool.'
            } else { Log '  | [OK] VERDICT: Normal background activity.' }
            break
        }
        '^(ccmexec|defendpointservice|smartscreen)\.exe$' {
            Log ("  | [OK] VERDICT: {0} is not a surveillance tool." -f $p.Label)
            Log '  |    Any IO activity is routine operational tasks.'
            break
        }
        default {
            if ($reactsLocal) {
                Log '  | [~] VERDICT: Activity detected - investigate this process.'
            } else { Log '  | [OK] VERDICT: No significant reaction.' }
        }
    }
    Log '  +-----------------------------------------------------'
}

# =============================================================================
#  FINAL SUMMARY
# =============================================================================
LogHeader 'FINAL SUMMARY: LOCAL TRANSFER VISIBILITY'
Log @'

  Understanding the difference:

  +----------------------------------------------------------+
  |  "REACTED" != "SAW YOUR CONTENT"                         |
  |                                                          |
  |  An IO spike means the process did SOME work.            |
  |  What that work was depends on the tool:                 |
  |                                                          |
  |  - Zscaler:    logged connection metadata (IP, port)     |
  |                Content inspection = internet only        |
  |  - MsSense:    logged behavioral telemetry               |
  |                (process X connected to IP Y)             |
  |  - MsMpEng:    scanned the file for malware              |
  |                (not related to the transfer)             |
  |  - Nexthink:   collected app usage metrics               |
  |                (no content visibility at all)            |
  |  - SCCM:       routine inventory/compliance check        |
  |  - BeyondTrust: checked app privilege level              |
  +----------------------------------------------------------+
'@

Log '  For a LOCAL network transfer (e.g., LocalSend over WiFi):'
Log ''

$contentInspectors = @()
$metadataLoggers = @()
$notRelevant = @()
foreach ($pname in $sortedAgents) {
    $p = Get-Profile $pname
    $blVal = if ($results['baseline'].ContainsKey($pname)) { $results['baseline'][$pname].TotalKb } else { 0 }
    $laVal = if ($results['lan'].ContainsKey($pname)) { $results['lan'][$pname].TotalKb } else { 0 }
    $laAbove = $laVal - $blVal
    switch -Regex ($pname) {
        '^(zsatunnel|zscaler)\.exe$' { if ($laAbove -gt 10) { $metadataLoggers += $p.Label } else { $notRelevant += $p.Label }; break }
        '^mssense\.exe$'             { if ($laAbove -gt 10) { $metadataLoggers += $p.Label } else { $notRelevant += $p.Label }; break }
        '^msmpeng\.exe$'             { if ($laAbove -gt 10) { $notRelevant += ("{0} (malware scan only)" -f $p.Label) } else { $notRelevant += $p.Label }; break }
        '^nxtcoordinator\.exe$'      { $notRelevant += ("{0} (app metrics only)" -f $p.Label); break }
        default                      { if ($laAbove -gt 10) { $metadataLoggers += $p.Label } else { $notRelevant += $p.Label } }
    }
}

if ($contentInspectors.Count -gt 0) {
    Log '  [RED] CAN SEE FILE CONTENTS:'
    foreach ($name in $contentInspectors) { Log ("      - {0}" -f $name) }
}
if ($metadataLoggers.Count -gt 0) {
    Log '  [YELLOW] CAN SEE METADATA (connection happened, bytes transferred):'
    Log '     But CANNOT see what was inside the file:'
    foreach ($name in $metadataLoggers) { Log ("      - {0}" -f $name) }
}
if ($notRelevant.Count -gt 0) {
    Log '  [GREEN] NOT RELEVANT to file transfer monitoring:'
    foreach ($name in $notRelevant) { Log ("      - {0}" -f $name) }
}

if ($contentInspectors.Count -eq 0) {
    Log @'

  +----------------------------------------------------------+
  |  CONCLUSION: No agent can see the CONTENTS of a local    |
  |  network file transfer. Some agents log that a transfer  |
  |  happened (metadata), but the file data itself is NOT    |
  |  inspected for local-to-local traffic.                   |
  |                                                          |
  |  Content inspection (reading what's inside your files)   |
  |  only happens for INTERNET-bound traffic via Zscaler's   |
  |  SSL interception.                                       |
  +----------------------------------------------------------+
'@
} else {
    Log @'

  +----------------------------------------------------------+
  |  [!] Content inspection detected on local transfers.     |
  |  Use the sanitizer before ANY transfer method.           |
  +----------------------------------------------------------+
'@
}

LogSeparator
Log ''
Log ("  Test completed: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))

# Cleanup
Remove-Item -Path $testDir -Recurse -Force -ErrorAction SilentlyContinue

# Save log
[System.IO.File]::WriteAllText($logPath, ($script:LogLines -join "`r`n"))
Log ''
Log ("  Full log saved to: {0}" -f $logPath)

Read-Host "`n  Press Enter to close"
