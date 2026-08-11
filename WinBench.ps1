#requires -Version 5.1
[CmdletBinding()]
param(
 [ValidateSet('Menu','Quick','Full','Telemetry')][string]$Mode='Menu',
 [string]$OutputRoot="$env:USERPROFILE\Desktop\WinBench-Reports",
 [ValidateRange(5,300)][int]$TelemetrySeconds=30,
 [ValidateRange(1,10)][int]$SampleInterval=1,
 [switch]$KeepTemp
)

$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'

function hdr($s){write-host "`n$('='*72)" -f darkgray;write-host "  $s" -f cyan;write-host $('='*72) -f darkgray}
function say($s){write-host "[*] $s" -f gray}
function good($s){write-host "[+] $s" -f green}
function warnx($s){write-host "[!] $s" -f yellow}

function val($o,$p,$d=$null){
 if($null -eq $o){return $d}
 $x=$o.PSObject.Properties[$p]
 if($null -eq $x){return $d}
 $x.Value
}
function gib($n){if($n -le 0){0}else{[math]::Round($n/1GB,2)}}

function nvsmi {
 try{$c=Get-Command nvidia-smi.exe -ea stop|select -first 1;if($c.Source){return $c.Source}}catch{}
 foreach($x in @("$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe","$env:SystemRoot\System32\nvidia-smi.exe")){
  if($x -and (test-path -literalpath $x)){return $x}
 }
 $null
}
function pwr {
 try{
  $x=(& powercfg.exe /GETACTIVESCHEME 2>$null|out-string).Trim()
  if($x -match '\((.+)\)'){return $Matches[1]}
  if($x){return $x}
 }catch{}
 'Unknown'
}

function sysinfo {
 say "Collecting system information..."
 $os=Get-CimInstance Win32_OperatingSystem -ea silentlycontinue
 $cs=Get-CimInstance Win32_ComputerSystem -ea silentlycontinue
 $cpu=Get-CimInstance Win32_Processor -ea silentlycontinue|select -first 1
 $gpu=@(Get-CimInstance Win32_VideoController -ea silentlycontinue)
 $bios=Get-CimInstance Win32_BIOS -ea silentlycontinue
 $bb=Get-CimInstance Win32_BaseBoard -ea silentlycontinue

 try{
  $pd=@(Get-PhysicalDisk -ea stop|%{
   [pscustomobject]@{FriendlyName=$_.FriendlyName;MediaType=$_.MediaType;BusType=$_.BusType;SizeGiB=gib $_.Size;HealthStatus=$_.HealthStatus}
  })
 }catch{
  $pd=@(Get-CimInstance Win32_DiskDrive -ea silentlycontinue|%{
   [pscustomobject]@{FriendlyName=$_.Model;MediaType=$_.MediaType;BusType=$_.InterfaceType;SizeGiB=gib $_.Size;HealthStatus='Unknown'}
  })
 }

 $gn=if($gpu.Count){($gpu|%{$_.Name})-join '; '}else{'Unknown'}
 [pscustomobject]@{
  Timestamp=(get-date).ToString('o');ComputerName=$env:COMPUTERNAME
  Manufacturer=val $cs Manufacturer Unknown;Model=val $cs Model Unknown
  OS=val $os Caption Unknown;OSVersion=val $os Version Unknown;BuildNumber=val $os BuildNumber Unknown
  CPU=val $cpu Name Unknown;PhysicalCores=val $cpu NumberOfCores 0;LogicalProcessors=val $cpu NumberOfLogicalProcessors $env:NUMBER_OF_PROCESSORS
  MemoryGiB=gib (val $cs TotalPhysicalMemory 0);GPU=$gn
  BIOSVersion=if($bios){$bios.SMBIOSBIOSVersion-join ', '}else{'Unknown'}
  BaseBoard=if($bb){"$($bb.Manufacturer) $($bb.Product)".Trim()}else{'Unknown'}
  PowerPlan=pwr;PowerShell=$PSVersionTable.PSVersion.ToString()
  IsElevated=([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  PhysicalDisks=$pd
 }
}

function counterok($x){try{$null=Get-Counter $x -MaxSamples 1 -ea stop;$true}catch{$false}}
function counters {
 $a=@(
  '\Processor(_Total)\% Processor Time',
  '\Processor Information(_Total)\% Processor Performance',
  '\Processor Information(_Total)\Processor Frequency',
  '\Memory\Available MBytes',
  '\Memory\% Committed Bytes In Use',
  '\PhysicalDisk(_Total)\Disk Bytes/sec',
  '\PhysicalDisk(_Total)\Avg. Disk Queue Length'
 )
 @($a|?{counterok $_})
}
function ckey($x){
 $p=$x.ToLowerInvariant()
 if($p-like '*\% processor time'){return 'CPUPercent'}
 if($p-like '*\% processor performance'){return 'CPUPerformancePercent'}
 if($p-like '*\processor frequency'){return 'CPUFrequencyMHz'}
 if($p-like '*\available mbytes'){return 'MemoryAvailableMB'}
 if($p-like '*\% committed bytes in use'){return 'MemoryCommittedPercent'}
 if($p-like '*\disk bytes/sec'){return 'DiskBytesPerSec'}
 if($p-like '*\avg. disk queue length'){return 'DiskQueueLength'}
 ($x-replace '[^a-zA-Z0-9]+','_').Trim('_')
}

function nvrow($nv){
 if(!$nv){return}
 try{
  $q='--query-gpu=index,name,utilization.gpu,temperature.gpu,power.draw,clocks.current.graphics,clocks.current.memory,memory.used,memory.total'
  $lines=@(& $nv $q '--format=csv,noheader,nounits' 2>$null)
  foreach($line in $lines){
   $p=@($line-split '\s*,\s*');if($p.Count-lt 9){continue}
   [pscustomobject]@{
    Index=[int]$p[0];Name=$p[1];UtilizationPct=[double]$p[2];TemperatureC=[double]$p[3];PowerW=[double]$p[4]
    GraphicsClockMHz=[double]$p[5];MemoryClockMHz=[double]$p[6];MemoryUsedMB=[double]$p[7];MemoryTotalMB=[double]$p[8]
   }
  }
 }catch{}
}

function telemetry($sec,$every,$nv){
 $cl=@(counters)
 $rows=New-Object 'System.Collections.Generic.List[object]'
 $sw=[Diagnostics.Stopwatch]::StartNew()
 while($sw.Elapsed.TotalSeconds-lt $sec){
  $r=[ordered]@{Timestamp=(get-date).ToString('o');ElapsedSec=[math]::Round($sw.Elapsed.TotalSeconds,2)}
  if($cl.Count){
   try{
    $z=Get-Counter -Counter $cl -MaxSamples 1 -ea stop
    foreach($i in $z.CounterSamples){$r[(ckey $i.Path)]=[math]::Round([double]$i.CookedValue,2)}
   }catch{$r.CounterError=$_.Exception.Message}
  }
  $n=@(nvrow $nv)
  if($n.Count){
   $g=$n[0]
   $r.GPUUtilizationPct=$g.UtilizationPct;$r.GPUTemperatureC=$g.TemperatureC;$r.GPUPowerW=$g.PowerW
   $r.GPUGraphicsClockMHz=$g.GraphicsClockMHz;$r.GPUMemoryClockMHz=$g.MemoryClockMHz;$r.GPUMemoryUsedMB=$g.MemoryUsedMB
   $r.NvidiaAllGPUs=$n|ConvertTo-Json -Compress -Depth 4
  }
  $rows.Add([pscustomobject]$r);sleep $every
 }
 $sw.Stop();$rows.ToArray()
}

function stat($rows,$prop){
 $v=@($rows|%{$_.PSObject.Properties[$prop]}|?{$null-ne $_ -and $null-ne $_.Value}|%{[double]$_.Value})
 if(!$v.Count){return}
 $m=$v|measure -Average -Minimum -Maximum
 [pscustomobject]@{Metric=$prop;Average=[math]::Round($m.Average,2);Minimum=[math]::Round($m.Minimum,2);Maximum=[math]::Round($m.Maximum,2);Samples=$v.Count}
}
function sumtele($rows){
 $out=@()
 foreach($p in @('CPUPercent','CPUPerformancePercent','CPUFrequencyMHz','MemoryAvailableMB','MemoryCommittedPercent','DiskBytesPerSec','DiskQueueLength','GPUUtilizationPct','GPUTemperatureC','GPUPowerW','GPUGraphicsClockMHz','GPUMemoryClockMHz','GPUMemoryUsedMB')){
  $s=stat $rows $p;if($s){$out+=$s}
 }
 $out
}

function cpubench($sec=10,$workers=0){
 if($workers-le 0){$workers=[Environment]::ProcessorCount}
 $workers=[math]::Max(1,[math]::Min($workers,64))
 say "CPU test: SHA-256 workload, $workers worker(s), $sec sec."

 $work={
  param($secs,$seed,$ticks)
  $ErrorActionPreference='Stop';$bs=4MB;$buf=New-Object byte[] $bs
  $rnd=New-Object System.Random $seed;$rnd.NextBytes($buf)
  $sha=[Security.Cryptography.SHA256]::Create();[long]$n=0
  $go=[datetime]::new([long]$ticks,[DateTimeKind]::Utc)
  while([datetime]::UtcNow-lt $go){sleep -Milliseconds 20}
  $sw=[Diagnostics.Stopwatch]::StartNew()
  try{while($sw.Elapsed.TotalSeconds-lt $secs){$null=$sha.ComputeHash($buf);$n++}}
  finally{$sha.Dispose();$sw.Stop()}
  [pscustomobject]@{Iterations=$n;Seconds=$sw.Elapsed.TotalSeconds;BytesPerIteration=$bs}
 }

 $all=[Diagnostics.Stopwatch]::StartNew();$jobs=@();$go=[datetime]::UtcNow.AddSeconds(5)
 try{
  0..($workers-1)|%{$jobs+=Start-Job $work -ArgumentList $sec,(1000+$_),$go.Ticks}
  $null=Wait-Job $jobs;$res=@($jobs|Receive-Job)
 }finally{
  if($jobs){$jobs|Remove-Job -Force -ea silentlycontinue}
  $all.Stop()
 }
 [long]$n=($res|measure Iterations -Sum).Sum
 [double]$b=$n*4MB;[double]$t=[math]::Max(.001,$sec)
 [pscustomobject]@{
  Test='CPU SHA-256 Multi-worker';Workers=$workers;RequestedSeconds=$sec;WallSeconds=[math]::Round($all.Elapsed.TotalSeconds,3)
  Iterations=$n;HashesPerSecond=[math]::Round($n/$t,2);ThroughputMiBPerSec=[math]::Round(($b/1MB)/$t,2)
 }
}

function membench($mb=64,$passes=12){
 say "Memory test: Buffer.BlockCopy, $mb MiB x $passes pass(es)."
 try{
  $bytes=$mb*1MB;$a=New-Object byte[] $bytes;$b=New-Object byte[] $bytes
  $r=New-Object System.Random 7331;$r.NextBytes($a)
  [Buffer]::BlockCopy($a,0,$b,0,$bytes)
  $sw=[Diagnostics.Stopwatch]::StartNew();for($i=0;$i-lt$passes;$i++){[Buffer]::BlockCopy($a,0,$b,0,$bytes)};$sw.Stop()
  $total=[double]$mb*$passes
  [pscustomobject]@{Test='Memory BlockCopy';BufferMiB=$mb;Passes=$passes;TotalCopiedMiB=[math]::Round($total,2);Seconds=[math]::Round($sw.Elapsed.TotalSeconds,4);ThroughputMiBps=[math]::Round($total/[math]::Max(.0001,$sw.Elapsed.TotalSeconds),2)}
 }catch{[pscustomobject]@{Test='Memory BlockCopy';Error=$_.Exception.Message}}
}

function diskbench($dir,$size=512,$bufmb=4){
 say "Disk test: sequential write/read, $size MiB temporary file."
 if(!(test-path -literalpath $dir)){ni -ItemType Directory -Path $dir -Force|out-null}
 $file=join-path $dir ("winbench-disk-"+[guid]::NewGuid().ToString('N')+".bin")
 $bs=$bufmb*1MB;$buf=New-Object byte[] $bs;(New-Object System.Random 2026).NextBytes($buf)
 $loops=[math]::Ceiling(($size*1MB)/$bs);[long]$target=[long]$size*1MB
 [long]$written=0;[long]$readn=0;$ws=$null;$rs=$null

 try{
  $opts=[IO.FileOptions]::SequentialScan-bor[IO.FileOptions]::WriteThrough
  $f=[IO.FileStream]::new($file,[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::None,$bs,$opts)
  try{
   $sw=[Diagnostics.Stopwatch]::StartNew()
   for($i=0;$i-lt$loops;$i++){
    $left=$target-$written;if($left-le 0){break}
    $n=[int][math]::Min([long]$bs,$left);$f.Write($buf,0,$n);$written+=$n
   }
   $f.Flush($true);$sw.Stop();$ws=$sw.Elapsed.TotalSeconds
  }finally{$f.Dispose()}

  $rb=New-Object byte[] $bs
  $f=[IO.FileStream]::new($file,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read,$bs,[IO.FileOptions]::SequentialScan)
  try{
   $sw=[Diagnostics.Stopwatch]::StartNew()
   while(($n=$f.Read($rb,0,$rb.Length))-gt 0){$readn+=$n}
   $sw.Stop();$rs=$sw.Elapsed.TotalSeconds
  }finally{$f.Dispose()}

  [pscustomobject]@{
   Test='Disk Sequential';Path=$dir;FileSizeMiB=[math]::Round($written/1MB,2)
   WriteSeconds=[math]::Round($ws,4);WriteMiBps=[math]::Round(($written/1MB)/[math]::Max(.0001,$ws),2)
   ReadSeconds=[math]::Round($rs,4);ReadMiBps=[math]::Round(($readn/1MB)/[math]::Max(.0001,$rs),2)
   WriteThroughUsed=$true;Note='Sequential local test; read speed can be affected by OS/storage caching.'
  }
 }catch{[pscustomobject]@{Test='Disk Sequential';Path=$dir;Error=$_.Exception.Message}}
 finally{if(!$KeepTemp-and(test-path -literalpath $file)){rm -literalpath $file -Force -ea silentlycontinue}}
}

function cpuwithtele($job,$seconds,$nv,$every=1){
 $defs=@{}
 foreach($n in 'telemetry','counters','counterok','ckey','nvrow'){
  $defs[$n]=(Get-Item "function:$n").ScriptBlock.ToString()
 }
 $tj=Start-Job -ArgumentList $seconds,$every,$nv,$defs -ScriptBlock{
  param($sec,$every,$nv,$d)
  foreach($k in $d.Keys){Set-Item "function:$k" ([scriptblock]::Create($d[$k]))}
  telemetry $sec $every $nv
 }
 try{
  $r=&$job;$t=@($tj|wait-job|receive-job)
  [pscustomobject]@{Result=$r;Telemetry=$t}
 }finally{Remove-Job $tj -Force -ea silentlycontinue}
}

function html($path,$sys,$bench,$ts,$mode){
 $css=@'
<style>
body{font-family:"Segoe UI",Arial,sans-serif;margin:28px;background:#111;color:#e8e8e8}h1,h2{font-weight:600}h1{margin-bottom:4px}.muted{color:#aaa;margin-top:0}.card{background:#1b1b1b;border:1px solid #333;border-radius:8px;padding:16px;margin:16px 0}table{width:100%;border-collapse:collapse;margin-top:8px}th,td{text-align:left;border-bottom:1px solid #333;padding:8px;font-size:13px}th{color:#ddd;background:#222}code{color:#ddd}.small{font-size:12px;color:#aaa}
</style>
'@
 $st=[pscustomobject]@{
  Computer=$sys.ComputerName;Manufacturer=$sys.Manufacturer;Model=$sys.Model
  OS="$($sys.OS) $($sys.OSVersion) (Build $($sys.BuildNumber))";CPU=$sys.CPU
  Cores="$($sys.PhysicalCores) physical / $($sys.LogicalProcessors) logical";Memory="$($sys.MemoryGiB) GiB"
  GPU=$sys.GPU;PowerPlan=$sys.PowerPlan;PowerShell=$sys.PowerShell;Elevated=$sys.IsElevated;BIOS=$sys.BIOSVersion;BaseBoard=$sys.BaseBoard
 }
 $a=$st|ConvertTo-Html -Fragment
 $b=if($bench.Count){$bench|ConvertTo-Html -Fragment}else{'<p>No benchmark tests were run.</p>'}
 $t=if($ts.Count){$ts|ConvertTo-Html -Fragment}else{'<p>No telemetry summary available.</p>'}
 $d=if($sys.PhysicalDisks.Count){$sys.PhysicalDisks|ConvertTo-Html -Fragment}else{'<p>No physical disk inventory available.</p>'}
 $body=@"
<!doctype html><html><head><meta charset="utf-8"><title>WinBench Report - $($sys.ComputerName)</title>$css</head>
<body><h1>WinBench Report</h1><p class="muted">$($sys.ComputerName) • Mode: $mode • $(get-date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<div class="card"><h2>System</h2>$a</div>
<div class="card"><h2>Benchmarks</h2>$b<p class="small">These are local repeatable measurements, not Cinebench-compatible scores and not intended to reproduce Maxon's workload.</p></div>
<div class="card"><h2>Telemetry Summary</h2>$t</div>
<div class="card"><h2>Physical Disks</h2>$d</div>
<div class="card"><h2>Interpretation Notes</h2><ul>
<li>Compare results from the same WinBench mode and similar background-load conditions.</li>
<li>Disk read results can be influenced by Windows and storage caching.</li>
<li>CPU temperature is omitted because Windows does not expose one universally reliable native CPU-temperature interface across systems.</li>
<li>NVIDIA telemetry appears only when <code>nvidia-smi</code> is already present.</li>
</ul></div></body></html>
"@
 [IO.File]::WriteAllText($path,$body,[Text.UTF8Encoding]::new($false))
}

function saveit($dir,$sys,$bench,$tele,$summary,$mode){
 if(!(test-path -literalpath $dir)){ni -ItemType Directory -Path $dir -Force|out-null}
 $jp=join-path $dir report.json;$hp=join-path $dir report.html;$cp=join-path $dir telemetry.csv
 [ordered]@{GeneratedAt=(get-date).ToString('o');Mode=$mode;System=$sys;Benchmarks=$bench;TelemetrySummary=$summary}|ConvertTo-Json -Depth 8|Set-Content -literalpath $jp -Encoding UTF8
 if($tele.Count){$tele|select * -ExcludeProperty NvidiaAllGPUs|Export-Csv -literalpath $cp -NoTypeInformation -Encoding UTF8}
 html $hp $sys $bench $summary $mode
 [pscustomobject]@{Directory=$dir;HTML=$hp;JSON=$jp;CSV=if(test-path -literalpath $cp){$cp}else{$null}}
}

function menu {
 write-host "`nWinBench" -f cyan
 write-host "Local Windows benchmark + diagnostics`n"
 write-host "  1) Quick benchmark"
 write-host "  2) Full benchmark"
 write-host "  3) Telemetry only"
 write-host "  4) Exit`n"
 while($true){
  switch(read-host Select){
   1{return 'Quick'} 2{return 'Full'} 3{return 'Telemetry'} 4{return 'Exit'}
   default{warnx 'Enter 1, 2, 3, or 4.'}
  }
 }
}

function runbench($mode){
 $stamp=get-date -Format yyyyMMdd-HHmmss
 $dir=join-path $OutputRoot "$env:COMPUTERNAME-$stamp";$tmp=join-path $dir temp
 ni -ItemType Directory -Path $tmp -Force|out-null

 hdr 'System Snapshot'
 $sys=sysinfo
 good $sys.CPU;say "$($sys.PhysicalCores) physical / $($sys.LogicalProcessors) logical CPU cores"
 say "$($sys.MemoryGiB) GiB RAM";say "GPU: $($sys.GPU)";say "Power plan: $($sys.PowerPlan)"
 $nv=nvsmi
 if($nv){good 'NVIDIA telemetry available.'}else{say 'nvidia-smi not found; NVIDIA-specific telemetry will be skipped.'}

 $bench=@();$tele=@()
 if($mode-eq'Telemetry'){
  hdr Telemetry;$tele=@(telemetry $TelemetrySeconds $SampleInterval $nv)
 }else{
  if($mode-eq'Quick'){$cpuSecs=8;$memPasses=10;$diskMB=256;$cpuTele=10}
  else{$cpuSecs=20;$memPasses=24;$diskMB=1024;$cpuTele=22}

  hdr 'CPU Benchmark'
  $pair=cpuwithtele {cpubench $cpuSecs $sys.LogicalProcessors} $cpuTele $nv $SampleInterval
  $c=$pair.Result;$tele+=@($pair.Telemetry);$bench+=$c
  good ("CPU: {0:N2} hashes/sec | {1:N2} MiB/sec"-f$c.HashesPerSecond,$c.ThroughputMiBPerSec)

  hdr 'Memory Benchmark'
  $m=membench 64 $memPasses;$bench+=$m
  if($m.PSObject.Properties['ThroughputMiBps']){good ("Memory copy: {0:N2} MiB/sec"-f$m.ThroughputMiBps)}else{warnx "Memory test failed: $($m.Error)"}

  hdr 'Disk Benchmark'
  $d=diskbench $tmp $diskMB 4;$bench+=$d
  if($d.PSObject.Properties['WriteMiBps']){good ("Disk write: {0:N2} MiB/sec | read: {1:N2} MiB/sec"-f$d.WriteMiBps,$d.ReadMiBps)}else{warnx "Disk test failed: $($d.Error)"}

  hdr 'Idle / Recovery Telemetry'
  $recover=if($mode-eq'Quick'){5}else{10};$tele+=@(telemetry $recover $SampleInterval $nv)
 }

 $summary=@(sumtele $tele)
 if(!$KeepTemp-and(test-path -literalpath $tmp)){rm $tmp -Recurse -Force -ea silentlycontinue}

 hdr Report
 $o=saveit $dir $sys $bench $tele $summary $mode
 good "HTML: $($o.HTML)";good "JSON: $($o.JSON)";if($o.CSV){good "CSV:  $($o.CSV)"}
 try{start $o.HTML}catch{warnx 'Could not automatically open the HTML report.'}
 $o
}

try{
 if($Mode-eq'Menu'){$Mode=menu;if($Mode-eq'Exit'){return}}
 hdr WinBench;say "Mode: $Mode";say "Reports: $OutputRoot";say 'No network connection or module installation is required.'
 $r=runbench $Mode
 write-host '';good Done;write-host "Report directory: $($r.Directory)"
}catch{
 write-host "`n[X] WinBench stopped: $($_.Exception.Message)" -f red
 write-host $_.ScriptStackTrace -f darkgray
 exit 1
}

# SIG # Begin signature block
# MIIb8gYJKoZIhvcNAQcCoIIb4zCCG98CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQURQggUS57p0pjvXEbqM4+Zw7M
# mjOgghZYMIIDGjCCAgKgAwIBAgIQSHk2u471qKhEmswFQlqdXzANBgkqhkiG9w0B
# AQsFADAlMSMwIQYDVQQDDBpNeSBQb3dlclNoZWxsIENvZGUgU2lnbmluZzAeFw0y
# NjA4MTEwMjAwMTlaFw0yNzA4MTEwMjIwMTlaMCUxIzAhBgNVBAMMGk15IFBvd2Vy
# U2hlbGwgQ29kZSBTaWduaW5nMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAyRFWb2ks5GKrrdHBwAb3I0YUJmEo7JWGJC4R2ivfSIs6DUtMS98M+HpKnOWp
# MpMRL3o3AOi3hFOGl5OzMnOSCKwe11fnFAHBZxtAA/ec4Cl+Xx8+93W+yGlgcG04
# bWztAygCOtd0LShsjZjFlD3iXLY7qFOkEZ/sKleLxC/IcoFiiYaDP6Mt4hjmMIEU
# GsGw323FERIBKwa8jM1LoIKsER00pq8QFWHqU2KmgGFCp0YZtyhDz7G73hgFI9N9
# 0kSGZvU0dJHfMOtBcIMG16O4slIeEPtgd5iotQ3tdD8TM9Q+8dhLIRLscT/2eXFf
# q0BQNWZVMnopCi4DU5QoFokAQQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFFmA95+8/hJgCm8KIIM5ab5OVBKv
# MA0GCSqGSIb3DQEBCwUAA4IBAQCmms7kh52g6Q+MKzpDfy73TIGrbETikrmQ1wBa
# IMbYzJhLJO06VdkZHY1snmvRINjWgFjC74XxnExMo6T67n9mNn9xKZ8Lgp4nx4vy
# muUEivXy6SHwVBZ+ZOjdXayPimLv530xxoncPFFa0ArcQduNa9hKaALOEemopI/z
# A7vhvn4E3fsyCxsbwbk2AT5HChwk1sjqKobEH64J1sRr1VFXSzZPw6SxVt05CGpU
# SuBFrSRQB89t4tGcXvIwK+F2gnpMGhIPiHzTY+V41aLlZJEhWiNAj4gZY9K1XRpT
# zN4mIvngkwx5lgyOixGe+5oUtrT/bq7DRgLIARjpVZHuNj9bMIIFjTCCBHWgAwIB
# AgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIw
# ODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3yithZwuEppz1Y
# q3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lX
# FllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDVySAdYyktzuxe
# TsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiODCu3T6cw2Vbu
# yntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I
# 9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmg
# Z92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse
# 5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADMfRyVw4/3IbKy
# Ebe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwh
# HbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXKchYiCd98THU/
# Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t9dmpsh3lGwID
# AQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM
# 3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDgYD
# VR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+
# MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUA
# A4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSI
# d229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n096wwepqLsl7U
# z9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxA
# GTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9vytsgjTVgHAID
# yyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW
# /VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIGtDCCBJygAwIBAgIQDcesVwX/IZkuQEMi
# DDpJhjANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGln
# aUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhE
# aWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjUwNTA3MDAwMDAwWhcNMzgwMTE0
# MjM1OTU5WjBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4x
# QTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQw
# OTYgU0hBMjU2IDIwMjUgQ0ExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtHgx0wqYQXK+PEbAHKx126NGaHS0URedTa2NDZS1mZaDLFTtQ2oRjzUXMmxC
# qvkbsDpz4aH+qbxeLho8I6jY3xL1IusLopuW2qftJYJaDNs1+JH7Z+QdSKWM06qc
# hUP+AbdJgMQB3h2DZ0Mal5kYp77jYMVQXSZH++0trj6Ao+xh/AS7sQRuQL37QXbD
# hAktVJMQbzIBHYJBYgzWIjk8eDrYhXDEpKk7RdoX0M980EpLtlrNyHw0Xm+nt5pn
# YJU3Gmq6bNMI1I7Gb5IBZK4ivbVCiZv7PNBYqHEpNVWC2ZQ8BbfnFRQVESYOszFI
# 2Wv82wnJRfN20VRS3hpLgIR4hjzL0hpoYGk81coWJ+KdPvMvaB0WkE/2qHxJ0ucS
# 638ZxqU14lDnki7CcoKCz6eum5A19WZQHkqUJfdkDjHkccpL6uoG8pbF0LJAQQZx
# st7VvwDDjAmSFTUms+wV/FbWBqi7fTJnjq3hj0XbQcd8hjj/q8d6ylgxCZSKi17y
# Vp2NL+cnT6Toy+rN+nM8M7LnLqCrO2JP3oW//1sfuZDKiDEb1AQ8es9Xr/u6bDTn
# YCTKIsDq1BtmXUqEG1NqzJKS4kOmxkYp2WyODi7vQTCBZtVFJfVZ3j7OgWmnhFr4
# yUozZtqgPrHRVHhGNKlYzyjlroPxul+bgIspzOwbtmsgY1MCAwEAAaOCAV0wggFZ
# MBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFO9vU0rp5AZ8esrikFb2L9RJ
# 7MtOMB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB/wQE
# AwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5j
# cnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJ
# YIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQAXzvsWgBz+Bz0RdnEwvb4LyLU0
# pn/N0IfFiBowf0/Dm1wGc/Do7oVMY2mhXZXjDNJQa8j00DNqhCT3t+s8G0iP5kvN
# 2n7Jd2E4/iEIUBO41P5F448rSYJ59Ib61eoalhnd6ywFLerycvZTAz40y8S4F3/a
# +Z1jEMK/DMm/axFSgoR8n6c3nuZB9BfBwAQYK9FHaoq2e26MHvVY9gCDA/JYsq7p
# GdogP8HRtrYfctSLANEBfHU16r3J05qX3kId+ZOczgj5kjatVB+NdADVZKON/gnZ
# ruMvNYY2o1f4MXRJDMdTSlOLh0HCn2cQLwQCqjFbqrXuvTPSegOOzr4EWj7PtspI
# HBldNE2K9i697cvaiIo2p61Ed2p8xMJb82Yosn0z4y25xUbI7GIN/TpVfHIqQ6Ku
# /qjTY6hc3hsXMrS+U0yy+GWqAXam4ToWd2UQ1KYT70kZjE4YtL8Pbzg0c1ugMZyZ
# Zd/BdHLiRu7hAWE6bTEm4XYRkA6Tl4KSFLFk43esaUeqGkH/wyW4N7OigizwJWeu
# kcyIPbAvjSabnf7+Pu0VrFgoiovRDiyx3zEdmcif/sYQsfch28bZeUz2rtY/9TCA
# 6TD8dC3JE3rYkrhLULy7Dc90G6e8BlqmyIjlgp2+VqsS9/wQD7yFylIz0scmbKvF
# oW2jNrbM1pD2T7m3XDCCBu0wggTVoAMCAQICEAqA7xhLjfEFgtHEdqeVdGgwDQYJ
# KoZIhvcNAQELBQAwaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBS
# U0E0MDk2IFNIQTI1NiAyMDI1IENBMTAeFw0yNTA2MDQwMDAwMDBaFw0zNjA5MDMy
# MzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7
# MDkGA1UEAxMyRGlnaUNlcnQgU0hBMjU2IFJTQTQwOTYgVGltZXN0YW1wIFJlc3Bv
# bmRlciAyMDI1IDEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDQRqwt
# Esae0OquYFazK1e6b1H/hnAKAd/KN8wZQjBjMqiZ3xTWcfsLwOvRxUwXcGx8AUjn
# i6bz52fGTfr6PHRNv6T7zsf1Y/E3IU8kgNkeECqVQ+3bzWYesFtkepErvUSbf+EI
# YLkrLKd6qJnuzK8Vcn0DvbDMemQFoxQ2Dsw4vEjoT1FpS54dNApZfKY61HAldytx
# NM89PZXUP/5wWWURK+IfxiOg8W9lKMqzdIo7VA1R0V3Zp3DjjANwqAf4lEkTlCDQ
# 0/fKJLKLkzGBTpx6EYevvOi7XOc4zyh1uSqgr6UnbksIcFJqLbkIXIPbcNmA98Os
# kkkrvt6lPAw/p4oDSRZreiwB7x9ykrjS6GS3NR39iTTFS+ENTqW8m6THuOmHHjQN
# C3zbJ6nJ6SXiLSvw4Smz8U07hqF+8CTXaETkVWz0dVVZw7knh1WZXOLHgDvundrA
# tuvz0D3T+dYaNcwafsVCGZKUhQPL1naFKBy1p6llN3QgshRta6Eq4B40h5avMcpi
# 54wm0i2ePZD5pPIssoszQyF4//3DoK2O65Uck5Wggn8O2klETsJ7u8xEehGifgJY
# i+6I03UuT1j7FnrqVrOzaQoVJOeeStPeldYRNMmSF3voIgMFtNGh86w3ISHNm0Ia
# adCKCkUe2LnwJKa8TIlwCUNVwppwn4D3/Pt5pwIDAQABo4IBlTCCAZEwDAYDVR0T
# AQH/BAIwADAdBgNVHQ4EFgQU5Dv88jHt/f3X85FxYxlQQ89hjOgwHwYDVR0jBBgw
# FoAU729TSunkBnx6yuKQVvYv1Ensy04wDgYDVR0PAQH/BAQDAgeAMBYGA1UdJQEB
# /wQMMAoGCCsGAQUFBwMIMIGVBggrBgEFBQcBAQSBiDCBhTAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMF0GCCsGAQUFBzAChlFodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdS
# U0E0MDk2U0hBMjU2MjAyNUNBMS5jcnQwXwYDVR0fBFgwVjBUoFKgUIZOaHR0cDov
# L2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5n
# UlNBNDA5NlNIQTI1NjIwMjVDQTEuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsG
# CWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAZSqt8RwnBLmuYEHs0QhEnmNA
# ciH45PYiT9s1i6UKtW+FERp8FgXRGQ/YAavXzWjZhY+hIfP2JkQ38U+wtJPBVBaj
# YfrbIYG+Dui4I4PCvHpQuPqFgqp1PzC/ZRX4pvP/ciZmUnthfAEP1HShTrY+2DE5
# qjzvZs7JIIgt0GCFD9ktx0LxxtRQ7vllKluHWiKk6FxRPyUPxAAYH2Vy1lNM4kze
# kd8oEARzFAWgeW3az2xejEWLNN4eKGxDJ8WDl/FQUSntbjZ80FU3i54tpx5F/0Kr
# 15zW/mJAxZMVBrTE2oi0fcI8VMbtoRAmaaslNXdCG1+lqvP4FbrQ6IwSBXkZagHL
# hFU9HCrG/syTRLLhAezu/3Lr00GrJzPQFnCEH1Y58678IgmfORBPC1JKkYaEt2Od
# Dh4GmO0/5cHelAK2/gTlQJINqDr6JfwyYHXSd+V08X1JUPvB4ILfJdmL+66Gp3CS
# BXG6IwXMZUXBhtCyIaehr0XkBoDIGMUG1dUtwq1qmcwbdUfcSYCn+OwncVUXf53V
# JUNOaMWMts0VlRYxe5nK+At+DI96HAlXHAL5SlfYxJ7La54i71McVWRP66bW+yER
# NpbJCjyCYG2j+bdpxo/1Cy4uPcU3AWVPGrbn5PhDBf3Froguzzhk++ami+r3Qrx5
# bIbY3TVzgiFI7Gq3zWcxggUEMIIFAAIBATA5MCUxIzAhBgNVBAMMGk15IFBvd2Vy
# U2hlbGwgQ29kZSBTaWduaW5nAhBIeTa7jvWoqESazAVCWp1fMAkGBSsOAwIaBQCg
# eDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJ
# BDEWBBRlRYQ6Hea5WwpvIjjPTrVW7QfGUTANBgkqhkiG9w0BAQEFAASCAQB7NJFE
# mz+gd+b6QxmrGfnPs9hYv4x2NtecU9yOiVXexNCRfttSB/Jf/l7+hmlOOHzz+6Nz
# X176X24F61sLNKiRmw6G4fVuWIzVHQgbcDJhfpLpZX1TQOvp4yeCtarPM90eE8N0
# //Lofc67g5aWOniYJB3YcqFk4AaqDMfs1EFmixElk6PFTE65LH/dOfrTn+kFkDa4
# LddDzmp9hajB4UokvnOh6aX8b/uyFw5BuurzZp88GOVCyv4+MKDoAtVv9qMQUiIV
# CQLgRY8ufHBNuOG6Wa7d1NKKXhMYMoJqE34ECCk8YvUwaLr1ho7+M2Xdgy6dG8cO
# Fkzf4Ool6rpOnmmfoYIDJjCCAyIGCSqGSIb3DQEJBjGCAxMwggMPAgEBMH0waTEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhE
# aWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAy
# MDI1IENBMQIQCoDvGEuN8QWC0cR2p5V0aDANBglghkgBZQMEAgEFAKBpMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI2MDgxMTAyMTM1
# MFowLwYJKoZIhvcNAQkEMSIEIHLqX/4sZEj/ImyIVVryMF+OQzJxW+X++CEIJXMJ
# 1n8TMA0GCSqGSIb3DQEBAQUABIICALJXfbu3WYgFrnM2JKVK6ej7OytH3xJRy4NX
# NHiKTxP1qC/k6ZsVZz8aDCLtiMDxp5QYA1lO+cVDXS9hF6/6UWeZNBGDPFqhAwyJ
# dJhVklRp4Na42ipS2SWWPicO5DzehqebNYUDnV9BTK+Crl6OF9a8h6tu4Otogcm+
# aU5ITed7RV7A3iQa8jkNNEpYyeZ7OwqSVxX7DccQHcEM4VCX1q0xoC6n2KHwE80Z
# rh323J4zqwe75fkhi6cCfCQwX9JKfjcrHhvcYKFjUG2xF/SGb+PA0M2Rj4KBrBh+
# 1E5jhptN3ll1ziGo1HUEBPd5yc2dBoOi2NF9c/OO733W3avbXAD60JMS5Istlpip
# PyFhIoY55rqQ/UW+1YYdt/opZiCYmg+0FBoQiOYdvN8KnGRQm0s/fz8DHfphQc5x
# jyKd4KksAJJ0yHgZbem0r3wb0ob7ryegff+i9zH8owZvIu5CPRg2B807ps7+6YMf
# ZhkmNejEEyiDL23KUL8VoFG9krbjXVVC3Kjg938UN0UY/mEdlLLxLIfThLJBRF+E
# evt5NzMFHsSWHV0Bl9hThEumhirv26V1lD2D5GNcUNmjIoCtcMyO9wXDACnOYDNW
# kiy05m9cjrF//lW8ak+r46veVXgZ4+c1xNhLjy4KJlAtYL/0ieRfX7zTvcHF0z3G
# HITYQV9J
# SIG # End signature block
