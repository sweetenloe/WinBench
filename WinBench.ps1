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

function gpubench($dir,$sec=8,$elevated=$false){
 say "GPU test: Direct3D ALU shader workload, $sec sec."
 $winsat=Get-Command WinSAT.exe -ea silentlycontinue|select -first 1
 if(!$winsat){
  return [pscustomobject]@{Test='GPU Direct3D (WinSAT)';RequestedSeconds=$sec;Error='WinSAT.exe is not available on this Windows installation.'}
 }
 if(!$elevated){
  return [pscustomobject]@{Test='GPU Direct3D (WinSAT)';RequestedSeconds=$sec;Error='GPU test requires an elevated PowerShell session. Run WinBench as Administrator.'}
 }

 $xml=join-path $dir ("winbench-gpu-"+[guid]::NewGuid().ToString('N')+'.xml')
 $args=@(
  'd3d','-totalobj','20','-objs','C(20)','-totaltex','10','-texpobj','C(1)',
  '-alushader','-noalpha','-NoDisp','-fixedseed','-time',[string]$sec,'-xml',$xml
 )
 $raw=@();$exitCode=$null
 try{
  $old=$ErrorActionPreference;$ErrorActionPreference='Continue'
  try{$raw=@(& $winsat.Source @args 2>&1|%{$_.ToString()});$exitCode=$LASTEXITCODE}
  finally{$ErrorActionPreference=$old}

  if($exitCode-ne 0){throw "WinSAT exited with code $exitCode. $($raw -join ' ')"}
  if(!(test-path -literalpath $xml)){throw 'WinSAT completed without producing its XML result.'}

  [xml]$doc=Get-Content -literalpath $xml -Raw
  $nodes=@($doc.SelectNodes("//*[local-name()='Results'][*[local-name()='FPS']]"))
  $fps=@($nodes|%{[double]$_.FPS}|?{$_-ge 0})
  if(!$fps.Count){throw 'WinSAT did not return a Direct3D frame-rate result.'}
  $m=$fps|measure -Average -Minimum -Maximum
  [long]$frames=($nodes|%{if($_.FramesRendered){[long]$_.FramesRendered}else{0}}|measure -Sum).Sum
  [pscustomobject]@{
   Test='GPU Direct3D (WinSAT)';Workload='ALU shader';RequestedSeconds=$sec;Subtests=$fps.Count
   AverageFPS=[math]::Round($m.Average,2);MinimumFPS=[math]::Round($m.Minimum,2);MaximumFPS=[math]::Round($m.Maximum,2)
   FramesRendered=$frames;ExitCode=$exitCode
  }
 }catch{
  [pscustomobject]@{Test='GPU Direct3D (WinSAT)';RequestedSeconds=$sec;ExitCode=$exitCode;Error=$_.Exception.Message}
 }finally{
  if(!$KeepTemp-and(test-path -literalpath $xml)){rm -literalpath $xml -Force -ea silentlycontinue}
 }
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
<li>The GPU benchmark is an off-screen Direct3D ALU-shader workload provided by WinSAT and requires an elevated PowerShell session.</li>
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
  if($mode-eq'Quick'){$cpuSecs=8;$memPasses=10;$diskMB=256;$cpuTele=10;$gpuSecs=8}
  else{$cpuSecs=20;$memPasses=24;$diskMB=1024;$cpuTele=22;$gpuSecs=15}

  hdr 'CPU Benchmark'
  $pair=cpuwithtele {cpubench $cpuSecs $sys.LogicalProcessors} $cpuTele $nv $SampleInterval
  $c=$pair.Result;$tele+=@($pair.Telemetry);$bench+=$c
  good ("CPU: {0:N2} hashes/sec | {1:N2} MiB/sec"-f$c.HashesPerSecond,$c.ThroughputMiBPerSec)

  hdr 'Memory Benchmark'
  $m=membench 64 $memPasses;$bench+=$m
  if($m.PSObject.Properties['ThroughputMiBps']){good ("Memory copy: {0:N2} MiB/sec"-f$m.ThroughputMiBps)}else{warnx "Memory test failed: $($m.Error)"}

  hdr 'GPU Benchmark'
  $pair=cpuwithtele {gpubench $tmp $gpuSecs $sys.IsElevated} ($gpuSecs+8) $nv $SampleInterval
  $g=$pair.Result;$tele+=@($pair.Telemetry);$bench+=$g
  if($g.PSObject.Properties['AverageFPS']){good ("GPU Direct3D: {0:N2} average FPS | {1:N0} frames"-f$g.AverageFPS,$g.FramesRendered)}else{warnx "GPU test skipped or failed: $($g.Error)"}

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
# MIIFswYJKoZIhvcNAQcCoIIFpDCCBaACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqFQT5CbN0GBYc
# 7TRkjqQvC6ks2JFCCNcLoA/2PZahUaCCAx4wggMaMIICAqADAgECAhBIeTa7jvWo
# qESazAVCWp1fMA0GCSqGSIb3DQEBCwUAMCUxIzAhBgNVBAMMGk15IFBvd2VyU2hl
# bGwgQ29kZSBTaWduaW5nMB4XDTI2MDgxMTAyMDAxOVoXDTI3MDgxMTAyMjAxOVow
# JTEjMCEGA1UEAwwaTXkgUG93ZXJTaGVsbCBDb2RlIFNpZ25pbmcwggEiMA0GCSqG
# SIb3DQEBAQUAA4IBDwAwggEKAoIBAQDJEVZvaSzkYqut0cHABvcjRhQmYSjslYYk
# LhHaK99IizoNS0xL3wz4ekqc5akykxEvejcA6LeEU4aXk7Myc5IIrB7XV+cUAcFn
# G0AD95zgKX5fHz73db7IaWBwbThtbO0DKAI613QtKGyNmMWUPeJctjuoU6QRn+wq
# V4vEL8hygWKJhoM/oy3iGOYwgRQawbDfbcUREgErBryMzUuggqwRHTSmrxAVYepT
# YqaAYUKnRhm3KEPPsbveGAUj033SRIZm9TR0kd8w60FwgwbXo7iyUh4Q+2B3mKi1
# De10PxMz1D7x2EshEuxxP/Z5cV+rQFA1ZlUyeikKLgNTlCgWiQBBAgMBAAGjRjBE
# MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQU
# WYD3n7z+EmAKbwoggzlpvk5UEq8wDQYJKoZIhvcNAQELBQADggEBAKaazuSHnaDp
# D4wrOkN/LvdMgatsROKSuZDXAFogxtjMmEsk7TpV2RkdjWyea9Eg2NaAWMLvhfGc
# TEyjpPruf2Y2f3EpnwuCnifHi/Ka5QSK9fLpIfBUFn5k6N1drI+KYu/nfTHGidw8
# UVrQCtxB241r2EpoAs4R6aikj/MDu+G+fgTd+zILGxvBuTYBPkcKHCTWyOoqhsQf
# rgnWxGvVUVdLNk/DpLFW3TkIalRK4EWtJFAHz23i0Zxe8jAr4XaCekwaEg+IfNNj
# 5XjVouVkkSFaI0CPiBlj0rVdGlPM3iYi+eCTDHmWDI6LEZ77mhS2tP9ursNGAsgB
# GOlVke42P1sxggHrMIIB5wIBATA5MCUxIzAhBgNVBAMMGk15IFBvd2VyU2hlbGwg
# Q29kZSBTaWduaW5nAhBIeTa7jvWoqESazAVCWp1fMA0GCWCGSAFlAwQCAQUAoIGE
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkE
# MSIEIPa05p5N944nytJ+13DRdx2Xl/B7NjTC7/+iusY7JJ4oMA0GCSqGSIb3DQEB
# AQUABIIBACBE5gm60frUKZGLFC/mJC4xFnvqYYXykBWweS5+zXtDAswEkDygZJvx
# t0WqSSHuWG/XIhyd28IvJeSXN/v/KpkA7g2zEB0u7hj9yTDdpgxpgrlGb4bjQLlN
# XgJjwnJFgzKyJ192D3HnAva9RiShJ0cWyv3H0fqBQbiw3O8i233M64jmGpmNrfUm
# zE6lu/H5Ex+jv7gRuPXUhGcGftZNFE+8R8snSRabQZFxnS/1rIk56GWEh7Toyq3j
# 1F/KowunowmgxBYaEzpVStdl74vVVAdB6IGXRWib1PAfsn3aFSbM+IRMqIf9deYo
# CzUq4XuvELqiVyTJkf3E61MActzt8NU=
# SIG # End signature block
