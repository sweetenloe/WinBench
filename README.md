# WinBench
### *Cause fuck paying a company for what PowerShell can already do*

GPU Test: `low risk`: short duration, normal Direct3D calls, and no changes to voltage, clocks, firmware, or drivers.

The CPU test is also low risk, but it deliberately loads every logical processor at close to 100%:
`Quick: 8 seconds`
`Full: 20 seconds`
No voltage, clock, BIOS, or power-limit changes

The script cannot reliably monitor CPU temperature on every Windows PC, so **stop if fans sound abnormal, you smell anything unusual, or the machine is already overheating.**

The RAM test is very **mild**. It allocates two `64 MiB` buffers—about `128 MiB` total—and repeatedly copies between them:
`Quick: 10 passes`
`Full: 24 passes`

**It doesn’t alter RAM timings, voltage, BIOS settings, or other files. On a healthy system it’s essentially harmless. If memory is already faulty, it might expose a crash or error, but it won’t damage the RAM.** 

#### It’ll push your PC for a few seconds, but it won’t change voltages or overclock anything. Healthy hardware will be fine—unstable or overheating hardware might crash or shut down like any other stress-tester. 

Here's every stress-testing Terms & Conditions in a nutshell:
** Run it at your own risk. **

### Future Additions:
- maybe utilizing a pip lib for benchmark standards
- maybe even using your telemetry data to make even more money off you like everyone else (*kidding it's just a powershell script, duh*)
- maybe having a shitty Terms & Conditions that 
