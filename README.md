# HelloWorldDriver
twinBASIC Kernel mode driver demo

This is a demonstration of using twinBASIC to create a kernel mode driver. I became fascinated with the idea after The_trick figured out how to make them in VB6, and when I saw twinBASIC could compile VB6 code to 64bit... the prospect was fascinating, since VB6 was limited to x86 and there's no WOW64 for kernel mode. We could make drivers for the Windows 64bit OS everyone is running, and even easier with far more features since tB has no runtime... if twinBASIC supported a few features to replicate the hacks that [made it possible in VB6](https://www.vbforums.com/showthread.php?788179-VB6-Kernel-mode-driver). So I made a [feature request](https://github.com/twinbasic/twinbasic/issues/1013), and the awesome Wayne Phillips was interested. There's still some issues with implementation, so the tB version doesn't work yet, but I needed to post the current code to help with debugging.

For testing purposes, I made a working VB6 version, a tB version with near-identical code (meaning x86-only for now), and a controller (written in tB) to load/unload the driver. There's a public release including binaries, with the VB6 version already patched to remove the msvbvm60 reference.

To use the project, 

*Build binaries, if not using release*
For the VB version:
1) The VBP includes the undocumented link switches and enables the optimizations needed, so just needs to be opened and compiled.
2) After compiling the .sys, use The_trick's Patcher project (included) to strip the msvbvm60.dll dependency from the .sys.
3) Build, for win32, the HelloWorldDriverController twinBASIC project. This is used to load both the VB6 version, and hopefully one day soon, the tB version.

The tB version only needs to be built. The required settings applied for making kernel mode drivers were creating a standard exe, removing the current references, enabling the settings Project: Native subsystem->YES, Project: Override entry point->DriverEntry, Project: Runtime binding of DLL declares->NO.

*Running the driver*
1) First, you need an x86 version of Windows. You can set one up in VirtualBox or other VM software. I've been testing on Windows 7, since it's less anal about unsigned drivers.
2) After you set up x86 Windows, you'll need to boot it using the 'Disable driver signature enforcement' advanced boot option (after this demo is fully working, it will include an undocumented method of installing self-signed drivers in Windows 10/11). VirtualBox makes it nigh impossible to get an F8 keypress in, so instead, open a command prompt as administrator, and enter the following:
`bcdedit /set {globalsettings} advancedoptions true`
Then restart. This will bring up the Advanced Boot Options menu, containing "Disable driver signature enforcement". The bcdedit command must be done before every restart.
3) Create a folder for the project, and put VBHWldDrv.sys, TBHWldDrv.sys (if testing that), and HelloWorldDriverController.exe all in the same folder.
4) Run HelloWorldDriverController.exe as administrator (it has a manifest option requiring this). 
5) If you're loading the VB version, above the big textbox there's a box with TBHWldDrv, change that to VBHWldDrv, and click set.
6) Click 'Load driver' (Connect is only for if it's already running, e.g. if it's been installed in the boot sequence, it isn't by default and isn't neccessary). This will create a service for a kernel driver and start it.
7) If it successfully loads and connects (there's a log that will tell you), you can send the version command to get a response back from the driver.
8) When done, click Unload and delete. This will remove the service created to load the driver. Then Exit.

Currently, with the tB version, step 6 will fail with invalid exe if you use the included binary, or if you rebuild with the latest tB your OS will blue screen with an access violation.  Wayne is looking into this. But the VB version is working.
