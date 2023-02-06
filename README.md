# HelloWorldDriver
twinBASIC Kernel mode driver demo

![image](https://user-images.githubusercontent.com/7834493/214309562-2d4b4fe7-6273-4201-8a82-ac1bb6a89a5b.png)

This is a demonstration of using [twinBASIC](https://twinbasic.com/) (Current releases and community [on GitHub](https://github.com/twinbasic) to create a kernel mode driver compatible with x64 versions of Windows. I became fascinated with the idea after The_trick figured out [how to make them in VB6](https://www.vbforums.com/showthread.php?788179-VB6-Kernel-mode-driver), and when I saw twinBASIC could compile VB6 code to 64bit... the prospect was fascinating, since VB6 was limited to x86 and there's no WOW64 for kernel mode. We could make drivers for the Windows 64bit OS everyone is running, and even easier with far more features since tB has no runtime... if twinBASIC supported a few features to replicate the hacks that made it possible in VB6. So I made a [feature request](https://github.com/twinbasic/twinbasic/issues/1013), and the awesome Wayne Phillips was interested. 

For testing purposes, I first made a working VB6 version (included), and then a tB version with all the definitions updated to x64 and taking advantage of a feature tB has to put in-project API declares into , and a controller (written in tB) to load/unload the driver.

**Running the project**

*Build the binaries:*

The tB version only needs to be built. The required settings applied for making kernel mode drivers were creating a standard exe, removing the current references, enabling the settings Project: Native subsystem->YES, Project: Override entry point->DriverEntry, Project: Runtime binding of DLL declares->NO.

If you want to build the VB6 version to compare:
The VBP includes the undocumented link switches and enables the optimizations needed, so just needs to be opened and compiled.
After compiling the .sys, use The_trick's Patcher project (included) to strip the msvbvm60.dll dependency from the .sys.

*Running the driver*

I've been testing on Windows 7 via VM software, since it's less anal about unsigned drivers, and VMs because errors in drivers typically result in a bluescreen instead of error message or app crash. It's recommended you do the same, but not required. 

1. Microsoft has been heading down the road to where you don't own your computer, they do. Windows Vista and newer do not normally allow unsigned drivers, or even self-signed drivers. To get them signed, you have to provide a ton of personal information and pay hundreds of dollars. One way around this is via the Advanced Boot Menu:
You'll need to boot using the 'Disable driver signature enforcement' advanced boot option. Reboot and press F8 right when Windows starts loading. VMs like VirtualBox makes it nigh impossible to get an F8 keypress in, so instead, you can also open a command prompt as administrator, and enter the following: 
`bcdedit /set {globalsettings} advancedoptions true`
Then restart. This will bring up the Advanced Boot Options menu, containing "Disable driver signature enforcement". You need to do this every boot for which you want to load unsigned drivers; and unfortunately it disables them globally.
NOTE: When you start the driver, Windows will pop up a box saying a signature is required. But this doesn't mean the driver didn't load, as the log will show.

2. Create a folder for the project, and put TBHWldDrv.sys and HelloWorldDriverController.exe all in the same folder (and/or VBHWldDrv.sys if you're testing that too). The folder cannot be named TBHWldDrv.

3. Run HelloWorldDriverController.exe as administrator (it has a manifest option requiring this).

4. Click 'Load driver' (Connect is only for if it's already running, e.g. if it's been installed in the boot sequence, it isn't by default and isn't necessary). This will create a service for a kernel driver and start it.

5. If it successfully loads and connects (there's a log that will tell you), you can send the version command to get a response back from the driver.

6. When done, click Unload and delete. This will remove the service created to load the driver. Then Exit.

**How it works**

The first step is the DriverEntry function. Normally VB and tB exes have a hidden function that's the first to run (even before Sub Main), this is used to set up various things like COM. But a driver can't have any of that; it must enter through the DriverEntry function. You won't be able to use any APIs besides ones in `ntoskrnl.exe`, the Windows kernel module. This is because there's a barrier between user mode and kernel mode, and you can't load user mode DLLs into a kernel mode driver. 

In VB, this dramatically limited what you can do, because virtually everything relies on the msvbvm60.dll runtime. But twinBASIC doesn't rely on an external runtime; all of the VB runtime stuff is built right into the exe. It also provides an option to put API declares in the Import Address Table, in VB they're late-bound and called using runtime APIs and only added to the IAT if they're in a TLB. This makes programming drivers quite a bit easier, because for instance you can use `VarPtr` directly; in VB you'd need to use something like `InterlockedExchange` to copy the pointer into a new variable. 

There are some limitations. You still can't use strings or arrays (besides 1D arrays inside UDTs) because these are managed with APIs behind the scenes, so this project uses The_trick's BinaryString method of putting strings into the 1D UDT arrays that don't use SAFEARRAY and thus don't need APIs. But generally, there's a hell of a lot more twinBASIC lets you do. 

In the DriverEntry function we do 4 things: Initialize our string types, create the `DEVICE_OBJECT` that descibes our driver to the system, create a symbolic link that allows the driver to communicate with user mode apps via `CreateFile`, and set up the function table for the IRP major functions. For most of these, we just implement default handlers that pass on IRPs (I/O request packets) to the next driver in the stack. The one we're primarily interested in for this demo is `IRP_MJ_DEVICE_CONTROL`: If you've ever used the DeviceIoControl API, this is where those commands are going, to device drivers. 


    Public Function DriverEntry(ByRef DriverObject As DRIVER_OBJECT, ByRef RegistryPath As UNICODE_STRING) As Long
    InitDebugStrings
    DbgPrint VarPtr(dbgsEntry)
    InitUnicodeStrings
    InitFuncs

    Dim ntStatus As Long

    ntStatus = IoCreateDevice(DriverObject, 0&, DeviceName, FILE_DEVICE_UNKNOWN, 0&, False, Device)
    If NT_SUCCESS(ntStatus) Then
        ntStatus = IoCreateSymbolicLink(DeviceLink, DeviceName)
        If Not NT_SUCCESS(ntStatus) Then
            IoDeleteDevice Device
            DriverEntry = ntStatus
            Exit Function
        End If

       Dim i As Long
       For i = 0 To IRP_MJ_MAXIMUM_FUNCTION
               DriverObject.MajorFunction(i) = AddressOf OnOther
       Next
        DriverObject.MajorFunction(IRP_MJ_CREATE) = AddressOf OnCreate
        DriverObject.MajorFunction(IRP_MJ_CLOSE) = AddressOf OnClose
        DriverObject.MajorFunction(IRP_MJ_DEVICE_CONTROL) = AddressOf OnDeviceControl

        DriverObject.DriverUnload = AddressOf OnUnload
    End If
     
    DriverEntry = ntStatus
    End Function

We define a custom command for our project using the `CTL_CODE` macro, which can be implemented easily thanks to tB having << and >> bitshift operators. The command we're defining is used to send the driver a verification number to check to make sure everything is ok as input. The output is the `HelloWorldVersion` UDT, a custom structure we use for the driver to pass it's version data to us to test that everything is working and we've successfully created a driver that takes commands and exchanges data with user mode applications.

In the driver controller, after starting the driver with the service APIs and connecting to it with `CreateFile`, the command is sent:

    Dim tVer As HelloWorldVersion
    Dim lVerify As Long
    Dim cbRet As Long
    Dim result As Long
    AppendLog "Sending IOCTL_HWRLD_VERSION to driver..."
    result = DeviceIoControl(hDev, IOCTL_HWRLD_VERSION, lVerify, 4&, tVer, LenB(tVer), cbRet, ByVal 0&)
    AppendLog "Result: ret=0x" & Hex$(result) & ",cbRead=" & cbRet & vbCrLf & "Version (Expecting 1.2.3.4)=" & tVer.Major & "." & tVer.Minor & "." & tVer.Build & "." & tVer.Revision

If we get the version numbers we expect, SUCCESS! Everything has worked.

In the driver, we receive that command:

    Public Function OnDeviceControl(ByRef DriverObject As DRIVER_OBJECT, ByRef pIrp As IRP) As Long
    DbgPrint VarPtr(dbgsDevIoEntry)
    Dim lpStack As LongPtr
    Dim ioStack As IO_STACK_LOCATION
    Dim ntStatus As Long

    pIrp.IoStatus.Information = 0
    lpStack = IoGetCurrentIrpStackLocation(pIrp)
    If lpStack Then
        DbgPrint VarPtr(dbgsStackOk)
        CopyMemory ioStack, ByVal lpStack, LenB(ioStack)
        If (ioStack.DeviceIoControl.IoControlCode = IOCTL_HWRLD_VERSION) And (pIrp.AssociatedIrp <> 0) Then
            DbgPrint VarPtr(dbgsCmdOk)
            Dim tVer As HelloWorldVersion
            Dim lpBuffer As LongPtr
            Dim cbIn As Long, cbOut As Long
            lpBuffer = pIrp.AssociatedIrp
            cbIn = ioStack.DeviceIoControl.InputBufferLength
            cbOut = ioStack.DeviceIoControl.OutputBufferLength
            If (lpBuffer = 0&) Or (cbIn <> 4) Or (cbOut <> LenB(tVer)) Then
                If (lpBuffer = 0&) Then
                    ntStatus = STATUS_BUFFER_ALL_ZEROS
                    DbgPrint VarPtr(dbgsNoBuffer)
                ElseIf (cbOut <> LenB(tVer)) Then
                    ntStatus = STATUS_INVALID_BUFFER_SIZE
                    DbgPrint VarPtr(dbgsBadSize)
                Else
                    ntStatus = STATUS_INVALID_DEVICE_REQUEST
                    DbgPrint VarPtr(dbgsNoVal)
                End If
            Else
                DbgPrint VarPtr(dbgsCopyOut)
                tVer.Major = 1
                tVer.Minor = 2
                tVer.Build = 3
                tVer.Revision = 4
                CopyMemory ByVal lpBuffer, tVer, cbOut
                pIrp.IoStatus.Information = LenB(tVer)
            End If
            
            pIrp.IoStatus.StatusPointer = ntStatus
            IoCompleteRequest pIrp, IO_NO_INCREMENT
            OnDeviceControl = ntStatus
            Exit Function
            
The current location on the stack has all the DeviceIoControl arguments from when we called that API, we check if everything is ok, if so, we copy the data to the i/o packet.

The last thing we handle is the Unload function, where we delete the symbolic link and device object:

    Public Sub OnUnload(DriverObject As DRIVER_OBJECT)
        If Device.Size = 0 Then Exit Sub
        IoDeleteSymbolicLink DeviceLink
        IoDeleteDevice ByVal DriverObject.DeviceObject
    End Sub


This project shows the use of the DbgPrint function so we can more easily debug our driver like a normal VB/tB project instead of the extraordinarily difficult WinDbg's kernel debugger. It's a CDecl function, but fortunately tB supports that natively so we don't have to worry about the workarounds VB requires for that. It has a `...` argument that maps to a ParamArray; but right now tB manages those behind the scenes with APIs, so can't use them here... but those are all optional, we can supply just the string, then use [DebugView](https://docs.microsoft.com/en-us/sysinternals/downloads/debugview) to see the output. Note that you must open DebugView (as administrator) after clicking load driver, or it won't attach properly. 

**Self-signed kernel mode drivers**

According to Geoff Chappell, it's possible to use a bunch of undocumented functionality to load self-signed drivers under Windows 10/11 x64, with SecureBoot enabled, without disabling signature checks. [Have a read of his article on it](https://www.geoffchappell.com/notes/windows/license/customkernelsigners.htm?tx=19,21,26,35,38,39,41,52,53,56) if you're interested in trying it. This is not for the faint of heart however. It's extremely complicated. You'll need the Windows SDK installed for the signing tools, and Windows 10/11 Enterprise or, oddly, Education editions, to compile the policy files (though you may be able to use the ones he provides). 

**Requirements**

twinBASIC Beta 95 or newer (v0.15.95)
This will build for x86 or x64. You cannot use the x86 version of the driver on x64; though you could probably use the controller.
I've only tested this on Windows 10+.

**Thanks**

This project reuses many definitions and techniques from o.g. BASIC driver developer The_trick. The rest of it is heavily based on Geoff Chappell's SelfSign driver from the article above. I'm brand new to driver development myself, and had never attempted it before. If I can follow those projects and do it, so can you. :)
