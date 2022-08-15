Attribute VB_Name = "modVBHWldDriver"
Option Explicit
'******************************************************************************
'VB6 Driver Demo: Hello World Driver
'Author: Jon Johnson (fafalone)
'
'This is a demonstration of using VB to create a kernel-mode driver.
'All it does it return it's version information.
'
'Definitions mainly from The trick's TrickMemReader VB6 driver, as well the
'method for initializing the Driver Name and Link Name. I converted those
'to x64-compatible and added some of my own.
'https://www.vbforums.com/showthread.php?788179-VB6-Kernel-mode-driver
'
'Structure and functionality are heavily based on Geoff Chappell's self-sign
'driver example, as I followed (and recommend you follow) that method to test
'self-signing to run on x64 platforms.
'https://www.geoffchappell.com/notes/windows/license/customkernelsigners.htm
'
'You can also disable signature verification by rebooting and using the
'Advanced Boot Menu; or enable testsigning mode.
'******************************************************************************

Private Type HelloWorldVersion
    Major As Integer
    Minor As Integer
    Build As Integer
    Revision As Integer
End Type

Private Type DebugCSTR
    aCh(0 To 49) As Byte
End Type

Private Type LARGE_INTEGER
    lowPart As Long
    highPart As Long
End Type

Private Type UNICODE_STRING
    Length              As Integer
    MaximumLength       As Integer
    lpBuffer            As Long
End Type

Private Type LIST_ENTRY
    Flink               As Long
    Blink               As Long
End Type

Private Type KDEVICE_QUEUE
    Type                As Integer
    Size                As Integer
    DeviceListHead      As LIST_ENTRY
    Lock                As Long
    Busy                As Long
End Type

Private Type KDPC
    Type                As Byte
    Importance          As Byte
    Number              As Integer
    DpcListEntry        As LIST_ENTRY
    DeferredRoutine     As Long
    DeferredContext     As Long
    SystemArgument1     As Long
    SystemArgument2     As Long
    DpcData             As Long
End Type

Private Type DISPATCHER_HEADER
    Lock                As Long
    SignalState         As Long
    WaitListHead        As LIST_ENTRY
End Type

Private Type KEVENT
    Header              As DISPATCHER_HEADER
End Type

Private Type IO_STATUS_BLOCK
    StatusPointer       As Long
    Information         As Long
End Type

'Private Type IRP
'    Type                As Integer
'    Size                As Integer  '2
'    MdlAddress          As Long  '8
'    Flags               As Long     '16
'    AssociatedIrp       As Long  '24
'    ThreadListEntry     As LIST_ENTRY '32
'    IoStatus            As IO_STATUS_BLOCK '48
'    RequestorMode       As Byte      '64
'    PendingReturned     As Byte
'    StackCount          As Byte
'    CurrentLocation     As Byte
'    Cancel              As Byte
'    CancelIrql          As Byte
'    ApcEnvironment      As Byte
'    AllocationFlags     As Byte
'    UserIosb            As Long   '72
'    UserEvent           As Long   '80
'    #If Win64 Then
'    Overlay(15)         As Byte  '88
'    #Else
'    Overlay(7)          As Byte
'    #End If
'    CancelRoutine       As Long   '104
'    UserBuffer          As Long   '112
'    'Tail
'    DriverContext(3)    As Long   '120
'    Thread              As Long   '152
'    AuxiliaryBuffer     As Long   '160
'    ListEntry           As LIST_ENTRY '168
'    CurrentStackLocation As Long  '184
'    OriginalFileObject  As Long    '192
'    Pad                 As Currency 'Tail union alternate _KAPC is 8 bytes larger on both x86 and x64
'End Type 'Expected size (x64=0xD0) (x86=0x70)
Public Type Tail
    DriverContext(3)    As Long
    Thread              As Long
    AuxiliaryBuffer     As Long
    ListEntry           As LIST_ENTRY
    lpCurStackLocation  As Long
    OriginalFileObject  As Long
End Type

Public Type IRP
    Type                As Integer
    Size                As Integer
    MdlAddress          As Long
    Flags               As Long
    AssociatedIrp       As Long
    ThreadListEntry     As LIST_ENTRY
    IoStatus            As IO_STATUS_BLOCK
    RequestorMode       As Byte
    PendingReturned     As Byte
    StackCount          As Byte
    CurrentLocation     As Byte
    Cancel              As Byte
    CancelIrql          As Byte
    ApcEnvironment      As Byte
    AllocationFlags     As Byte
    UserIosb            As Long
    UserEvent           As Long
    Overlay             As Currency
    CancelRoutine       As Long
    UserBuffer          As Long
    Tail                As Tail
End Type
Private Type DEVICEIOCTL       'Comments: x64 alignments
'    #If Win64 Then
'    zPtrAlign1 As Long
'    #End If
    OutputBufferLength  As Long  '0x8
'    #If Win64 Then
'    zPtrAlign2 As Long
'    #End If
    InputBufferLength   As Long   '0x10
'    #If Win64 Then
'    zPtrAlign3 As Long
'    #End If
    IoControlCode       As Long   '0x18
    Type3InputBuffer    As Long '0x20
End Type

Private Type IO_STACK_LOCATION
    MajorFunction       As Byte       '0x0
    MinorFunction       As Byte       '0x1
    Flags               As Byte       '0x2
    Control             As Byte       '0x3
    DeviceIoControl     As DEVICEIOCTL '0x8
    DeviceObject        As Long    '0x28
    FileObject          As Long    '0x30
    CompletionRoutine   As Long    '0x38
    Context             As Long    '0x40
End Type 'Expected size (x64) = 0x48

Private Const IRP_MJ_MAXIMUM_FUNCTION As Long = &H1B

Private Type DRIVER_OBJECT
    Type                As Integer
    Size                As Integer
    DeviceObject        As Long
    Flags               As Long
    DriverStart         As Long
    DriverSize          As Long
    DriverSection       As Long
    DriverExtension     As Long
    DriverName          As UNICODE_STRING
    HardwareDatabase    As Long
    FastIoDispatch      As Long
    DriverInit          As Long
    DriverStartIo       As Long
    DriverUnload        As Long
    MajorFunction(IRP_MJ_MAXIMUM_FUNCTION)   As Long
End Type 'Expected size x64=0x150, x86=0xa8

Private Type DEVICE_OBJECT
    Type                As Integer
    Size                As Integer
    ReferenceCount      As Long
    DriverObject        As Long
    NextDevice          As Long
    AttachedDevice      As Long
    CurrentIrp          As Long
    Timer               As Long
    Flags               As Long
    Characteristics     As Long
    Vpb                 As Long
    DeviceExtension     As Long
    DeviceType          As Long
    StackSize           As Byte
'    #If Win64 Then
'    Queue(71)           As Byte 'This includes alignment padding, so be careful if copying out.
'    #Else
    Queue(39)           As Byte
'    #End If
    AlignRequirement    As Long
    DeviceQueue         As KDEVICE_QUEUE
    Dpc                 As KDPC
    ActiveThreadCount   As Long
    SecurityDescriptor  As Long
    DeviceLock          As KEVENT
    SectorSize          As Integer
    Spare1              As Integer
    DeviceObjExtension  As Long
    Reserved            As Long
End Type 'Expected size x64=0x150, x86=0xb8

Private Type FILE_OBJECT
    Type                 As Integer
    Size                 As Integer
    DeviceObject         As Long
    Vpb                  As Long
    FsContext            As Long
    FsContext2           As Long
    SectionObjectPointer As Long
    PrivateCacheMap      As Long
    FinalStatus          As Long
    RelatedFileObject    As Long
    LockOperation        As Byte
    DeletePending        As Byte
    ReadAccess           As Byte
    WriteAccess          As Byte
    DeleteAccess         As Byte
    SharedRead           As Byte
    SharedWrite          As Byte
    SharedDelete         As Byte
    Flags                As Long
    FileName             As UNICODE_STRING
    CurrentByteOffset    As LARGE_INTEGER
    Waiters              As Long
    Busy                 As Long
    LastLock             As KEVENT
    Event                As KEVENT
    CompletionContext    As Long
    IrpListLock          As Long
    IrpList              As LIST_ENTRY
    FileObjectExtension  As Long
End Type 'Expected size x64=0xD8, x86=0x80


Private Type BinaryString
    D(255)              As Integer
End Type

Private Const FILE_DEVICE_UNKNOWN    As Long = &H22
Private Const IO_NO_INCREMENT        As Long = &H0

Private Const IRP_MJ_CREATE = &H0
Private Const IRP_MJ_CREATE_NAMED_PIPE = &H1
Private Const IRP_MJ_CLOSE = &H2
Private Const IRP_MJ_READ = &H3
Private Const IRP_MJ_WRITE = &H4
Private Const IRP_MJ_QUERY_INFORMATION = &H5
Private Const IRP_MJ_SET_INFORMATION = &H6
Private Const IRP_MJ_QUERY_EA = &H7
Private Const IRP_MJ_SET_EA = &H8
Private Const IRP_MJ_FLUSH_BUFFERS = &H9
Private Const IRP_MJ_QUERY_VOLUME_INFORMATION = &HA
Private Const IRP_MJ_SET_VOLUME_INFORMATION = &HB
Private Const IRP_MJ_DIRECTORY_CONTROL = &HC
Private Const IRP_MJ_FILE_SYSTEM_CONTROL = &HD
Private Const IRP_MJ_DEVICE_CONTROL = &HE
Private Const IRP_MJ_INTERNAL_DEVICE_CONTROL = &HF
Private Const IRP_MJ_SHUTDOWN = &H10
Private Const IRP_MJ_LOCK_CONTROL = &H11
Private Const IRP_MJ_CLEANUP = &H12
Private Const IRP_MJ_CREATE_MAILSLOT = &H13
Private Const IRP_MJ_QUERY_SECURITY = &H14
Private Const IRP_MJ_SET_SECURITY = &H15
Private Const IRP_MJ_POWER = &H16
Private Const IRP_MJ_SYSTEM_CONTROL = &H17
Private Const IRP_MJ_DEVICE_CHANGE = &H18
Private Const IRP_MJ_QUERY_QUOTA = &H19
Private Const IRP_MJ_SET_QUOTA = &H1A
Private Const IRP_MJ_PNP = &H1B

Private Const METHOD_BUFFERED = &H0
Private Const FILE_ACCESS_ANY = &H0
Private Const IOCTL_HWRLD_VERSION As Long = &H80002000

Private DeviceName       As UNICODE_STRING   ' // Device name unicode string
Private DeviceLink       As UNICODE_STRING   ' // Device link unicode string
Private Device           As DEVICE_OBJECT    ' // Device object

Private strName As BinaryString     ' // Device name string
Private strLink As BinaryString     ' // Device link string

Private dbgsEntry As DebugCSTR 'Entry point success
Private dbgsDevIoEntry As DebugCSTR 'DeviceIoControl received

Private Const STATUS_SUCCESS = 0&
Private Const STATUS_INVALID_PARAMETER = &HC000000D
Private Const STATUS_INVALID_DEVICE_REQUEST = &HC0000010
Private Const STATUS_BUFFER_ALL_ZEROS = &H117
'Drivers can only import from ntoskrnl.exe. You need to be careful which language features you use; strings and arrays besides 1-D inside UDTs
'utilize APIs behind the scenes from other libraries (currently, I opened a feature request for non-SAFEARRAY arrays), and cannot be used in drivers.
'[ UseGetLastError (False) ]
'Private Declare PtrSafe Function DbgPrint CDecl Lib "ntoskrnl.exe" (ByVal Format As Long /* Currently unsupported in kernel mode: , ByVal ParamArray Args As Any()*/) As Long
'Private Declare PtrSafe Sub IoCompleteRequest Lib "ntoskrnl.exe" (pIrp As IRP, ByVal PriorityBoost As Byte)
'Private Declare PtrSafe Function IoCreateDevice Lib "ntoskrnl.exe" (DriverObject As DRIVER_OBJECT, ByVal DeviceExtensionSize As Long, DeviceName As UNICODE_STRING, ByVal DeviceType As Long, ByVal DeviceCharacteristics As Long, ByVal Exclusive As Long, DeviceObject As DEVICE_OBJECT) As Long
'Private Declare PtrSafe Function IoCreateSymbolicLink Lib "ntoskrnl.exe" (SymbolicLinkName As UNICODE_STRING, DeviceName As UNICODE_STRING) As Long
'Private Declare PtrSafe Function IoDeleteSymbolicLink Lib "ntoskrnl.exe" (SymbolicLinkName As UNICODE_STRING) As Long
'Private Declare PtrSafe Sub IoDeleteDevice Lib "ntoskrnl.exe" (DeviceObject As DEVICE_OBJECT)
'Private Declare PtrSafe Sub RtlInitUnicodeString Lib "ntoskrnl.exe" (DestinationString As UNICODE_STRING, SourceString As Any)
'Private Declare PtrSafe Sub CopyMemory Lib "ntoskrnl.exe" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Main()
End Sub

Private Sub InitDebugStrings()
dbgsEntry.aCh(0) = &H45: dbgsEntry.aCh(1) = &H6E: dbgsEntry.aCh(2) = &H74: dbgsEntry.aCh(3) = &H72: dbgsEntry.aCh(4) = &H79: dbgsEntry.aCh(5) = &H20
dbgsEntry.aCh(6) = &H70: dbgsEntry.aCh(7) = &H6F: dbgsEntry.aCh(8) = &H69: dbgsEntry.aCh(9) = &H6E: dbgsEntry.aCh(10) = &H74: dbgsEntry.aCh(11) = &H20
dbgsEntry.aCh(12) = &H73: dbgsEntry.aCh(13) = &H75: dbgsEntry.aCh(14) = &H63: dbgsEntry.aCh(15) = &H63: dbgsEntry.aCh(16) = &H65: dbgsEntry.aCh(17) = &H73
dbgsEntry.aCh(18) = &H73

dbgsDevIoEntry.aCh(0) = &H44: dbgsDevIoEntry.aCh(1) = &H65: dbgsDevIoEntry.aCh(2) = &H76: dbgsDevIoEntry.aCh(3) = &H69: dbgsDevIoEntry.aCh(4) = &H63
dbgsDevIoEntry.aCh(5) = &H65: dbgsDevIoEntry.aCh(6) = &H49: dbgsDevIoEntry.aCh(7) = &H6F: dbgsDevIoEntry.aCh(8) = &H43: dbgsDevIoEntry.aCh(9) = &H6F
dbgsDevIoEntry.aCh(10) = &H6E: dbgsDevIoEntry.aCh(11) = &H74: dbgsDevIoEntry.aCh(12) = &H72: dbgsDevIoEntry.aCh(13) = &H6F: dbgsDevIoEntry.aCh(14) = &H6C
dbgsDevIoEntry.aCh(15) = &H20: dbgsDevIoEntry.aCh(16) = &H72: dbgsDevIoEntry.aCh(17) = &H65: dbgsDevIoEntry.aCh(18) = &H63: dbgsDevIoEntry.aCh(19) = &H65
dbgsDevIoEntry.aCh(20) = &H69: dbgsDevIoEntry.aCh(21) = &H76: dbgsDevIoEntry.aCh(22) = &H65: dbgsDevIoEntry.aCh(23) = &H64
End Sub

Private Sub InitUnicodeStrings()
'\Device\VBHWldDrv
strName.D(0) = &H5C: strName.D(1) = &H44: strName.D(2) = &H65: strName.D(3) = &H76: strName.D(4) = &H69: strName.D(5) = &H63: strName.D(6) = &H65
strName.D(7) = &H5C: strName.D(8) = &H56: strName.D(9) = &H42: strName.D(10) = &H48: strName.D(11) = &H57: strName.D(12) = &H6C: strName.D(13) = &H64
strName.D(14) = &H44: strName.D(15) = &H72: strName.D(16) = &H76
RtlInitUnicodeString DeviceName, strName

'\DosDevices\VBHWldDrv
strLink.D(0) = &H5C: strLink.D(1) = &H44: strLink.D(2) = &H6F: strLink.D(3) = &H73: strLink.D(4) = &H44: strLink.D(5) = &H65: strLink.D(6) = &H76
strLink.D(7) = &H69: strLink.D(8) = &H63: strLink.D(9) = &H65: strLink.D(10) = &H73: strLink.D(11) = &H5C: strLink.D(12) = &H56: strLink.D(13) = &H42
strLink.D(14) = &H48: strLink.D(15) = &H57: strLink.D(16) = &H6C: strLink.D(17) = &H64: strLink.D(18) = &H44: strLink.D(19) = &H72: strLink.D(20) = &H76
RtlInitUnicodeString DeviceLink, strLink
End Sub

'Private Sub InitFuncs()
'IOCTL_HWRLD_VERSION = &H80002000 ' &H220004 'CTL_CODE(FILE_DEVICE_UNKNOWN, 1&, METHOD_BUFFERED, FILE_ACCESS_ANY)
'End Sub

Private Function NT_SUCCESS(ByVal Status As Long) As Boolean
    NT_SUCCESS = Status >= STATUS_SUCCESS
End Function

'Private Function CTL_CODE(ByVal DeviceType As Long, ByVal lFunction As Long, ByVal Method As Long, ByVal Access As Long) As Long
'    CTL_CODE = ((DeviceType << 16) Or (Access << 14) Or (lFunction << 2) Or Method)
'End Function

Private Function FARPROC(ByVal lpAdr As Long) As Long
FARPROC = lpAdr
End Function

Public Function IoGetCurrentIrpStackLocation(ByRef pIrp As IRP) As Long
IoGetCurrentIrpStackLocation = pIrp.Tail.lpCurStackLocation
End Function

Public Function DriverEntry(ByRef DriverObject As DRIVER_OBJECT, ByRef RegistryPath As UNICODE_STRING) As Long
    'InitDebugStrings
    'Dim dbgsPtr As Long
    'InterlockedExchange dbgsPtr, dbgsEntry
    'DbgPrint dbgsPtr
    InitUnicodeStrings
'    InitFuncs
'
    Dim ntStatus As Long

    ntStatus = IoCreateDevice(DriverObject, 0&, DeviceName, FILE_DEVICE_UNKNOWN, 0&, False, Device)
    If NT_SUCCESS(ntStatus) Then
        ntStatus = IoCreateSymbolicLink(DeviceLink, DeviceName)
        If Not NT_SUCCESS(ntStatus) Then
            IoDeleteDevice Device
            DriverEntry = ntStatus
            Exit Function
        End If

'        Dim i As Long
'        For i = 0 To IRP_MJ_MAXIMUM_FUNCTION
'                DriverObject.MajorFunction(i) = FARPROC(AddressOf OnOther)
'        Next
        DriverObject.MajorFunction(IRP_MJ_CREATE) = FARPROC(AddressOf OnCreate)
        DriverObject.MajorFunction(IRP_MJ_CLOSE) = FARPROC(AddressOf OnClose)
        DriverObject.MajorFunction(IRP_MJ_DEVICE_CONTROL) = FARPROC(AddressOf OnDeviceControl)

        DriverObject.DriverUnload = FARPROC(AddressOf OnUnload)
    End If

    DriverEntry = ntStatus
End Function

Public Function OnCreate(ByRef DriverObject As DRIVER_OBJECT, ByRef pIrp As IRP) As Long
'    Dim lpStack As Long
'    Dim ioStack As IO_STACK_LOCATION
'    Dim tFO As FILE_OBJECT
'    lpStack = IoGetCurrentIrpStackLocation(pIrp)
'    If lpStack Then
'        CopyMemory ioStack, ByVal lpStack, LenB(ioStack)
'        If ioStack.FileObject Then
'            CopyMemory tFO, ByVal ioStack.FileObject, LenB(tFO)
            Dim ntStatus As Long
'            If tFO.FileName.Length Then
                ntStatus = STATUS_SUCCESS
'            Else
'                ntStatus = STATUS_INVALID_PARAMETER
'            End If
            pIrp.IoStatus.Information = 0
            pIrp.IoStatus.StatusPointer = ntStatus
            IoCompleteRequest pIrp, IO_NO_INCREMENT
            OnCreate = ntStatus
'        End If
'    End If
End Function

Public Function OnClose(ByRef DriverObject As DRIVER_OBJECT, ByRef pIrp As IRP) As Long
    pIrp.IoStatus.Information = 0
    pIrp.IoStatus.StatusPointer = STATUS_SUCCESS
    IoCompleteRequest pIrp, IO_NO_INCREMENT
    OnClose = STATUS_SUCCESS
End Function

Public Function OnDeviceControl(ByRef DriverObject As DRIVER_OBJECT, ByRef pIrp As IRP) As Long
    'DbgPrint VarPtr(dbgsDevIoEntry)
    'Dim dbgsPtr As Long
    'InterlockedExchange dbgsPtr, dbgsDevIoEntry
    'DbgPrint dbgsPtr
    Dim lpStack As Long
    Dim ioStack As IO_STACK_LOCATION
    Dim ntStatus As Long

    pIrp.IoStatus.Information = 0
    lpStack = IoGetCurrentIrpStackLocation(pIrp)
    If lpStack Then
        CopyMemory ioStack, ByVal lpStack, Len(ioStack)
        If ioStack.DeviceIoControl.IoControlCode = IOCTL_HWRLD_VERSION And pIrp.AssociatedIrp <> 0 Then
                Dim tVer As HelloWorldVersion
                Dim lpBuffer As Long
                Dim cbIn As Long, cbOut As Long
                lpBuffer = pIrp.AssociatedIrp
                cbIn = ioStack.DeviceIoControl.InputBufferLength
                cbOut = ioStack.DeviceIoControl.OutputBufferLength
                If (lpBuffer = 0&) Or (cbIn <> 4) Or (cbOut <> LenB(tVer)) Then
                    If (lpBuffer = 0&) Or (cbOut <> LenB(tVer)) Then
                        ntStatus = STATUS_BUFFER_ALL_ZEROS
                    Else
                        ntStatus = STATUS_INVALID_PARAMETER
                    End If
                Else
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
        End If

    End If
    pIrp.IoStatus.Information = 0
    ntStatus = STATUS_INVALID_PARAMETER
    pIrp.IoStatus.StatusPointer = ntStatus
    IoCompleteRequest pIrp, IO_NO_INCREMENT
    OnDeviceControl = ntStatus
End Function

Public Function OnOther(ByRef DriverObject As DRIVER_OBJECT, ByRef pIrp As IRP) As Long
    Dim ntStatus As Long
    ntStatus = STATUS_INVALID_DEVICE_REQUEST
    pIrp.IoStatus.Information = 0
    pIrp.IoStatus.StatusPointer = ntStatus
    IoCompleteRequest pIrp, IO_NO_INCREMENT
    OnOther = ntStatus
End Function

Public Sub OnUnload(DriverObject As DRIVER_OBJECT)
    If Device.Size = 0 Then Exit Sub
    IoDeleteSymbolicLink DeviceLink
    IoDeleteDevice ByVal DriverObject.DeviceObject
End Sub

