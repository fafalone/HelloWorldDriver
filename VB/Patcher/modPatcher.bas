Attribute VB_Name = "modPatcher"
' // modPatcher.bas - module for removing runtime from VB6 applications
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

Private Type IMAGE_DOS_HEADER
    e_magic                     As Integer
    e_cblp                      As Integer
    e_cp                        As Integer
    e_crlc                      As Integer
    e_cparhdr                   As Integer
    e_minalloc                  As Integer
    e_maxalloc                  As Integer
    e_ss                        As Integer
    e_sp                        As Integer
    e_csum                      As Integer
    e_ip                        As Integer
    e_cs                        As Integer
    e_lfarlc                    As Integer
    e_ovno                      As Integer
    e_res(0 To 3)               As Integer
    e_oemid                     As Integer
    e_oeminfo                   As Integer
    e_res2(0 To 9)              As Integer
    e_lfanew                    As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUnitializedData       As Long
    AddressOfEntryPoint         As Long
    BaseOfCode                  As Long
    BaseOfData                  As Long
    ImageBase                   As Long
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    W32VersionValue             As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    SubSystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As Long
    SizeOfStackCommit           As Long
    SizeOfHeapReserve           As Long
    SizeOfHeapCommit            As Long
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
End Type

Private Type IMAGE_FILE_HEADER
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type

Private Type IMAGE_NT_HEADERS
    Signature                   As Long
    FileHeader                  As IMAGE_FILE_HEADER
    OptionalHeader              As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_IMPORT_DESCRIPTOR
    Characteristics             As Long
    TimeDateStamp               As Long
    ForwarderChain              As Long
    pName                       As Long
    FirstThunk                  As Long
End Type

Private Type IMAGE_BOUND_IMPORT_DESCRIPTOR
    TimeDateStamp               As Long
    OffsetModuleName            As Integer
    NumberOfModuleForwarderRefs As Integer
End Type

Private Type IMAGE_BOUND_FORWARDER_REF
    TimeDateStamp               As Long
    OffsetModuleName            As Integer
    Reserved                    As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress              As Long
    size                        As Long
End Type

Private Type IMAGE_SECTION_HEADER
    SectionName(7)              As Byte
    VirtualSize                 As Long
    VirtualAddress              As Long
    SizeOfRawData               As Long
    PointerToRawData            As Long
    PointerToRelocations        As Long
    PointerToLinenumbers        As Long
    NumberOfRelocations         As Integer
    NumberOfLinenumbers         As Integer
    Characteristics             As Long
End Type

Private Type OPENFILENAME
    lStructSize                 As Long
    hwndOwner                   As Long
    hInstance                   As Long
    lpstrFilter                 As Long
    lpstrCustomFilter           As Long
    nMaxCustFilter              As Long
    nFilterIndex                As Long
    lpstrFile                   As Long
    nMaxFile                    As Long
    lpstrFileTitle              As Long
    nMaxFileTitle               As Long
    lpstrInitialDir             As Long
    lpstrTitle                  As Long
    Flags                       As Long
    nFileOffset                 As Integer
    nFileExtension              As Integer
    lpstrDefExt                 As Long
    lCustData                   As Long
    lpfnHook                    As Long
    lpTemplateName              As Long
End Type

Private Type ptDat
    Prv1                        As Long
    Prv2                        As Long
End Type

Private Type LARGE_INTEGER
    lowpart                     As Long
    highpart                    As Long
End Type

Private Const IMAGE_DIRECTORY_ENTRY_IMPORT          As Long = 1
Private Const FILE_SHARE_READ                       As Long = &H1
Private Const GENERIC_READ                          As Long = &H80000000
Private Const GENERIC_WRITE                         As Long = &H40000000
Private Const FILE_MAP_WRITE                        As Long = &H2
Private Const FILE_MAP_READ                         As Long = &H4
Private Const PAGE_READWRITE                        As Long = &H4&
Private Const INVALID_HANDLE_VALUE                  As Long = -1
Private Const OPEN_EXISTING                         As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL                 As Long = &H80
Private Const IMAGE_DOS_SIGNATURE                   As Long = &H5A4D
Private Const IMAGE_NT_SIGNATURE                    As Long = &H4550&
Private Const IMAGE_NT_OPTIONAL_HDR32_MAGIC         As Long = &H10B&

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
                         Alias "GetOpenFileNameW" ( _
                         ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef dst As Any) As Long
Private Declare Function lstrcpyn Lib "kernel32" _
                         Alias "lstrcpynA" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any, _
                         ByVal iMaxLength As Long) As Long
Private Declare Function lstrlen Lib "kernel32" _
                         Alias "lstrlenA" ( _
                         ByRef lpString As Any) As Long
Private Declare Function CreateFile Lib "kernel32" _
                         Alias "CreateFileW" ( _
                         ByVal lpFileName As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwShareMode As Long, _
                         ByRef lpSecurityAttributes As Any, _
                         ByVal dwCreationDisposition As Long, _
                         ByVal dwFlagsAndAttributes As Long, _
                         ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" _
                         Alias "CreateFileMappingW" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileMappingAttributes As Any, _
                         ByVal flProtect As Long, _
                         ByVal dwMaximumSizeHigh As Long, _
                         ByVal dwMaximumSizeLow As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" ( _
                         ByVal hFileMappingObject As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwFileOffsetHigh As Long, _
                         ByVal dwFileOffsetLow As Long, _
                         ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" ( _
                         ByVal lpBaseAddress As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" ( _
                         ByRef lp As Any, _
                         ByVal ucb As Long) As Long
Private Declare Function CopyFile Lib "kernel32" _
                         Alias "CopyFileW" ( _
                         ByVal lpExistingFileName As Long, _
                         ByVal lpNewFileName As Long, _
                         ByVal bFailIfExists As Long) As Long
Private Declare Function CheckSumMappedFile Lib "imagehlp" ( _
                         ByRef BaseAddress As Any, _
                         ByVal FileLength As Long, _
                         ByRef HeaderSum As Long, _
                         ByRef CheckSum As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileSize As LARGE_INTEGER) As Long
                         
Private Declare Sub CopyMemory Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" _
                    Alias "RtlZeroMemory" ( _
                    ByRef Destination As Any, _
                    ByVal Length As Long)
                    
Dim base            As Long             ' // Module base address
Dim lpFirstSection  As Long             ' // Pointer to first section
Dim secCount        As Long             ' // Sections count
Dim dirCount        As Long             ' // Number of directories
Dim size            As LARGE_INTEGER    ' // Size of file

' // Shows open file dialog and returns selected file name
Public Function GetFile( _
                ByVal hwnd As Long) As String
    Dim ofn     As OPENFILENAME:    Dim Title   As String
    Dim Out     As String:          Dim Filter  As String
    Dim i       As Long

    Out = String(260, vbNullChar)
    Title = "Open file"
    Filter = "Executable file" & vbNullChar & "*.exe;*.dll;*.ocx;*.sys" & vbNullChar

    ofn.nMaxFile = 260
    ofn.hwndOwner = hwnd
    ofn.lpstrTitle = StrPtr(Title)
    ofn.lpstrFile = StrPtr(Out)
    ofn.lStructSize = Len(ofn)
    ofn.lpstrFilter = StrPtr(Filter)

    If GetOpenFileName(ofn) Then

        i = InStr(1, Out, vbNullChar, vbBinaryCompare)
        If i Then GetFile = Left$(Out, i - 1)

    End If

End Function

' // Remove runtime import (MSVBVM60) from module
Public Function RemoveRuntimeFromIAT( _
                ByRef FileName As String) As Boolean
    Dim iDir  As IMAGE_DATA_DIRECTORY
    Dim bDir  As IMAGE_DATA_DIRECTORY
    
    ' // Load PE to memory
    base = LoadPE(FileName)
    If base = 0 Then Exit Function
    ' // Create backup
    CreateBackup FileName

    ' // Get import directory
    If GetImportDir(iDir, bDir) Then

        ' // Remove runtime
        If ClearRuntimeImport(iDir) And ClearRuntimeBoundImport(bDir) Then
            
            ' // Update checksum
            If Not UpdateCheckSum() Then

                MsgBox "Error during updating checksum", vbExclamation

            Else

                RemoveRuntimeFromIAT = True

            End If

        End If

    End If

    ' // Close module
    ClosePE base

End Function

' // Load PE module to memory and return pointer to first byte
Private Function LoadPE( _
                 ByRef FileName As String) As Long
    Dim hFile   As Long:    Dim hMap    As Long

    ' // Open PE file
    hFile = CreateFile(StrPtr(FileName), GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ, _
                        ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If hFile = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to open the file '" & FileName & "'", vbCritical
        Exit Function
    End If

    ' // Obtain size of file
    GetFileSizeEx hFile, size

    If CBool(size.highpart) Or size.lowpart < 0 Then
        MsgBox "File too big"
        CloseHandle hFile
        Exit Function
    End If

    ' // Create file mapping
    hMap = CreateFileMapping(hFile, ByVal 0&, PAGE_READWRITE, 0, 0, 0)

    CloseHandle hFile

    If hMap = 0 Then
        MsgBox "Unable to create file mapping"
        Exit Function
    End If

    ' // Map file to memory
    LoadPE = MapViewOfFile(hMap, FILE_MAP_WRITE, 0, 0, 0)

    If LoadPE = 0 Then
        MsgBox "Unable to map file"
        Exit Function
    End If

    CloseHandle hMap

End Function

' // Unload module
Private Sub ClosePE( _
            ByVal base As Long)
    UnmapViewOfFile base
End Sub

' // Create backup file
Private Sub CreateBackup( _
            ByRef FileName As String)
    Dim Title   As String:  Dim Path    As String
    Dim Ext     As String:  Dim NewName As String

    Title = GetFileTitle(FileName)
    Path = GetFilePath(FileName)
    Ext = GetExtension(FileName)

    NewName = Path & Title

    ' // Try to copy file
    Do
        NewName = NewName & "_backup"
    Loop Until CopyFile(StrPtr(FileName), StrPtr(NewName & "." & Ext), True)

End Sub

' // Extract path from file name
Private Function GetFilePath( _
                 ByRef Path As String) As String
    Dim L As Long

    L = InStrRev(Path, "\")
    If L = Len(Path) Or L = 0 Then GetFilePath = Path: Exit Function
    GetFilePath = Mid$(Path, 1, L)

End Function

' // Extract file name from path
Private Function GetFileTitle( _
                 ByRef Path As String, _
                 Optional ByRef UseExtension As Boolean = False) As String
    Dim L   As Long:    Dim P   As Long

    L = InStrRev(Path, "\")
    If UseExtension Then P = Len(Path) + 1 Else P = InStrRev(Path, ".")

    If P > L Then
        L = IIf(L = 0, 1, L + 1)
        GetFileTitle = Mid$(Path, L, P - L)
    ElseIf P = L Then
        GetFileTitle = Path
    Else
        GetFileTitle = Mid$(Path, L + 1)
    End If

End Function

' // Extract file extension
Private Function GetExtension( _
                 ByRef Path As String) As String
    Dim L   As Long:    Dim P   As Long

    L = InStrRev(Path, "\")
    P = InStrRev(Path, ".")
    If P > L Then GetExtension = Mid$(Path, P + 1)

End Function

' // Get import data derictory of PE file
Private Function GetImportDir( _
                 ByRef iDir As IMAGE_DATA_DIRECTORY, _
                 ByRef bDir As IMAGE_DATA_DIRECTORY) As Boolean
    Dim dosHdr()    As IMAGE_DOS_HEADER:    Dim dPtr()      As Long
    Dim dOld        As ptDat

    ReDim dosHdr(0):    ReDim dPtr(0)

    ' // Set pointer to dos header
    dOld = PtGet(dPtr, Not Not dosHdr): dPtr(0) = base

    ' // Check memory permissions
    If IsBadReadPtr(ByVal dPtr(0), Len(dosHdr(0))) = 0 Then

        ' // Check signature and alignment
        If dosHdr(0).e_magic = IMAGE_DOS_SIGNATURE And (dosHdr(0).e_lfanew And &H3) = 0 Then

            Dim ntHdr()     As IMAGE_NT_HEADERS:    Dim nPtr()      As Long
            Dim nOld        As ptDat

            ReDim ntHdr(0):    ReDim nPtr(0)

            ' // Set pointer to NT headers
            nOld = PtGet(nPtr, Not Not ntHdr): nPtr(0) = base + dosHdr(0).e_lfanew

            ' // Check memory permissions
            If IsBadReadPtr(ByVal nPtr(0), Len(ntHdr(0))) = 0 Then

                ' // Check signature and size of optional header
                If ntHdr(0).Signature = IMAGE_NT_SIGNATURE And _
                   ntHdr(0).FileHeader.SizeOfOptionalHeader >= Len(ntHdr(0).OptionalHeader) Then

                    ' // Check bitness
                    If ntHdr(0).OptionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR32_MAGIC Then

                        ' // Check directories count
                        If ntHdr(0).OptionalHeader.NumberOfRvaAndSizes >= 0 Then
                            Dim lpDatDir    As Long

                            ' // Get pointer to first section
                            lpDatDir = nPtr(0) + Len(ntHdr(0))
                            dirCount = ntHdr(0).OptionalHeader.NumberOfRvaAndSizes
                            secCount = ntHdr(0).FileHeader.NumberOfSections
                            
                            lpFirstSection = lpDatDir + Len(iDir) * dirCount

                            ' // Get import directory
                            If dirCount > 1 Then
                                
                                ' // Move to import data directory
                                lpDatDir = lpDatDir + Len(iDir)

                                ' // Check memory permissions
                                If IsBadReadPtr(ByVal lpDatDir, Len(iDir)) = 0 Then
                                    CopyMemory iDir, ByVal lpDatDir, Len(iDir)
                                End If
                                
                                If dirCount > 11 Then
                                
                                    ' // Move to bound import directory
                                    lpDatDir = lpDatDir + Len(iDir) * 10
                                
                                    ' // Check memory permissions
                                    If IsBadReadPtr(ByVal lpDatDir, Len(iDir)) = 0 Then
                                        CopyMemory bDir, ByVal lpDatDir, Len(iDir)
                                        GetImportDir = True
                                    End If
                                
                                End If
                                
                            Else

                                MsgBox "Import directory not found"

                            End If

                        End If

                    End If

                End If

            End If

            ' // Release NT headers pointer
            PtRelease nPtr, nOld

        End If

    End If

    ' // Release dos header pointer
    PtRelease dPtr, dOld

End Function

' // Remove runtime import and functions name
Private Function ClearRuntimeImport( _
                 ByRef iDir As IMAGE_DATA_DIRECTORY) As Boolean
    Dim pDsc()      As Long:                    Dim dsc()       As IMAGE_IMPORT_DESCRIPTOR
    Dim od          As ptDat:                   Dim ptr         As Long
    Dim rva         As Long:                    Dim sz          As Long
    Dim found       As Boolean:                 Dim iat         As Long
    Dim prev        As IMAGE_IMPORT_DESCRIPTOR

    ReDim pDsc(0):  ReDim dsc(0)

    rva = iDir.VirtualAddress
    sz = iDir.size
    ptr = RVA2RAW(rva) + base

    ' // If import directory is presented
    If sz > 0 And rva > 0 Then

        Dim i       As Long:    Dim Name    As String

        ' // Set pointer to import descriptor
        od = PtGet(pDsc, Not Not dsc): pDsc(0) = ptr

        ' // Go thru each imported library
        Do Until dsc(0).Characteristics = 0 And _
                 dsc(0).FirstThunk = 0 And _
                 dsc(0).ForwarderChain = 0 And _
                 dsc(0).pName = 0 And _
                 dsc(0).TimeDateStamp = 0

            ' // If name is presented
            If dsc(0).pName Then

                ' // Get library name
                ptr = RVA2RAW(dsc(0).pName) + base
                Name = GetString(ptr)

                ' // Compare with runtime name
                If StrComp(Name, "MSVBVM60.DLL", vbTextCompare) = 0 Then

                    found = True
                    rva = IIf(dsc(0).Characteristics, dsc(0).Characteristics, dsc(0).FirstThunk)
                    iat = dsc(0).FirstThunk
                    ClearFunctions rva, iat
                    ClearRuntimeImport = True

                End If


            End If

            i = i + 1

            ' // Next import descriptor
            pDsc(0) = pDsc(0) + Len(dsc(0))

            If found Then

                ' // Shift descriptor to free place
                CopyMemory ByVal pDsc(0) - Len(dsc(0)), dsc(0), Len(dsc(0))

            End If

        Loop

        ' // Release import descriptor pointer
        PtRelease pDsc, od

    End If
    
    ClearRuntimeImport = True
    
End Function

' // Delete all functions names
Private Function ClearFunctions( _
                 ByVal rva As Long, _
                 ByVal iat As Long) As Boolean
    Dim pTnk()      As Long:    Dim thnk()      As Long
    Dim ot          As ptDat:   Dim Name        As String
    Dim ptr         As Long

    ReDim pTnk(0):  ReDim thnk(0)

    ' // Set pointer to IMAGE_THUNK_DATA
    ot = PtGet(pTnk, Not Not thnk)

    pTnk(0) = RVA2RAW(rva) + base

    Do While thnk(0)

        If thnk(0) > 0 Then

            ' // If import by name then get pointer to string
            ptr = RVA2RAW(thnk(0)) + base + 2
            ' // Get string
            Name = GetString(ptr)
            ' // Zero place
            ZeroMemory ByVal ptr - 2, Len(Name) + 2

        End If

        ' // Zero thunk
        thnk(0) = 0
        pTnk(0) = pTnk(0) + 4

    Loop

    ' // Release pointer
    PtRelease pTnk, ot

End Function

' // Remove runtime bound import
Private Function ClearRuntimeBoundImport( _
                 ByRef bDir As IMAGE_DATA_DIRECTORY) As Boolean
    Dim pDsc()      As Long:                            Dim dsc()       As IMAGE_BOUND_IMPORT_DESCRIPTOR
    Dim od          As ptDat:                           Dim ptr         As Long
    Dim rva         As Long:                            Dim sz          As Long
    Dim descAddr    As Long:                            Dim iat         As Long
    Dim start       As Long:                            Dim fwdRef      As IMAGE_BOUND_FORWARDER_REF
    Dim i           As Long:                            Dim Name        As String
    Dim descSize    As Long
    
    ReDim pDsc(0):  ReDim dsc(0)

    rva = bDir.VirtualAddress
    sz = bDir.size
    start = RVA2RAW(rva) + base

    ' // If bound import directory is presented
    If sz > 0 And rva > 0 Then

        ' // Set pointer to bound import descriptor
        od = PtGet(pDsc, Not Not dsc): pDsc(0) = start

        ' // Go thru each imported library
        Do Until dsc(0).NumberOfModuleForwarderRefs = 0 And _
                 dsc(0).OffsetModuleName = 0 And _
                 dsc(0).TimeDateStamp = 0

            ' // If name is presented
            If dsc(0).OffsetModuleName Then

                ' // Get library name
                ptr = start + dsc(0).OffsetModuleName
                Name = GetString(ptr)

                ' // Compare with runtime name
                If StrComp(Name, "MSVBVM60.DLL", vbTextCompare) = 0 Then

                    descAddr = pDsc(0)
                    descSize = descSize + Len(dsc(0)) + Len(fwdRef) * dsc(0).NumberOfModuleForwarderRefs
                    
                End If


            End If

            i = i + 1
            
            ' // Next bound import descriptor
            pDsc(0) = pDsc(0) + Len(dsc(0)) + Len(fwdRef) * dsc(0).NumberOfModuleForwarderRefs
            
        Loop
        
        CopyMemory ByVal descAddr, ByVal descAddr + descSize, pDsc(0) - descAddr - descSize + 8
        
        ' // Release import descriptor pointer
        PtRelease pDsc, od

    End If
    
    ClearRuntimeBoundImport = True
    
End Function

' // Update checksum of PE
Private Function UpdateCheckSum() As Boolean
    Dim chksum  As Long:    Dim lpNtHdr As Long

    lpNtHdr = CheckSumMappedFile(ByVal base, size.lowpart, 0, chksum)

    If lpNtHdr Then

        ' // Store checksum to header
        GetMem4 chksum, ByVal lpNtHdr + &H58
        UpdateCheckSum = True

    End If

End Function

' // Relative virtual address to raw offset
Private Function RVA2RAW( _
                 ByVal rva As Long) As Long
    Dim i           As Long:                    Dim pSec()      As Long
    Dim sec()       As IMAGE_SECTION_HEADER:    Dim os          As ptDat

    ReDim pSec(0):  ReDim sec(0)

    ' // Set pointer to section header
    os = PtGet(pSec, Not Not sec): pSec(0) = lpFirstSection

    For i = 0 To secCount - 1

        If rva >= sec(0).VirtualAddress And rva < sec(0).VirtualAddress + sec(0).VirtualSize Then

            RVA2RAW = sec(0).PointerToRawData + (rva - sec(0).VirtualAddress)
            PtRelease pSec, os

            Exit Function

        End If

        pSec(0) = pSec(0) + Len(sec(0))

    Next

    RVA2RAW = rva

    ' // Release pointer
    PtRelease pSec, os

End Function

' // Get ASCII string by pointer
Private Function GetString( _
                 ByVal ptr As Long) As String
    Dim L   As Long

    L = lstrlen(ByVal ptr)

    If L Then

        GetString = Space(L)
        If lstrcpyn(ByVal GetString, ByVal ptr, L + 1) = 0 Then GetString = vbNullString

    End If

End Function

' // Create SA pointer
Private Function PtGet( _
                 ByRef Pointer() As Long, _
                 ByVal VarAddr As Long) As ptDat
    Dim i As Long

    i = (Not Not Pointer) + &HC
    GetMem4 ByVal i, PtGet.Prv1
    GetMem4 VarAddr + &HC, ByVal i
    PtGet.Prv2 = Pointer(0)

End Function

' // Release SA pointer
Private Sub PtRelease( _
            ByRef Pointer() As Long, _
            ByRef prev As ptDat)

    Pointer(0) = prev.Prv2
    GetMem4 prev.Prv1, ByVal (Not Not Pointer) + &HC

End Sub

