Attribute VB_Name = "modMemory"
Option Explicit

Private Const MEM_PRIVATE& = &H20000
Private Const MEM_COMMIT& = &H1000
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 260
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPALL = &HF

Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function VirtualQueryEx& Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


'NT, 2K, XP

Public Sub GetProcessesNT() 'Enumerates Processes
  Dim lngSnapShot&, uProcess As PROCESSENTRY32, lngContinue&, lngPID&, lstItem As ListItem
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

  If lngSnapShot <> 0 Then
    uProcess.dwSize = Len(uProcess)
    lngContinue = Process32First(lngSnapShot, uProcess) 'Get the process to start off with

      Do While lngContinue 'Do while there are processes
          lngPID = uProcess.th32ProcessID 'Get Process Id
            
            Set lstItem = frmProcesses.lvwMain.ListItems.Add(, , Left(uProcess.szExeFile, InStr(1, uProcess.szExeFile, vbNullChar) - 1)) 'Add process name
              lstItem.SubItems(1) = lngPID 'Add Process Id
              
                If GetMemoryUsageNT(lngPID) > 0 Then  'Get memory usage from Process Id
                  lstItem.SubItems(2) = Format(GetMemoryUsageNT(lngPID), "###,###") & " Kb" 'Can be determined
                Else
                  lstItem.SubItems(2) = "Unknown" 'Cannot be determined
                  
                  'Make bold to make it easier to see
                  lstItem.Bold = True
                  lstItem.ListSubItems.Item(1).Bold = True
                  lstItem.ListSubItems.Item(2).Bold = True
                End If
                
          lngContinue = Process32Next(lngSnapShot, uProcess) 'Get the next process to work with
        DoEvents 'So it doesnt hang
      Loop
    CloseHandle (lngSnapShot) 'Close process handle
  End If
End Sub

Private Function GetMemoryUsageNT(ProcID&) As String
  Dim hProcess&, ProcMemCounter As PROCESS_MEMORY_COUNTERS
    ProcMemCounter.cb = LenB(ProcMemCounter)
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcID&) 'Open a process
        GetProcessMemoryInfo hProcess, ProcMemCounter, ProcMemCounter.cb 'Heres where it gets the memory
    GetMemoryUsageNT = ProcMemCounter.WorkingSetSize / 1024 'Sets the function return to the memory usage
  CloseHandle hProcess 'Close process handle
End Function



'9x and ME

Public Sub GetProcesses9X()
  Dim hSnapshot&, ProcessInfo As PROCESSENTRY32, lngContinue&, strExeName$, hProcess&, lngMemory&, lstItem As ListItem
    
    ProcessInfo.dwSize = Len(ProcessInfo)
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0)
    ProcessInfo.dwSize = Len(ProcessInfo)
    lngContinue = Process32First(hSnapshot, ProcessInfo)

    While lngContinue <> 0
        strExeName = Left(ProcessInfo.szExeFile, InStr(ProcessInfo.szExeFile, vbNullChar) - 1)
        strExeName = GetFileName(strExeName)
             
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessInfo.th32ProcessID)
        lngMemory = GetProcessMemUsage9X(hProcess, strExeName)
        
            Set lstItem = frmProcesses.lvwMain.ListItems.Add(, , strExeName) 'Add process name
              lstItem.SubItems(1) = ProcessInfo.th32ProcessID
                If lngMemory > 0 Then  'Get memory usage from Process Id
                  lstItem.SubItems(2) = Format(lngMemory, "###,###") & " Kb" 'Can be determined
                Else
                  lstItem.SubItems(2) = "Unknown" 'Cannot be determined
                  lstItem.Bold = True
                  lstItem.ListSubItems.Item(1).Bold = True
                  lstItem.ListSubItems.Item(2).Bold = True
                End If
        
        ProcessInfo.dwSize = Len(ProcessInfo)
        lngContinue = Process32Next(hSnapshot, ProcessInfo)
      DoEvents
    Wend
  CloseHandle (hSnapshot)
End Sub

Private Function GetProcessMemUsage9X(hProcess As Long, exe As String) As Long
    Dim lngMem#, lngPrivateBytes#, lngReturn#, lngLenMbi#, lngProcess#, Si As SYSTEM_INFO, MBI As MEMORY_BASIC_INFORMATION

    Call GetSystemInfo(Si)
    
    lngLenMbi = Len(MBI)
    lngMem = Si.lpMinimumApplicationAddress
    
    While lngMem < Si.lpMaximumApplicationAddress
      MBI.RegionSize = 0
      lngReturn = VirtualQueryEx(hProcess, lngMem, MBI, lngLenMbi)
        If lngReturn = lngLenMbi Then
            If ((MBI.lType = MEM_PRIVATE) And (MBI.State = MEM_COMMIT)) Then  ' this block is In use by this process
              lngPrivateBytes = lngPrivateBytes + MBI.RegionSize
            End If
          lngMem = MBI.BaseAddress + MBI.RegionSize
        Else
          Exit Function
        End If
      DoEvents
    Wend
  GetProcessMemUsage9X = CStr(lngPrivateBytes / 1024)
End Function

'Turns C:\Windows\Explorer.exe to Explorer.exe

Private Function GetFileName(FullPath As String) As String
  On Error Resume Next
    Dim strConverted$, strChr$, strLength&, intIncrement%
      strLength = Len(FullPath)
      strChr = Mid$(FullPath, strLength, 1)
        
        While strChr <> "\" And intIncrement < strLength
          strConverted = strChr & strConverted
            intIncrement = intIncrement + 1
          strChr = Mid$(FullPath, strLength - intIncrement, 1)
         DoEvents
       Wend
    GetFileName = strConverted
End Function

'Gets the windows version

Public Function GetWindowsVersion()
  Dim OSInfo As OSVERSIONINFO, strVersion As String, lngReturn&
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
      lngReturn = GetVersionEx(OSInfo)
        Select Case OSInfo.dwPlatformId
          Case 1
            strVersion = "9X"
          Case 2
            strVersion = "NT"
        End Select
  GetWindowsVersion = strVersion
End Function
