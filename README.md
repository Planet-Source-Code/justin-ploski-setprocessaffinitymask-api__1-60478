<div align="center">

## SetProcessAffinityMask API


</div>

### Description

This code allows you to set the process affinity on a running thread for multi processor computers.

Get the process ID of a running thread by filename. Get the process handle by the PID. Set your custom Affinity Mask. Apply the affinity mask to the process thread.
 
### More Info
 
Process Name is the only variable input.

You'll need to customize the affinity mask - see comments by the MyMask assignment for information.

This code is Windows API that allows you to specify a running process name and obtain the PID (Process ID). From the PID, it then obtains a handle on the process. You then set a custom affinity BitMask for the process, and pass the handle and affinity mask to the SetProcessAffinity function. Use GetCurrentProcess() API (always returns long -1) in place of the application handle to set the affinity on the current application.

See the comments when setting the MyMask variable to customize which processors will be used.

This is my first submission. I've been leeching off PlanetSourceCode for years, so I figured it's time to give something back. I've seen alot of questions but not many answers related to process affinity for multiprocessors. Please comment if you find this code useful.

The SetProcessAffinity API returns 0 if the affinity was set correctly. Anything &lt;&gt; 0 is an error.

If using a single processor machine, the only valid process affinity is CPU0.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Justin Ploski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/justin-ploski.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/justin-ploski-setprocessaffinitymask-api__1-60478/archive/master.zip)

### API Declarations

```
Option Explicit
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Function SetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwProcessAffinityMask As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WTS_CURRENT_SERVER_HANDLE = 0&amp;
Private Type WTS_PROCESS_INFO
 SessionID As Long
 ProcessID As Long
 pProcessName As Long
 pUserSid As Long
 End Type
```


### Source Code

```
Public Sub Main()
 Call SetAffinityByEXE("notepad.exe")
End Sub
Private Sub SetAffinityByEXE(strImageName As String)
 Const PROCESS_QUERY_INFORMATION = 1024
 Const PROCESS_VM_READ = 16
 Const MAX_PATH = 260
 Const STANDARD_RIGHTS_REQUIRED = &HF0000
 Const SYNCHRONIZE = &H100000
 Const PROCESS_ALL_ACCESS = &H1F0FFF
 Const TH32CS_SNAPPROCESS = &H2&
 Const hNull = 0
 Const WIN95_System_Found = 1
 Const WINNT_System_Found = 2
 Const Default_Log_Size = 10000000
 Const Default_Log_Days = 0
 Const SPECIFIC_RIGHTS_ALL = &HFFFF
 Const STANDARD_RIGHTS_ALL = &H1F0000
 Dim BitMasks() As Long, NumMasks As Long, LoopMasks As Long
 Dim MyMask As Long
 Const AffinityMask As Long = &HF ' 00001111b
 Dim lngPID As Long
 Dim lngHwndProcess
 lngPID = GetProcessID(strImageName)
 If lngPID = 0 Then
 MsgBox "Could not get process ID of " & strImageName, vbCritical, "Error"
 Exit Sub
 End If
 lngHwndProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, lngPID)
 If lngHwndProcess = 0 Then
 MsgBox "Could not obtain a handle for the Process ID: " & lngPID, vbCritical, "Error"
 Exit Sub
 End If
 BitMasks() = GetBitMasks(AffinityMask)
 'Use CPU0
 MyMask = BitMasks(0)
 'Use CPU1
 'MyMask = BitMasks(1)
 'Use CPU0 and CPU1
 'MyMask = BitMasks(0) Or BitMasks(1)
 'The CPUs to use are specified by the array index.
 'To use CPUs 0, 2, and 4, you would use:
 'MyMask = BitMasks(0) Or BitMasks(2) Or BitMasks(4)
 'To Set Affinity, pass the application handle and your custom affinity mask:
 'SetProcessAffinityMask(lngHwndProcess, MyMask)
 'Use GetCurrentProcess() API instead of lngHwndProcess to set affinity on the current app.
 If SetProcessAffinityMask(lngHwndProcess, MyMask) = 1 Then
 MsgBox "Affinity Set", vbInformation, "Success"
 Else
 MsgBox "Failed To Set Affinity", vbCritical, "Failure"
 End If
End Sub
Private Function GetBitMasks(ByVal inValue As Long) As Long()
 Dim RetArr() As Long, NumRet As Long
 Dim LoopBits As Long, BitMask As Long
 Const HighBit As Long = &H80000000
 ReDim RetArr(0 To 31) As Long
 For LoopBits = 0 To 30
 BitMask = 2 ^ LoopBits
 If (inValue And BitMask) Then
 RetArr(NumRet) = BitMask
 NumRet = NumRet + 1
 End If
 Next LoopBits
 If (inValue And HighBit) Then
 RetArr(NumRet) = HighBit
 NumRet = NumRet + 1
 End If
 If (NumRet > 0) Then ' Trim unused array items and return array
 If (NumRet < 32) Then ReDim Preserve RetArr(0 To NumRet - 1) As Long
 GetBitMasks = RetArr
 End If
End Function
Private Function GetProcessID(strProcessName As String) As Long
 Dim RetVal As Long
 Dim Count As Long
 Dim i As Integer
 Dim lpBuffer As Long
 Dim p As Long
 Dim udtProcessInfo As WTS_PROCESS_INFO
 Dim lngProcessID As Long
 Dim strTempProcessName As String
 RetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
 If RetVal Then ' WTSEnumerateProcesses was successful
 p = lpBuffer
 For i = 1 To Count
 ' Count is the number of Structures in the buffer
 ' WTSEnumerateProcesses returns a pointer, so copy it to a
 ' WTS_PROCESS_INO UDT so you can access its members
 CopyMemory udtProcessInfo, ByVal p, LenB(udtProcessInfo)
 ' Add items to the ListView control
 lngProcessID = CLng(udtProcessInfo.ProcessID)
 ' Since pProcessName contains a pointer, call GetStringFromLP to get the
 ' variable length string it points to
 If udtProcessInfo.ProcessID = 0 Then
  'MsgBox "System Idle Process"
 Else
  strTempProcessName = GetStringFromLP(udtProcessInfo.pProcessName)
  If UCase(strTempProcessName) = UCase(strProcessName) Then
  GetProcessID = lngProcessID
  End If
 End If
 p = p + LenB(udtProcessInfo)
 Next i
 WTSFreeMemory lpBuffer 'Free your memory buffer
 Else
 MsgBox "Error", vbCritical, "Fatal Error"
 End If
End Function
Private Function GetStringFromLP(ByVal StrPtr As Long) As String
 Dim b As Byte
 Dim tempStr As String
 Dim bufferStr As String
 Dim Done As Boolean
 Done = False
 Do
 ' Get the byte/character that StrPtr is pointing to.
 CopyMemory b, ByVal StrPtr, 1
 If b = 0 Then ' If you've found a null character, then you're done.
 Done = True
 Else
 tempStr = Chr$(b) ' Get the character for the byte's value
 bufferStr = bufferStr & tempStr 'Add it to the string
 StrPtr = StrPtr + 1 ' Increment the pointer to next byte/char
 End If
 Loop Until Done
 GetStringFromLP = bufferStr
End Function
```

