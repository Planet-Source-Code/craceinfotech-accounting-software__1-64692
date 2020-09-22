Attribute VB_Name = "Backup"
'Public CompanyYear As String
'Public ZipFile As String

'***********************************************
'This Software is developed by craceinfotech.
'Web site : http://www.craceinfotech.com
'email    : craceinfotech.yahoo.com
'date     : 18.03.2006
'***********************************************

    Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&

   Public Function ExecCmd(cmdline$)
      Dim proc As PROCESS_INFORMATION
      Dim start As STARTUPINFO
      Dim RET As Long

      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)

      ' Start the shelled application:
      RET& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

      ' Wait for the shelled application to finish:
         RET& = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, RET&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = RET&
   End Function

'   Sub Form_Click()
'      Dim retval As Long
'      'retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -a d:\natarajan\test\backups\backup.zip @d:\natarajan\test\backups\filelist.txt")
'      retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -e d:\natarajan\test\backups\backup.zip d:\natarajan\test\backups\test")
'      MsgBox "Process Finished, Exit Code " & retval
'   End Sub

'c:\progra~1\WinZip\winzip32 -a d:\natarajan\test\backups\backup.zip @d:\natarajan\test\backups\filelist.txt
'c:\progra~1\WinZip\winzip32 -e [options] filename[.zip] folder
'c:\progra~1\WinZip\winzip32 -e d:\natarajan\test\backups\backup.zip d:\natarajan\test\backups\TEST



'Dim fLen As Integer, filepath As String
'filepath = "C:myfile.txt"
'On Error Resume Next
'fLen = Len(Dir$(filepath))
'If Err Or fLen = 0 Then
''file dosent exist
'Else
''file exists
'End If
