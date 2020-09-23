VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: Environment
'

''
' The Environment class provides functions for retrieving information about the local machine
'
' @remarks This class cannot be directly instantiated. In order to access the methods, use
' the class name directly.
' <pre>
' Debug.Print Environment.MachineName
' Debug.Print Environment.UserName
' </pre>
'
Option Explicit

Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameW" (ByVal lpSystemName As Long, ByVal lpAccountName As Long, ByVal Sid As Long, ByRef cbSid As Long, ByVal ReferencedDomainName As Long, ByRef cbReferencedDomainName As Long, ByRef peUse As Long) As Long
Private Declare Function GetProcessWindowStation Lib "user32.dll" () As Long
Private Declare Function GetUserObjectInformation Lib "user32.dll" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, ByRef pvInfo As Any, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariableAPI Lib "kernel32.dll" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function GetEnvironmentVariableAPI Lib "kernel32.dll" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As Long

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

Private Const ERR_PATHNOTFOUND              As Long = 76
Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31
Private Const MAX_PATH                      As Long = 260
Private Const UNLEN                         As Long = 256
Private Const UOI_FLAGS                     As Long = 1
Private Const WSF_VISIBLE                   As Long = 1
Private Const CHAR_EQUAL                    As Long = 61
Private Const CHAR_QUOTE                    As Long = 34
Private Const CHAR_SPACE                    As Long = 32
Private Const WM_SETTINGCHANGE              As Long = &H1A
Private Const HWND_BROADCAST                As Long = &HFFFF&
Private Const SMTO_NORMAL                   As Long = &H0
Private Const ERROR_ENVVAR_NOT_FOUND        As Long = 203&


''
' This enum represents a list of folders accessible using <i>GetFolderPath</i>.
' These definitions are from the Microsoft Dot NET 1.1 MSDN Library.
'
' @param ApplicationData The directory that serves as a common repository for application-specific data for the current roaming user.
' A roaming user works on more than one computer on a network. A roaming user's profile is kept on a server on the network and is loaded onto a system when the user logs on.
' @param CommonApplicationData The directory that serves as a common repository for application-specific data that is used by all users.
' @param CommonProgramFiles The directory for components that are shared across applications.
' @param Cookies The directory that serves as a common repository for Internet cookies.
' @param Desktop The logical Desktop rather than the physical file system location.
' @param DesktopDirectory The directory used to physically store file objects on the desktop.
' Do not confuse this directory with the desktop folder itself, which is a virtual folder.
' @param Favorites The directory that serves as a common repository for the user's favorite items.
' @param History The directory that serves as a common repository for Internet history items.
' @param InternetCache The directory that serves as a common repository for temporary Internet files.
' @param LocalApplicationData The directory that serves as a common repository for application-specific data that is used by the current, non-roaming user.
' @param MyComputer The "My Computer" folder.
' @param MyMusic The "My Music" folder.
' @param MyPictures The "My Pictures" folder.
' @param Personal The directory that serves as a common repository for documents (My Documents).
' @param ProgramFiles The program files folder.
' @param Programs The directory that contains the user's program groups.
' @param Recent The directory that contains the user's most recently used documents.
' @param SendTo The directory that contains the Send To menu items.
' @param StartMenu The directory that contains the Start menu items.
' @param Startup The directory that corresponds to the user's Startup program group.
' The system starts these programs whenever a user logs on or starts Windows NT or later, or starts Windows 98.
' @param SystemDir The system directory.
' @param Templates The directory that serves as a common repository for document templates.
' @param Windows The windows directory.
'
Public Enum SpecialFolder
   ApplicationData = &H1A
   CommonApplicationData = &H23
   CommonProgramFiles = &H2B
   Cookies = &H21
   Desktop = &H0
   DesktopDirectory = &H10
   Favorites = &H6
   History = &H22
   InternetCache = &H20
   LocalApplicationData = &H1C
   MyComputer = &H11
   MyMusic = &HD
   MyPictures = &H27
   Personal = &H5
   ProgramFiles = &H26
   Programs = &H2
   Recent = &H8
   SendTo = &H9
   StartMenu = &HB
   Startup = &H7
   SystemDir = &H25
   Templates = &H15
   Windows = &H24
End Enum

''
' Used to indicate the source or destination of the
' environment variables when read from or written to the system.
'
' @param Machine The source of the variable is machine wide. These
' are variables that are defined for the system and all users.
' @param Process The source of the variable is for the current process.
' This combines both machine and user level variables, giving priority
' to the user level variables.
' @param User The source of the variable is of the current user.
' @remarks <b>Machine</b> and <b>User</b> are not supported on Windows 95, 98 or ME.
'
Public Enum EnvironmentVariableTarget
    Machine = 0
    Process = 1
    User = 2
End Enum


Private mWordBuffer As WordBuffer
Private mOSVersion  As New OperatingSystem



''
' Returns the command specified when executing the current application,
' including any arguments typed after the application name.
'
' @return The command line typed in at a command prompt.
'
Public Property Get CommandLine() As String
    CommandLine = SysAllocString(GetCommandLine)
End Property

''
' Returns the current directy of the executing application.
'
' @return The current directory.
'
Public Property Get CurrentDirectory() As String
    CurrentDirectory = CurDir
End Property

''
' Sets the current directory for the executing application.
'
' @param RHS The new directory to change to.
'
Public Property Let CurrentDirectory(ByVal RHS As String)
    If Len(RHS) = 0 Then _
        Throw Cor.NewArgumentException("Path cannot be an empty string.", "CurrentDirectory")
    
    On Error GoTo errTrap
    Call ChDir(RHS)
    Exit Property
    
errTrap:
    If Err.Number = ERR_PATHNOTFOUND Then
        Throw Cor.NewFileNotFoundException("The specified path was not found.", RHS)
    Else
        Throw Cor.NewArgumentException(Err.Description, "CurrentDirectory")
    End If
End Property

''
' Returns the name of the local computer.
'
' @return The name of the local computer.
'
Public Property Get MachineName() As String
    Dim Size As Long
    Size = MAX_COMPUTERNAME_LENGTH + 1
    
    Dim ret As String
    ret = String$(Size, 0)
    Call GetComputerName(ret, Size)
    MachineName = Left$(ret, Size)
End Property

''
' Returns string of a newline character for the local computer.
'
' @return The newline characters.
' @remarks Being this always going to be in VB, the return is just vbCrLf.
'
Public Property Get NewLine() As String
    NewLine = vbCrLf
End Property

''
' Returns an object representing the current operating system.
'
' @return The OperatingSystem object that represents the current OS.
'
Public Property Get OSVersion() As OperatingSystem
    Set OSVersion = mOSVersion
End Property

''
' Returns the System directory for the current OS.
'
' @return The system directory.
' @remarks Generally this will be C:\Windows\System32 or C:\Winnt\System32 for NT4 system.
'
Public Property Get SystemDirectory() As String
    Dim ret As String
    ret = String$(MAX_PATH + 1, 0)
    
    Dim Size As Long
    Size = GetSystemDirectory(ret, MAX_PATH + 1)
    If Size > 0 Then SystemDirectory = Left$(ret, Size)
End Property

''
' Returns the number of milliseconds the system has been running since it started.
'
' @return Number of milliseconds since the system booted up.
' @remarks Once the maximum value is reached, it will wrap around to be negative,
' at which point 0 is returned. Wrapping will take place in approximatively 24.85 days.
'
Public Property Get TickCount() As Long
    TickCount = timeGetTime
    If TickCount < 0 Then TickCount = 0
End Property

''
' Returns the name of the local machine within a domain
'
' @return The computer name within a domain.
' @remarks Typically the domain name is the name of the machine, but
' can change depending on the application being deployed.<br><br>
' <b>UserDomainName</b> is not supported on Windows 95, 98, ME platforms
' and will throw a <b>PlatformNotSupportedException</b> exception.
'
Public Property Get UserDomainName() As String
    VerifyNTMachine
    
    Dim Size As Long
    Size = 128
    
    Dim Sid As String
    Sid = String$(Size, 0)
    
    Dim buf As String
    Do
        buf = String$(Size, 0)
        If LookupAccountName(0, StrPtr(UserName), StrPtr(Sid), Len(Sid), StrPtr(buf), Size, 0&) = BOOL_FALSE Then IOError Err.LastDllError
    Loop While Size > Len(buf)
    
    UserDomainName = SysAllocString(StrPtr(buf))
End Property

''
' Determines if the current application is allowed to have user interaction.
'
' @return Indication if user interaction is allowed.
' @remarks This is used primarily for when the application is a service. It can
' be used to supress message dialogs and user entry when not allowed.
'
Public Property Get UserInteractive() As Boolean
    Dim hObj As Handle
    Dim Info As UserObjectFlags
    
    If OSVersion.Platform = Win32NT Then
        hObj = GetProcessWindowStation
        If GetUserObjectInformation(hObj, UOI_FLAGS, Info, LenB(Info), 0&) Then
            UserInteractive = (Info.dwFlags And WSF_VISIBLE)
        End If
    Else
        UserInteractive = True
    End If
End Property

''
' Returns the name of the currently logged in user of the local computer.
'
' @return The currently logged in user's name.
'
Public Property Get UserName() As String
    Dim Size As Long
    Size = UNLEN + 1
    
    Dim ret As String
    ret = String$(Size, 0)
    Call GetUserName(ret, Size)
    UserName = Left$(ret, Size)
End Property

''
' Returns the amount of physical memory the current application is allocated.
'
' @return The amount of memory allocated to the application.
' @remarks This method only works for NT installations. 95, 98, and ME will
' receive a 0 when this method is called.
'
Public Property Get WorkingSet() As Long
    If OSVersion.Platform <> Win32NT Then Exit Property
    
    Dim h As Handle
    h = GetCurrentProcess
    
    Dim Info As PROCESS_MEMORY_COUNTERS
    If GetProcessMemoryInfo(h, Info, LenB(Info)) Then
        WorkingSet = Info.WorkingSetSize
    End If
End Property

''
' Replaces environment tags within a string with the actual values.
'
' @param Name A string containing environment variable tags to be replaced.
' @return A string containing actual environment variable values in place of tags.
' @remarks The environment variables are tagged using %.
' A string such as 'My OS is %OS%.' would become 'My OS is Windows_NT.' for NT machines.
'
Public Function ExpandEnvironmentVariables(ByVal Name As String) As String
    Const DEF_SIZE As Long = 128
    
    If Len(Name) = 0 Then Exit Function
    
    Dim buf As String
    buf = SysAllocStringLen(0, DEF_SIZE)
    
    Dim Size As Long
    Size = ExpandEnvironmentStrings(Name, buf, DEF_SIZE)
    
    If Size > DEF_SIZE Then
        buf = SysAllocStringLen(0, Size)
        Size = ExpandEnvironmentStrings(Name, buf, Size)
    End If
    
    ExpandEnvironmentVariables = Left$(buf, Size - 1)
End Function

''
' Returns an array of parsed Arguments from the command line.
'
' @return An array of command line Arguments, including the application name.
' @remarks The application is the first Arguments in the array. Arguments are
' separated by spaces. If spaces are embedded in an Arguments and should not
' be used to separate the Arguments, then the Arguments must be surrounded in quotes.
'
Public Function GetCommandLineArgs() As String()
    Dim Arguments As ArrayList
    Set Arguments = New ArrayList
    
    Dim Line As String
    Line = SysAllocString(GetCommandLine)
    
    Dim Max As Long
    Max = Len(Line)
    
    Dim buf As WordBuffer
    Call InitWordBuffer(buf, StrPtr(Line), Max)
    
    Dim prevPos     As Long
    Dim pos         As Long
    Dim inQ         As Boolean
    Do While pos < Max
        Select Case buf.Data(pos)
            Case CHAR_QUOTE
                If inQ Then
                    Call Arguments.Add(Mid$(Line, prevPos + 1, pos - prevPos))
                    inQ = False
                Else
                    inQ = True
                End If
                prevPos = pos + 1
            Case CHAR_SPACE
                If Not inQ Then
                    If buf.Data(pos - 1) <> CHAR_SPACE Then
                        If pos - prevPos > 0 Then Call Arguments.Add(Mid$(Line, prevPos + 1, pos - prevPos))
                    End If
                    Do While buf.Data(pos) = CHAR_SPACE
                        pos = pos + 1
                        If pos = Max Then Exit Do
                    Loop
                    prevPos = pos
                    pos = pos - 1
                End If
        End Select
        pos = pos + 1
    Loop
    
    If buf.Data(pos - 1) <> CHAR_SPACE And buf.Data(pos - 1) <> CHAR_QUOTE Then
        Call Arguments.Add(Mid$(Line, prevPos + 1, pos - prevPos))
    End If
    
    GetCommandLineArgs = Arguments.ToArray(ciString)
End Function

''
' Returns an environment variable value.
'
' @param Name The name of the environment variable to return.
' @param Target The system region to retrieve the variable from.
' @return The value of the environment variable.
' @remarks If the variable does not exist, then an empty string is returned.
' <p>A Target of Process will retrieve the variable from the current process. Other
' targets will retrieve the variable value from the registry.</p>
'
Public Function GetEnvironmentVariable(ByVal Name As String, Optional ByVal Target As EnvironmentVariableTarget = EnvironmentVariableTarget.Process) As String
    Dim rk As RegistryKey
    
    Select Case Target
        Case EnvironmentVariableTarget.Process
            ' use the GetEnvironmentVariable api to allow for realtime
            ' updating of the environment variables for the current
            ' process. VB's Environ does not reflect variable changes.
            Dim Size As Long
            Size = 128
            Do
                Dim buf As String
                buf = String$(Size, 0)
                Size = GetEnvironmentVariableAPI(Name, buf, Len(buf))
                If Size = 0 Then
                    Dim E As Long
                    E = Err.LastDllError
                    If E = ERROR_ENVVAR_NOT_FOUND Or E = 0 Then Exit Function
                    IOError E, "Name"
                End If
            Loop While Size > Len(buf)
            GetEnvironmentVariable = Left$(buf, Size)
            
        Case EnvironmentVariableTarget.Machine
            Call VerifyNTMachine
            Set rk = Registry.LocalMachine.OpenSubKey("system\currentcontrolset\control\session manager\environment")
            GetEnvironmentVariable = rk.GetValue(Name, vbNullString)
            
        Case EnvironmentVariableTarget.User
            Call VerifyNTMachine
            Set rk = Registry.CurrentUser.OpenSubKey("environment")
            GetEnvironmentVariable = rk.GetValue(Name, vbNullString)
            
        Case Else
            Throw Cor.NewArgumentException("Invalid Target value.", "Target")
    End Select
End Function

''
' Returns an IDictionary object containing all environment variables as name/value pairs.
'
' @param Target The system region to retrieve the environment variable from.
' @return A dictionary containing the name/value pairs.
' @remarks The hidden environment variables are not included in the dictionary.
' <p>A Target of Process will retrieve the variable from the current process. Other
' targets will retrieve the variable value from the registry.</p>
'
Public Function GetEnvironmentVariables(Optional ByVal Target As EnvironmentVariableTarget = Process) As IDictionary
    Dim ret As New Hashtable
    Dim rk As RegistryKey
    
    If Target = EnvironmentVariableTarget.Process Then
        ' use the GetEnvironmentStrings api to allow for realtime
        ' updating of the environment variables for the current
        ' process. VB's Environ does not reflect variable changes.
        Set ret = InternalGetEnvironmentVariables
    Else
        Call VerifyNTMachine
        Select Case Target
            Case EnvironmentVariableTarget.Machine: Set rk = Registry.LocalMachine.OpenSubKey("system\currentcontrolset\control\session manager\environment")
            Case EnvironmentVariableTarget.User:    Set rk = Registry.CurrentUser.OpenSubKey("environment")
            Case Else: Throw Cor.NewArgumentException("Invalid Target value.", "Target")
        End Select
        
        Dim names() As String
        names = rk.GetValueNames
        
        Dim i As Long
        For i = 0 To UBound(names)
            Call ret.Add(names(i), rk.GetValue(names(i)))
        Next i
    End If
    
    Set GetEnvironmentVariables = ret
End Function

''
' Returns the path for a specified folder type.
'
' @param Folder The folder to return the path of.
' @return The path for the specified folder.
'
Public Function GetFolderPath(ByVal Folder As SpecialFolder) As String
    Dim buf As String
    
    buf = String$(MAX_PATH, 0)
    If SHGetFolderPath(0, Folder, 0, 0, buf) Then Exit Function
    GetFolderPath = SysAllocString(StrPtr(buf))
End Function

''
' Returns the logical drives installed on the local machine.
'
' @return An array containing all of the logical drives installed.
'
Public Function GetLogicalDrives() As String()
    GetLogicalDrives = Directory.GetLogicalDrives
End Function

''
' Sets, creates, or deletes an environment variable in the specified region.
'
' @param Variable The name of the variable to set the value for.
' @param Value The value of the variable.
' @param Target The region the variable is located in the system.
' @remarks Setting the variable in the Process region only affects the current
' process that is setting the variable. Setting a variable in a User or Machine
' region will set the values in the registry.
' <p>by setting the value to an empty string, the variable will be deleted.</p>
'
Public Sub SetEnvironmentVariable(ByVal Variable As String, ByVal Value As String, Optional ByVal Target As EnvironmentVariableTarget = Process)
    If Len(Variable) = 0 Then _
        Throw Cor.NewArgumentException("Variable name cannot be an empty string.", "Variable")
    If InStr(Variable, "=") > 0 Then _
        Throw Cor.NewArgumentException("Variable name cannot contain an equal '=' sign.", "Variable")
    If Len(Variable) > 255 Then _
        Throw Cor.NewArgumentException("Variable name cannot exceed 255 characters.", "Variable")
    If Len(Value) > 32767 Then _
        Throw Cor.NewArgumentException("Value cannot exceed a length of 32767 characters including the terminating null.", "Value")
    
    Select Case Target
        Case EnvironmentVariableTarget.Process
            If SetEnvironmentVariableAPI(Variable, Value) = BOOL_FALSE Then _
                Throw Cor.NewArgumentException(GetErrorMessage(Err.LastDllError), "Value")
                
        Case EnvironmentVariableTarget.Machine, EnvironmentVariableTarget.User
            Call VerifyNTMachine
            
            Dim rk As RegistryKey
            If Target = EnvironmentVariableTarget.Machine Then
                Set rk = Registry.LocalMachine.OpenSubKey("system\currentcontrolset\control\session manager\environment", True)
            Else
                Set rk = Registry.CurrentUser.OpenSubKey("environment", True)
            End If
            
            If Len(Value) > 0 Then
                Call rk.SetValue(Variable, Value)
            Else
                Call rk.DeleteValue(Variable, False)
            End If
            
            Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment", SMTO_NORMAL, 1000, 0)
            
        Case Else
            Throw Cor.NewArgumentException("Invalid target specified.", "Target")
    End Select
End Sub

''
' Returns the number of processors on the current machine.
'
' @return The processor count.
' @remarks This method counts the number of process subkey entries
' located in the 'Hardware\Description\System\CentralProcessor'
' region of the HKEY_LOCAL_MACHINE registry key.
'
Public Property Get ProcessorCount() As Long
    Dim rk As RegistryKey
    Set rk = Registry.LocalMachine.OpenSubKey("Hardware\Description\System\CentralProcessor")
    ProcessorCount = rk.SubKeyCount
End Property



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetResourceString(ByVal resId As Long, ParamArray args() As Variant) As String
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16))
    GetResourceString = cString.FormatArray(LoadResString(resId), vArgs)
End Function

Friend Sub VerifyNTMachine()
    If mOSVersion.Platform <> Win32NT Then Throw Cor.NewNotSupportedException("This operation is only supported on Windows NT platforms.")
End Sub

Friend Function IsNT() As Boolean
    IsNT = mOSVersion.Platform = Win32NT
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' we need to use the GetEnvironmentStrings API method instead of
' iterating through the variables using Environ because Environ
' does not reflect changes to the variables after the application starts.
Private Function InternalGetEnvironmentVariables() As IDictionary
    Dim SA As SafeArray1d
    With SA
        .cbElements = 1
        .cDims = 1
        .pvData = GetEnvironmentStrings
        .cElements = &H7FFFFFFF
    End With
    
    Dim Chars() As Byte
    SAPtr(Chars) = VarPtr(SA)
    
    Dim ret As Hashtable
    Set ret = New Hashtable
    
    Dim isHidden    As Boolean
    Dim pos         As Long
    Dim startPos    As Long
    Dim Name        As String
    Do While Chars(pos) <> 0
        startPos = pos
        Do
            If Chars(pos) = CHAR_EQUAL Then
                If startPos < pos Then Exit Do
                isHidden = True
            End If
            pos = pos + 1
        Loop
        
        If Not isHidden Then Name = BytesToString(Chars, startPos, pos - startPos)
        
        startPos = pos + 1
        Do While Chars(pos) <> 0: pos = pos + 1: Loop   ' find the end of the string.
        
        If Not isHidden Then Call ret.Add(Name, BytesToString(Chars, startPos, pos - startPos))
        
        pos = pos + 1
        isHidden = False
    Loop
    
    FreeEnvironmentStrings SA.pvData
    SAPtr(Chars) = 0
    Set InternalGetEnvironmentVariables = ret
End Function

Private Function BytesToString(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Size As Long) As String
    Dim i As Long
    
    If Size = 0 Then Exit Function
    
    BytesToString = SysAllocStringLen(0, Size)
    mWordBuffer.SA.pvData = StrPtr(BytesToString)
    For i = 0 To Size - 1
        mWordBuffer.Data(i) = Bytes(Index + i)
    Next i
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call InitWordBuffer(mWordBuffer, 0, &H7FFFFFFF)
End Sub
