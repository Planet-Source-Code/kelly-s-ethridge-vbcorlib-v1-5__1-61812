VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: Directory
'

''
' Provides a set of static methods for manipulating and retrieving directory information.
'
Option Explicit
Private Const ALT_DIRECTORY_SEPARATOR_CHAR      As Integer = 47
Private Const DIRECTORY_SEPARATOR_CHAR          As Integer = 92
Private Const FILE_FLAG_BACKUP_SEMANTICS        As Long = &H2000000
Private Const ERROR_NO_MORE_FILES               As Long = 18&

Private Enum SystemEntryTypes
    FileEntry = 1
    DirectoryEntry = 2
End Enum



''
' Creates a directory and any subdirectories in the specified path.
'
' @param sPath The directory path to be created.
' @return A DirectoryInfo object that represents the created directory.
' @remarks The path to be created can be a relative or absolute path.
'
Public Function CreateDirectory(ByVal sPath As String) As DirectoryInfo
    sPath = cString.Trim(sPath)
    Call Path.VerifyPath(sPath)
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    sPath = cString.TrimEnd(sPath, Path.DirectorySeparatorString & Path.AltDirectorySeparatorString)
    If Len(sPath) > MAX_PATH Then _
        Throw New PathTooLongException
    
    Dim list As ArrayList
    Set list = New ArrayList
    
    Dim dirsep As Long
    dirsep = Path.DirectorySeparatorChar
    
    Dim altdirsep As Long
    altdirsep = Path.AltDirectorySeparatorChar
    
    Dim volsep As Long
    volsep = Path.VolumeSeparatorChar
    
    Dim i As Long
    If Path.IsPathRooted(sPath) Then
        i = Len(Path.GetPathRoot(sPath)) + 1
    End If
    
    Dim Max As Long
    Max = Len(sPath)
    
    Dim pathBuf As WordBuffer
    Call InitWordBuffer(pathBuf, StrPtr(sPath), Max)
    
    Dim foundsep As Boolean
    Dim p As String
    Do While i < Max
        Select Case pathBuf.Data(i)
            Case dirsep, altdirsep
                If Not foundsep Then
                    p = Left$(sPath, i)
                    If Not Exists(p) Then Call list.Add(p)
                    foundsep = True
                End If
            Case Else
                foundsep = False
        End Select
        i = i + 1
    Loop
    If Not Exists(sPath) Then Call list.Add(sPath)
    
    For i = 0 To list.Count - 1
        If CreateDirectoryA(list(i), ByVal vbNullPtr) = BOOL_FALSE Then IOError Err.LastDllError
    Next i
    
    Set CreateDirectory = Cor.NewDirectoryInfo(sPath)
End Function

''
' Checks if the path exists.
'
' @param sPath The path to check for existence.
' @return Indication of the existence of the path.
' @remarks The path can be relative or absolute.
'
Public Function Exists(ByVal sPath As String) As Boolean
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    
    If file.GetFileData(sPath, Data) = NO_ERROR Then
        Exists = CBool(Data.dwFileAttributes And FileAttributes.DirectoryAttr)
    End If
End Function

''
' Deletes directories and files within directories.
'
' @param sPath The top directory to be deleted.
' @param Recursive If set to True, then all directories and files in the top
' directory will be deleted along with the top directory. Otherwise, the specified
' directory must be empty to be deleted.
'
Public Sub Delete(ByVal sPath As String, Optional ByVal Recursive As Boolean = False)
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath))
    
    sPath = cString.TrimEnd(Path.InternalGetFullPath(sPath), "/\")
    
    If Recursive Then
        Dim FileHandle As Handle
        Dim Data As WIN32_FIND_DATA
        FileHandle = FindFirstFile(sPath & "\*", Data)
        If FileHandle <> INVALID_HANDLE Then
            Do
                Dim FileName As String
                FileName = GetFileNameFromFindData(Data)
                If Len(FileName) > 0 Then
                    Dim FullPath As String
                    FullPath = Path.Combine(sPath, FileName)
                    If Data.dwFileAttributes And FileAttributes.DirectoryAttr Then
                        Call Delete(FullPath, True)
                    Else
                        If DeleteFile(FullPath) = BOOL_FALSE Then IOError Err.LastDllError, FullPath
                    End If
                End If
            Loop While FindNextFile(FileHandle, Data) <> BOOL_FALSE
            
            Dim E As Long
            E = Err.LastDllError    ' get this before FindClose can change it.
            Call FindClose(FileHandle)
            If E <> ERROR_NO_MORE_FILES Then IOError E, FullPath
        End If
    End If
    If RemoveDirectory(sPath) = BOOL_FALSE Then
        E = Err.LastDllError
        If E = ERROR_PATH_NOT_FOUND Then Exit Sub
        IOError E, sPath
    End If
End Sub

''
' Returns the time of creation for the specified directory in local time.
'
' @param sPath The directory to retrieve the creation time of.
' @return A cDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTime(ByVal sPath As String) As cDateTime
    Set GetCreationTime = file.GetCreationTime(sPath)
End Function

''
' Returns the time of creation for the specified directory in UTC time.
'
' @param sPath The directory to retrieve the creation time of.
' @return A cDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTimeUtc(ByVal sPath As String) As cDateTime
    Set GetCreationTimeUtc = file.GetCreationTimeUtc(sPath)
End Function

''
' Returns the time the directory was last accessed in local time.
'
' @param sPath The directory to retrieve the last time it was accessed.
' @return A cDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTime(ByVal sPath As String) As cDateTime
    Set GetLastAccessTime = file.GetLastAccessTime(sPath)
End Function

''
' Returns the time the directory was last accessed in UTC time.
'
' @param sPath The directory to retrieve the last time it was accessed.
' @return A cDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTimeUtc(ByVal sPath As Stream) As cDateTime
    Set GetLastAccessTimeUtc = file.GetLastAccessTimeUtc(sPath)
End Function

''
' Returns the time the directory was last written to in local time.
'
' @param sPath The directory to retrieve the last time it was written to.
' @return A cDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTime(ByVal sPath As String) As cDateTime
    Set GetLastWriteTime = file.GetLastWriteTime(sPath)
End Function

''
' Returns the time the directory was last written to in UTC time.
'
' @param sPath The directory to retrieve the last time it was written to.
' @return A cDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTimeUtc(ByVal sPath As String) As cDateTime
    Set GetLastWriteTimeUtc = file.GetLastWriteTimeUtc(sPath)
End Function

''
' Returns the current directory the application is set to.
'
' @return The current set directory.
'
Public Function GetCurrentDirectory() As String
    GetCurrentDirectory = CurDir
End Function

''
' Returns a list of all the directories found in the specified directory
' that matches the search pattern.
'
' @param sPath The directory to find the requested directories.
' @param SearchPattern A pattern to compare all the directories against, returning
' those that match the pattern.
' @return An array of the matching directories. If no directories match the pattern, then
' an empty zero-length array.
' @remarks The path can be absolute or relative.
'
Public Function GetDirectories(ByVal sPath As String, Optional ByVal SearchPattern As String = "*") As String()
    GetDirectories = InternalGetFiles(sPath, SearchPattern, DirectoryEntry)
End Function

''
' Returns the root directory of the specified path.
'
' @param sPath The path to return the root of.
' @return The root of the specified directory.
' @remarks If the path is relative, then the current directory will
' be used as the root.
'
Public Function GetDirectoryRoot(ByVal sPath As String) As String
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    If Not Path.IsPathRooted(sPath) Then
        sPath = Path.GetFullPath(sPath)
    End If
    GetDirectoryRoot = Path.GetPathRoot(sPath)
End Function

''
' Returns a list of all the files that match a pattern within the specified directory.
'
' @param sPath The directory to search for the matching files.
' @param SearchPattern A pattern to match the files against.
' @return The list of matching file names. If no filenames matched, then an empty,
' zero-length array is returned.
' @remarks The path can be absolute or relative. The path is not included in the
' filename result.
'
Public Function GetFiles(ByVal sPath As String, Optional ByVal SearchPattern As String = "*") As String()
    GetFiles = InternalGetFiles(sPath, SearchPattern, FileEntry)
End Function

''
' Returns both directories and files that match the specified pattern.
'
' @param sPath The directory to search for the files and directories.
' @param SearchPattern The pattern the filenames and directories must match.
' @return The list of files and directories that matched the pattern. If no files or directories
' matched the pattern, then an empty, zero-length array is returned.
' @remarks The path can be absolute or relative.
'
Public Function GetFileSystemEntries(ByVal sPath As String, Optional ByVal SearchPattern As String = "*") As String()
    GetFileSystemEntries = InternalGetFiles(sPath, SearchPattern, DirectoryEntry Or FileEntry)
End Function

''
' Returns a list of the logical drives installed on this machine.
'
' @return The list of logical drives on this machine.
'
Public Function GetLogicalDrives() As String()
    Dim ret() As String
    ReDim ret(25) As String ' just allocate the maximum letters possible.
    
    Dim mask As Long
    mask = VBCorLibTypes.GetLogicalDrives
    
    Dim letter As Long
    letter = CHAR_UPPER_A
    
    Dim i As Long
    Do While mask <> 0
        If mask And 1 Then
            ret(i) = Chr$(letter) & ":\"
            i = i + 1
        End If
        letter = letter + 1
        mask = mask \ 2
    Loop
    
    ReDim Preserve ret(0 To i - 1)
    GetLogicalDrives = ret
End Function

''
' Returns the parent of the specified directory.
'
' @param sPath The directory to retrieve the parent of.
' @return A DirectoryInfo object representing the parent directory.
' @remarks The path can be absolute or relative.
'
Public Function GetParent(ByVal sPath As String) As DirectoryInfo
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")

    Call Path.VerifyPath(sPath)
    sPath = Path.GetFullPath(sPath)
    
    Dim DirName As String
    DirName = Path.GetDirectoryName(sPath)
    If cString.IsNull(DirName) Then Exit Function
    Set GetParent = Cor.NewDirectoryInfo(DirName)
End Function

''
' Moves a directory and all of its contents to another location.
'
' @param SourceDirectory The directory to be moved.
' @param DestinationDirectory The directory to be moved to.
' @remarks The destination directory is the source directory renamed
' after the moving of the source directory.
'
Public Sub Move(ByVal SourceDirectory As String, ByVal DestinationDirectory As String)
    If Len(SourceDirectory) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "SourceDirectory")
    If Len(DestinationDirectory) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "DestinationDirectory")
    
    SourceDirectory = LCase$(Path.InternalGetFullRootedPath(SourceDirectory))
    DestinationDirectory = LCase$(Path.InternalGetFullRootedPath(DestinationDirectory))
    
    If SourceDirectory = DestinationDirectory Then _
        Throw Cor.NewIOException("Source and Destination directories must be different.")
    If Path.GetPathRoot(SourceDirectory) <> Path.GetPathRoot(DestinationDirectory) Then _
        Throw Cor.NewIOException("Source and Destination directories must be on the same root drive.")
        
    If MoveFile(SourceDirectory, DestinationDirectory) = BOOL_FALSE Then
        Dim E As Long
        E = Err.LastDllError
        If E = ERROR_FILE_NOT_FOUND Then
            IOError ERROR_PATH_NOT_FOUND, SourceDirectory
        Else
            IOError E
        End If
    End If
End Sub

''
' Changes the current directory to the specified directory.
'
' @param sPath The new directory to set as the current directory.
' @remarks The path can be absolute or relative.
'
Public Sub SetCurrentDirectory(ByVal sPath As String)
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    Call Path.VerifyPath(sPath)
    If Not Exists(sPath) Then _
        Throw New DirectoryNotFoundException
    
    'ChDir sPath
    If SetCurrentDirectoryA(sPath) = BOOL_FALSE Then IOError Err.LastDllError, sPath
End Sub

''
' Sets the creation time of the specified directory to the local time.
'
' @param sPath The directory to set the creation time for.
' @param CreationTime A Date or cDateTime object of the new time.
' @remarks The time will have the UTC offset removed before setting the
' the directory to it. When the time is read, it will be in local time to the timezone.
'
Public Sub SetCreationTime(ByVal sPath As String, ByVal CreationTime As Variant)
    Call InternalSetFileTime(sPath, cDateTime.GetcDateTime(CreationTime).ToFileTime)
End Sub

''
' Sets the creation time of the specified directory to UTC time.
'
' @param sPath The directory to set the creation time for in UTC.
' @param CreationTime A Date or cDateTime object of the new UTC time.
' @remarks The directory is set to the actual time passed in.
'
Public Sub SetCreationTimeUtc(ByVal sPath As String, ByVal CreationTime As Variant)
    Call InternalSetFileTime(sPath, cDateTime.GetcDateTime(CreationTime).ToFileTimeUtc)
End Sub

''
' Sets the time the directory was last accessed in local time.
'
' @param sPath The directory to set the new time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
' @remarks The UTC offset is removed from the time before setting the directory.
'
Public Sub SetLastAccessTime(ByVal sPath As String, ByVal LastAccessTime As Variant)
    Call InternalSetFileTime(sPath, , cDateTime.GetcDateTime(LastAccessTime).ToFileTime)
End Sub

''
' Sets the time the directory was last accessed in UTC time.
'
' @param sPath The directory to set the new time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
' @remarks The directory is set to the actual time passed in.
'
Public Sub SetLastAccessTimeUtc(ByVal sPath As String, ByVal LastAccessTime As Variant)
    Call InternalSetFileTime(sPath, , cDateTime.GetcDateTime(LastAccessTime).ToFileTimeUtc)
End Sub

''
' Sets the time the directory was last written to in local time.
'
' @param sPath The directory to set the new time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTime(ByVal sPath As String, ByVal LastWriteTime As Variant)
    Call InternalSetFileTime(sPath, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTime)
End Sub

''
' Sets the time the directory was last written to in UTC time.
'
' @param sPath The directory to set the new time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTimeUtc(ByVal sPath As String, ByVal LastWriteTime As Variant)
    Call InternalSetFileTime(sPath, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTimeUtc)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifySearchPattern(ByRef s As String)
    Dim i As Long
    i = InStr(s, "..")
    If i > 0 Then
        If i = Len(s) - 1 Then Throw Cor.NewArgumentException("Cannot end search pattern with '..'")
        Select Case Asc(Mid$(s, i + 2, 1))
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR
                Throw Cor.NewArgumentException("'..' cannot be followed immediately by a directory separator.")
        End Select
    End If
End Sub

Private Function InternalGetFiles(ByRef sPath As String, ByRef SearchPattern As String, ByVal EntryTypes As SystemEntryTypes) As String()
    Dim FullPath As String
    FullPath = Path.InternalGetFullPath(sPath)
    
    Call VerifySearchPattern(SearchPattern)
    
    Dim Data        As WIN32_FIND_DATA
    Dim FileHandle  As Long
    FileHandle = FindFirstFile(Path.Combine(sPath, SearchPattern), Data)
    
    Dim list As ArrayList
    Set list = New ArrayList
    
    Dim include As Boolean
    If FileHandle <> INVALID_HANDLE Then
        Do
            If Data.dwFileAttributes And FileAttributes.DirectoryAttr Then
                include = (EntryTypes And SystemEntryTypes.DirectoryEntry)
            Else
                include = (EntryTypes And SystemEntryTypes.FileEntry)
            End If
            
            If include Then
                Dim s As String
                s = GetFileNameFromFindData(Data)
                If Len(s) > 0 Then Call list.Add(Path.Combine(sPath, s))
            End If
        Loop While FindNextFile(FileHandle, Data) <> BOOL_FALSE
    End If
    
    Dim E As Long
    E = Err.LastDllError    ' cache the error before FindClose changes it.
    
    Call FindClose(FileHandle)
    
    Select Case E
        Case 0, ERROR_NO_MORE_FILES, ERROR_FILE_NOT_FOUND   ' ignore these errors.
        Case Else: IOError E, sPath
    End Select

    Dim ret() As String
    If list.Count > 0 Then
        ret = list.ToArray(ciString)
    Else
        ret = cArray.CreateInstance(ciString)
    End If
    
    InternalGetFiles = ret
End Function

Private Sub InternalSetFileTime(ByRef sPath As String, Optional ByVal CreationTime As Variant, Optional ByVal LastAccessTime As Variant, Optional ByVal LastWriteTime As Variant)
    Dim cnt As Long
    Dim lat As Long
    Dim lwt As Long
    
    If Not IsMissing(CreationTime) Then cnt = VarPtr(CreationTime) + VARIANTDATA_OFFSET
    If Not IsMissing(LastAccessTime) Then lat = VarPtr(LastAccessTime) + VARIANTDATA_OFFSET
    If Not IsMissing(LastWriteTime) Then lwt = VarPtr(LastWriteTime) + VARIANTDATA_OFFSET
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    Call Path.VerifyPath(sPath)
    
    Dim DirHandle As Long
    DirHandle = CreateFile(sPath, FileAccess.WriteAccess, FileShare.None, ByVal 0&, FileMode.OpenExisting, FILE_FLAG_BACKUP_SEMANTICS, 0)
    If DirHandle = INVALID_HANDLE Then IOError Err.LastDllError, sPath
    If SetFileTime(DirHandle, ByVal cnt, ByVal lat, ByVal lwt) = BOOL_FALSE Then
        Close CloseHandle(DirHandle)
        IOError Err.LastDllError, sPath
    End If
    Call CloseHandle(DirHandle)
End Sub

Private Function GetFileNameFromFindData(ByRef Data As WIN32_FIND_DATA) As String
    Dim ret As String
    
    Select Case Data.cFileName(0)
        Case 46:    Exit Function   ' skip periods
        Case 0:     ret = StrConv(Data.cAlternate, vbUnicode)
        Case Else:  ret = StrConv(Data.cFileName, vbUnicode)
    End Select
    GetFileNameFromFindData = Left$(ret, InStr(ret, vbNullChar) - 1)
End Function

