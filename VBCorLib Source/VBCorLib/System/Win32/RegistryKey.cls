VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKey"
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
'    Module: RegistryKey
'

''
' Represents a key level node inside the Windows Registry.
'
' @remarks This class allows for easy access and manipulation of keys and key values
' within the Windows Registry. By using a key naming convention similar to folders the
' keys and values can be traversed and modified.
' <p>There a set of Root nodes that separate the Registry to common areas for specific
' needs. The Root nodes are the starting place for any keys in the registry. A key name
' could look like 'HKEY_LOCAL_MACHINE\Software'.
' <pre>
' Iterate the values found in a registry key.
'
' Private Sub Main()
'     Dim rk As RegistryKey
'     Dim names() As String
'     Dim i As Long
'
'     '' Open a registry subkey for enumeration.
'     Set rk = Registry.LocalMachine.OpenSubKey("software\microsoft\windows\currentversion\run")
'
'     '' Retrieve all names for the values in the key.
'     names = rk.GetValueNames
'
'     '' enumerate the names and get the value for each,
'     '' displaying the pair as [name] = [value].
'     For i = 0 To UBound(names)
'         Console.WriteLine "{0} = {1}", names(i), rk.GetValue(names(i))
'     Next i
'
'     '' Close the registry key.
'     rk.CloseKey
'
'     '' Wait for user to hit return.
'     Console.ReadLine
' End Sub
'
' ''This code produces the following output.
' ''The list will vary from machine to machine.
'
' ''CplBCL50 = C:\Program Files\EzButton\CplBCL50.EXE
' ''ccApp = "C:\Program Files\Common Files\Symantec Shared\ccApp.exe"
' ''Symantec NetDriver Monitor = C:\PROGRA~1\SYMNET~1\SNDMon.exe
' </pre>
'
' @see Registry
' @see RegistryKeyStatic
'
Option Explicit
Implements IObject
Implements IEnumerable


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal HKey As Long, ByVal lpClass As String, ByRef lpcbClass As Any, ByRef lpReserved As Any, ByRef lpcSubKeys As Any, ByRef lpcbMaxSubKeyLen As Any, ByRef lpcbMaxClassLen As Any, ByRef lpcValues As Any, ByRef lpcbMaxValueNameLen As Any, ByRef lpcbMaxValueLen As Any, ByRef lpcbSecurityDescriptor As Any, ByRef lpftLastWriteTime As Any) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As Any, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByRef lpftLastWriteTime As Currency) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal HKey As Long) As Long

Private Const REG_OPTION_NON_VOLATILE       As Long = 0

'Private Const REG_NONE                      As Long = 0
'Private Const REG_UNKNOWN                   As Long = 0
'Private Const REG_SZ                        As Long = 1
'Private Const REG_DWORD                     As Long = 4
'Private Const REG_BINARY                    As Long = 3
'Private Const REG_MULTI_SZ                  As Long = 7
'Private Const REG_EXPAND_SZ                 As Long = 2
'Private Const REG_QWORD                     As Long = 11
'
'Private Const ERROR_SUCCESS                 As Long = 0
'Private Const ERROR_FILE_NOT_FOUND          As Long = 2
'Private Const ERROR_ACCESS_DENIED           As Long = 5
'Private Const ERROR_INVALID_HANDLE          As Long = 6
'Private Const ERROR_INVALID_PARAMETER       As Long = 87
'Private Const ERROR_CALL_NOT_IMPLEMENTED    As Long = 120
'Private Const ERROR_INSUFFICIENT_BUFFER     As Long = 122
'Private Const ERROR_MORE_DATA               As Long = 234
'Private Const ERROR_NO_MORE_ITEMS           As Long = 259
'Private Const ERROR_CANTOPEN                As Long = 1011
'Private Const ERROR_CANTREAD                As Long = 1012
'Private Const ERROR_CANTWRITE               As Long = 1013
'Private Const ERROR_REGISTRY_RECOVERED      As Long = 1014
'Private Const ERROR_REGISTRY_CORRUPT        As Long = 1015
'Private Const ERROR_REGISTRY_IO_FAILED      As Long = 1016
'Private Const ERROR_NOT_REGISTRY_FILE       As Long = 1017
'Private Const ERROR_KEY_DELETED             As Long = 1018

Private Const MAX_KEY_LENGTH                As Long = 255

Private Const DEFAULT_VALUE_NAME            As String = ""


''
' The value types that can be set and retrieved from the Registry.
'
' @param UnknownKind Determines the value type based on the variable type or registry setting.
' @param StringKind Attempts to set or retrieve the value as a string.
' @param DWordKind Attempts to set or retrieve the value as a DWord (Long).
' @param BinaryKind Attempts to set or retrieve the value as a Byte array.
' @param MultiStringKind Attempts to set or retrieve the value as a String array.
' @param ExpandStringKind Attempts to set or retrieve the value as an expandable string.
' @param QWordKind Attempts to set or retrieve the value as a 64-bit value.
'
Public Enum RegistryValueKind
    UnknownKind = 0
    StringKind = 1
    DWordKind = 4
    BinaryKind = 3
    MultiStringKind = 7
    ExpandStringKind = 2
    QWordKind = 11
End Enum


Private mHKey       As Long
Private mRootKey    As Long
Private mName       As String
Private mRootName   As String
Private mWritable   As Boolean



''
' Returns the name of the key.
'
' @return The name of the key.
'
Public Property Get Name() As String
    Call VerifyKey
    
    If Len(mName) = 0 Then
        Name = mRootName
    Else
        Name = mRootName & "\" & mName
    End If
End Property

''
' Returns a <b>RegistryKey</b> of the requested SubKey with the write permission specified.
'
' @param Name The name of the SubKey to be opened. Any slashes will will be trimmed from both ends.
' @param Writable The permission to modify the opened SubKey.
' @return A new <b>RegistryKey</b> of the opened SubKey.
' @remarks The requested SubKey is opened starting at the location of the current <b>RegistryKey</b>
' node within the Registry.
' <p>If the requested SubKey was not found, then <b>Nothing</b> is returned. No error is thrown.</p>
'
Public Function OpenSubKey(ByVal Name As String, Optional ByVal Writable As Boolean = False) As RegistryKey
    Call VerifyKey
    
    If Len(Name) > MAX_KEY_LENGTH Then _
        Throw Cor.NewArgumentException("Key name cannot exceed 255 characters.", "Name")
    
    Set OpenSubKey = Cor.NewRegistryKey(mHKey, Me.Name, 0, Name, Writable)
    
    On Error GoTo errNoSubKey
    Call OpenSubKey.VerifyKey
    Exit Function
    
errNoSubKey:
    Set OpenSubKey = Nothing
End Function

''
' Closes the <b>RegistryKey</b> object.
'
' @remarks When the key is closed, then it can no longer be manipulated using the <b>RegistryKey</b>
' <p>If the <b>RegistryKey</b> is a system root key, then it will not be closed.</p>
'
Public Sub CloseKey()
    If (Not IsSystemKey) And (mRootKey <> 0) Then
        Call RegCloseKey(mHKey)
        mRootKey = 0
        mHKey = 0
    End If
End Sub

''
' Flushes any changes to the Windows Registry back to disk.
'
' @remarks Generally this does not need to be called. Windows writes back to the disk Registry
' at regular intervals and when the system shuts down. values of other keys written to the
' Registry may also be flushed back to disk as well.
' <p>Only call this function if it is absolutely necessary to be sure the data has been written
' to the Registry. Excessive calls to this function can impact system performance.</p>
'
Public Sub Flush()
    Call VerifyKey
    
    Dim E As Long
    E = RegFlushKey(mHKey)
    If E <> NO_ERROR Then IOError E, "RegistryKey.Flush"
End Sub

''
' Creates a new Windows Registry node.
'
' @param SubKey The name of the new node to be created. Any slashes will be trimmed from the ends.
' @return A <b>RegistryKey</b> to manipulate the new Registry node.
' @remarks If the Registry node already exists, then it is returned.
'
Public Function CreateSubKey(ByVal SubKey As String) As RegistryKey
    Call VerifyWritable
    Call VerifyKey
    
    SubKey = cString.Trim(SubKey, "\")
    
    If Len(SubKey) > MAX_KEY_LENGTH Then _
        Throw Cor.NewArgumentException("Key name cannot exceed 255 characters.", "Name")
    
    Dim E As Long
    Dim NewKey As Long
    E = RegCreateKeyEx(mHKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, NewKey, 0&)
    If E <> NO_ERROR Then IOError E, "RegistryKey"
    
    Set CreateSubKey = Cor.NewRegistryKey(NewKey, Me.Name, NewKey, SubKey, mWritable)
End Function

''
' Deletes a node from the Windows Registry.
'
' @param SubKey The node to be deleted.
' @param ThrowOnMissingSubKey Whether to throw an exception if the node was not found.
' @remarks The key node must be empty to be deleted. No sub keys or values are deleted.
' To delete all SubKeys use the <b>DeleteSubKeyTree</b> to recursively delete SubKeys
' and values.
'
Public Sub DeleteSubKey(ByVal SubKey As String, Optional ByVal ThrowOnMissingSubKey As Boolean = True)
    Call VerifyWritable
    
    Dim rk As RegistryKey
    Set rk = Me.OpenSubKey(SubKey)
    If rk Is Nothing Then
        If ThrowOnMissingSubKey Then _
            Throw Cor.NewArgumentException("The specified SubKey was not found.", "SubKey")
    Else
        If rk.SubKeyCount > 0 Then _
            Throw Cor.NewInvalidOperationException("Cannot delete. SubKey contains keys.")
        
        Dim E As Long
        E = RegDeleteKey(mHKey, SubKey)
        If E <> NO_ERROR Then IOError E, "RegistryKey"
        rk.CloseKey
    End If
End Sub

''
' Deletes all SubKeys within the specified SubKey to delete.
'
' @param SubKey The Registry node to be deleted.
' @remarks All subkeys and values are deleted recursively. Caution should be used.
'
Public Sub DeleteSubKeyTree(ByVal SubKey As String)
    Call VerifyWritable
    Call VerifyKey
    
    SubKey = cString.Trim(SubKey, "\")
    
    Dim rk As RegistryKey
    Set rk = Me.OpenSubKey(SubKey, True)
    
    Dim SubKeys() As String
    SubKeys = rk.GetSubKeyNames
    If cArray.GetLength(SubKeys) > 0 Then
        Dim i As Long
        For i = 0 To UBound(SubKeys)
            rk.DeleteSubKeyTree SubKeys(i)
        Next i
    End If
    Call rk.CloseKey
    
    Dim E As Long
    E = RegDeleteKey(mHKey, SubKey)
    If E <> NO_ERROR Then IOError E, "RegistryKey.DeleteSubKeyTree"
End Sub

''
' Returns a list of SubKey names within the current SubKey.
'
' @return The list of SubKeys.
' @remarks Values within the SubKey are not included in the list.
' <p>If not SubKeys exist then a zero-length array is returned.</p>
'
Public Function GetSubKeyNames() As String()
    Call VerifyKey
    
    Dim buf As String
    buf = String$(MAX_KEY_LENGTH, 0)
    
    Dim KeyCount As Long
    KeyCount = Me.SubKeyCount
    
    Dim ret() As String
    ret = cArray.CreateInstance(ciString, KeyCount)
    
    Do While KeyCount > 0
        KeyCount = KeyCount - 1
        
        Dim BufferSize As Long
        BufferSize = MAX_KEY_LENGTH
        
        Dim E As Long
        E = RegEnumKeyEx(mHKey, KeyCount, buf, BufferSize, 0, vbNullString, 0, 0@)
        If E <> NO_ERROR Then IOError E, "GetSubKeyNames"
        
        ret(KeyCount) = Left$(buf, BufferSize)
    Loop

    GetSubKeyNames = ret
End Function

''
' Returns a list of value names within the current SubKey.
'
' @return The list of value names.
' @remarks SubKeys within the SubKey are not included in the list.
'
' <p>Bad Spec: The MSDN says that if a default value exists, it will
' not be included as part of the names returned. However, a simple
' test in Dot NET shows that the default is included in the list.
' An empty string is the name of the default value. We include
' the default name ("") in the returned list.</p>
'
Public Function GetValueNames() As String()
    Call VerifyKey
    
    Dim buf As String
    buf = String$(MAX_KEY_LENGTH, 0)
    
    Dim ValueCount As Long
    ValueCount = Me.ValueCount
    
    Dim ret() As String
    ret = cArray.CreateInstance(ciString, ValueCount)
    
    Do While ValueCount > 0
        ValueCount = ValueCount - 1
        
        Dim BufferSize As Long
        BufferSize = MAX_KEY_LENGTH
        
        Dim E As Long
        E = RegEnumValue(mHKey, ValueCount, buf, BufferSize, 0, 0, ByVal 0&, 0&)
        If E <> NO_ERROR Then IOError E, "RegistryKey.GetValueNames"
        ret(ValueCount) = Left$(buf, BufferSize)
    Loop
    
    GetValueNames = ret
End Function

''
' Sets the value of a key value within the SubKey.
'
' @param Name The name of the value to set, or an empty string to set the key default value.
' @param Value The value to write to the Registry.
' @param ValueKind The type of value to be written to the registry.
' @remarks The registry can hold several types of data. <b>RegistryKey</b> supports
' many data types.<br>
' REG_DWORD (Integer numbers)<br>
' REG_QWORD (64-bit integer)<br>
' REG_BINARY (Byte arrays)<br>
' REG_SZ (Strings)<br>
' REG_MULTI_SZ (String arrays)<br>
' REG_EXPAND_SZ (Expandable string)<br>
' <p>Another value type is any class that implements the <b>cObject</b> interface. The
' <i>ToString</i> method is used to write the string representation to the Registry. This
' makes it easy to write such values as <b>cDateTime</b> with ease.</p>
' <p>If saving as a <b>REG_QWORD</b> then numeric values passed in may be adjusted to
' and mapped to an internal vbCurrency datatype to represent the 64-bit value accurately
' in the registry. If a vbLong value of 1 is passed in, then it will be set to a vbCurrency
' value of 0.0001. This will give the proper bit alignment with the 64-bit registry value.
' In order to avoid having the value mapped to a vbCurrency, pass the value in as a vbCurrency.
' In this case, the value is not changed and the 64-bits are saved to the registry as is.
' A vbCurrency of 1 will be saved to the registry as is, not as 0.0001.</p>
'
Public Sub SetValue(ByVal Name As String, ByRef Value As Variant, Optional ByVal ValueKind As RegistryValueKind = UnknownKind)
    Call VerifyWritable
    Call VerifyKey
    
    If Len(Name) >= MAX_KEY_LENGTH Then _
        Throw Cor.NewArgumentException("Registry Value name must be less than 255 characters.", "Name")

    If ValueKind = UnknownKind Then
        Call InternalSetValue(Name, Value, False)
    Else
        Call ClearException
        On Error GoTo errTrap
        Select Case ValueKind
            Case StringKind
                Call InternalSetValue(Name, Convert.ToString(Value), False)
                
            Case DWordKind
                Call InternalSetValue(Name, CLng(Value), False)
                
            Case BinaryKind
                If VarType(Value) <> (vbByte Or vbArray) Then _
                    Throw Cor.NewArgumentException("The value did not match the registry type specified by ValueKind.", "value")
                
                Call InternalSetValue(Name, Value, False)
                
            Case MultiStringKind
                If VarType(Value) <> (vbString Or vbArray) Then _
                    Throw Cor.NewArgumentException("The value did not match the registry type specified by ValueKind.", "value")
                    
                Call InternalSetValue(Name, Value, False)
                
            Case ExpandStringKind
                Call InternalSetValue(Name, Convert.ToString(Value), True)
                
            Case QWordKind
                If VarType(Value) = vbCurrency Then
                    Call InternalSetValue(Name, Value, False)
                Else
                    Call InternalSetValue(Name, 0.0001@ * CCur(Value), False)
                End If
            
            Case Else
                Throw Cor.NewArgumentException("Invalid ValueKind specified.", "ValueKind")
        End Select
    End If
    Exit Sub
    
errTrap:
    Dim ex As Exception
    If Catch(ex) Then
        If TypeOf ex Is IOException Then Throw ex
    End If
    Throw Cor.NewArgumentException("The value did not match the registry type specified by ValueKind.", "value")
End Sub

''
' Sets the keys default value.
'
' @param Value The default value for the key.
' @remarks The registry can hold several types of data. <b>RegistryKey</b> supports
' many data types.<br>
' REG_DWORD (Integer numbers)<br>
' REG_QWORD (64-bit integer)<br>
' REG_BINARY (Byte arrays)<br>
' REG_SZ (Strings)<br>
' REG_MULTI_SZ (String arrays)<br>
' REG_EXPAND_SZ (Expandable string)<br>
' <p>Another value type is any class that implements the <b>cObject</b> interface. The
' <i>ToString</i> method is used to write the string representation to the Registry. This
' makes it easy to write such values as <b>cDateTime</b> with ease.</p>
' <p>If saving as a <b>REG_QWORD</b> then numeric value passed in may be adjusted to
' and mapped to an internal vbCurrency datatype to represent the 64-bit value accurately
' in the registry. If a vbLong value of 1 is passed in, then it will be set to a vbCurrency
' value of 0.0001. This will give the proper bit alignment with the 64-bit registry value.
' In order to avoid having the value mapped to a vbCurrency, pass the value in as a vbCurrency.
' In this case, the value is not changed and the 64-bits are saved to the registry as is.
' A vbCurrency of 1 will be saved to the registry as is, not as 0.0001.</p>
'
Public Sub SetDefaultValue(ByRef Value As Variant)
    Call SetValue(DEFAULT_VALUE_NAME, Value)
End Sub

''
' Returns the value of a key within a SubKey.
'
' @param Name The name of the value to retrieve, or an empty string to retrieve the key default value.
' @param Default The value to be returned if the key name was not found.
' @param DoNotExpand Indicates if values of type REG_EXPAND_SZ should have their environment variables expanded.
' @return The value found in the Registry, or Empty if the value does not exist.
' @remarks <b>RegistryKey</b> supports the return of 6 different key value types.<br>
' REG_DWORD (Integer numbers)<br>
' REG_QWORD (64-bit integer)<br>
' REG_BINARY (Byte arrays)<br>
' REG_SZ (Strings)<br>
' REG_MULTI_SZ (String arrays)<br>
' REG_EXPAND_SZ (String containing an environment variable)<br>
' The specific type will be created and returned based on the type in the registry.
' <p>If retrieving a REG_QWORD, a vbCurrency variable is returned to represent the
' 64-bit value. This means that the value will have the bias of being divided by
' 10000. To get the true 64-bit value, the vbCurrency value will need to be multiplied
' by 10000. The result may be an overflow.
'
Public Function GetValue(ByVal Name As String, Optional ByRef Default As Variant, Optional ByVal DoNotExpand As Boolean = False) As Variant
    Dim ValueType As Long
    Dim ValueSize As Long
    Dim buf() As Byte
    
    Call VerifyKey

    If Len(Name) >= MAX_KEY_LENGTH Then _
        Throw Cor.NewArgumentException("Registry Value name must be less than 255 characters.", "Name")

    ValueSize = 128
    Do
        ReDim buf(0 To ValueSize - 1)
        
        Dim E As Long
        E = RegQueryValueEx(mHKey, Name, 0, ValueType, buf(0), ValueSize)
    Loop While (E = ERROR_MORE_DATA) And (mHKey <> HKEY_PERFORMANCE_DATA)
    
    If E <> ERROR_SUCCESS Then
        If Not IsMissing(Default) Then Call VariantCopyInd(GetValue, Default)
        Exit Function
    End If
    
    If (ValueSize > 0) And (ValueType <> REG_DWORD) And (ValueType <> REG_QWORD) Then
        If ValueSize <> UBound(buf) + 1 Then ReDim Preserve buf(0 To ValueSize - 1)
    End If
    
    Select Case ValueType
        Case REG_DWORD
            GetValue = AsLong(buf(0))
        
        Case REG_SZ
            If ValueSize > 0 Then
                GetValue = Left$(StrConv(buf, vbUnicode), ValueSize - 1)
            Else
                GetValue = ""
            End If
            
        Case REG_EXPAND_SZ
            If ValueSize > 0 Then
                GetValue = Left$(StrConv(buf, vbUnicode), ValueSize - 1)
                If DoNotExpand Then Exit Function
                GetValue = Environment.ExpandEnvironmentVariables(GetValue)
            Else
                GetValue = ""
            End If
            
        Case REG_MULTI_SZ
            If ValueSize > 2 Then
                Dim s As String
                s = StrConv(buf, vbUnicode)
                s = Left$(s, Len(s) - 2)
                GetValue = Split(s, vbNullChar)
            Else
                GetValue = cArray.CreateInstance(ciString)
            End If
            
        Case REG_BINARY, REG_NONE
            If ValueSize > 0 Then
                GetValue = buf
            Else
                GetValue = cArray.CreateInstance(ciByte)
            End If
        
        Case REG_QWORD
            GetValue = AsCurr(buf(0))
            
        Case Else
            Throw Cor.NewInvalidCastException("Could not convert registry value to a supported datatype.")
    End Select
End Function

''
' Returns the default value for the key.
'
' @param DoNotExpand Indicates if values of type REG_EXPAND_SZ should have their environment variables expanded.
' @return The default value for the key, or Empty if the value does not exist.
' @remarks <b>RegistryKey</b> supports the return of 6 different key value types.<br>
' REG_DWORD (Integer numbers)<br>
' REG_QWORD (64-bit integer)<br>
' REG_BINARY (Byte arrays)<br>
' REG_SZ (Strings)<br>
' REG_MULTI_SZ (String arrays)<br>
' REG_EXPAND_SZ (String containing an environment variable)<br>
' The specific type will be created and returned based on the type in the registry.
' <p>If retrieving a REG_QWORD, a vbCurrency variable is returned to represent the
' 64-bit value. This means that the value will have the bias of being divided by
' 10000. To get the true 64-bit value, the vbCurrency value will need to be multiplied
' by 10000. The result may be an overflow.
'
Public Function GetDefaultValue(Optional ByVal DoNotExpand As Boolean = False) As Variant
    Call Helper.MoveVariant(GetDefaultValue, GetValue("", , DoNotExpand))
End Function

''
' Deletes the value from the registry.
'
' @param Name The name of the value to be deleted, or an empty string to delete the default value for the key.
' @param ThrowOnMissingValue Whether to throw an exception if the value was not found.
'
Public Sub DeleteValue(ByVal Name As String, Optional ByVal ThrowOnMissingValue As Boolean = True)
    Call VerifyWritable
    Call VerifyKey
    
    Dim E As Long
    E = RegDeleteValue(mHKey, Name)
    If (E = ERROR_FILE_NOT_FOUND) And ThrowOnMissingValue Then _
        Throw Cor.NewArgumentException("Value was not found.", "Name")
End Sub

''
' Deletes the default value for the key.
'
Public Sub DeleteDefaultValue()
    Call DeleteValue(DEFAULT_VALUE_NAME, False)
End Sub

''
' Returns the number of SubKeys within the current key.
'
' @return The number of SubKeys.
'
Public Property Get SubKeyCount() As Long
    Call VerifyKey
    
    Dim E As Long
    E = RegQueryInfoKey(mHKey, vbNullString, ByVal 0&, ByVal 0&, SubKeyCount, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
    If E <> NO_ERROR Then IOError E, "RegistryKey"
End Property

''
' Returns the number of values within the current key.
'
' @remarks The number of values.
'
Public Property Get ValueCount() As Long
    Call VerifyKey
    
    Dim E As Long
    E = RegQueryInfoKey(mHKey, vbNullString, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ValueCount, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&)
    If E <> NO_ERROR Then IOError E, "RegistryKey"
End Property

''
' Returns the type of value stored in the registry.
'
' @param Name The name of the value to get the type of.
' @return The type of value in the registry.
'
Public Function GetValueKind(ByVal Name As String) As RegistryValueKind
    Dim E As Long
    E = RegQueryValueEx(mHKey, Name, 0, GetValueKind, ByVal 0&, 0)
    If E <> ERROR_SUCCESS Then IOError E
End Function

''
' Returns the last time the subkey was written to.
'
' @return A cDateTime object.
' @remarks This method returns cDateTime.MinValue on Windows 95,98 and ME.
'
Public Function GetLastWriteTime() As cDateTime
    Call VerifyKey
    
    If Environment.IsNT Then
        Dim E As Long
        Dim Time As Currency
        E = RegQueryInfoKey(mHKey, vbNullString, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, Time)
        If E <> NO_ERROR Then IOError E, "RegistryKey"
        Set GetLastWriteTime = cDateTime.FromFileTime(Time)
    Else
        Set GetLastWriteTime = cDateTime.MinValue
    End If
End Function

''
' Returns an enumerator to enumerate the name/value pairs.
'
' @param DoNotExpand Flag indicating if any expandable environment
' variables should be expanded or not before being returned.
' @param EnumType This allows to enumerate only the names or the values.
' @return An IEnumerator object.
'
Public Function GetEnumerator(Optional ByVal DoNotExpand As Boolean = False, Optional ByVal EnumType As IDictionaryEnumType = detEntries) As Object
    Dim ret As New RegistryKeyEnumerator
    Call ret.Init(Me, DoNotExpand, EnumType)
    Set GetEnumerator = ret
End Function

''
' Returns a For..Each compatible enumerator.
'
' @return A For..Each compatible enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal RootKey As Long, ByVal RootName As String, ByVal Key As Long, ByVal Name As String, ByVal Writable As Boolean)
    mRootKey = RootKey
    mRootName = cString.Trim(RootName, "\")
    mHKey = Key
    mName = cString.Trim(Name, "\")
    mWritable = Writable
End Sub

Friend Sub InitRemote(ByVal HKey As Long)
    mRootKey = HKey
    mHKey = HKey
    mWritable = True
    SetRootName
End Sub

Friend Sub VerifyKey()
    If mRootKey = 0 Then _
        Throw Cor.NewObjectDisposedException("RegistryKey", "The Registry Key is already disposed.")
    
    If mHKey = 0 Then
        Dim acc As Long
        If mWritable Then
            acc = KEY_ALL_ACCESS
        Else
            acc = KEY_READ
        End If
        
        Dim E As Long
        E = RegOpenKeyEx(mRootKey, mName, 0, acc, mHKey)
        If E <> NO_ERROR Then IOError E, "RegistryKey"
        
        Call SetRootName
    End If
End Sub

Friend Property Get HKey() As Long
    HKey = mHKey
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsSystemKey() As Boolean
    If mHKey < &H80000006 Then Exit Function
    If mHKey > &H80000000 Then Exit Function
    IsSystemKey = True
End Function

Private Sub VerifyWritable()
    If Not mWritable Then Throw Cor.NewUnauthorizedAccessException("The Registry Key is not writable.")
End Sub

Private Function SetValueString(ByRef Name As String, ByRef Value As Variant, ByVal AsExpandString As Boolean) As Long
    Dim b() As Byte
    b = StrConv(Value, vbFromUnicode)
    ReDim Preserve b(0 To UBound(b) + 1)
    SetValueString = RegSetValueEx(mHKey, Name, 0, IIf(AsExpandString, REG_EXPAND_SZ, REG_SZ), b(0), Len(Value))
End Function

Private Sub SetRootName()
    Select Case mHKey
        Case HKEY_CLASSES_ROOT:         mRootName = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_CONFIG:       mRootName = "HKEY_CURRENT_CONFIG"
        Case HKEY_CURRENT_USER:         mRootName = "HKEY_CURRENT_USER"
        Case HKEY_DYN_DATA:             mRootName = "HKEY_DYN_DATA"
        Case HKEY_LOCAL_MACHINE:        mRootName = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS:                mRootName = "HKEY_USERS"
        Case HKEY_PERFORMANCE_DATA:     mRootName = "HKEY_PERFORMANCE_DATA"
    End Select
End Sub

Private Sub InternalSetValue(ByRef Name As String, ByRef Value As Variant, ByVal AsExpandString As Boolean)
    Dim E As Long
    
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            E = RegSetValueEx(mHKey, Name, 0, REG_DWORD, CLng(Value), 4)
        
        Case vbObject
            If TypeOf Value Is IObject Then
                Dim Obj As IObject
                Set Obj = Value
                E = SetValueString(Name, Obj.ToString, False)
            Else
                Throw Cor.NewArgumentException("Objects must implement the IObject interface to be saved to the Registry.", "Value")
            End If
        
        Case vbString
            E = SetValueString(Name, Value, AsExpandString)
        
        Case vbByte Or vbArray
            Dim lpData As Long
            Dim Size As Long
            If Not cArray.IsNull(Value) Then
                Size = cArray.GetLength(Value)
                If Size > 0 Then lpData = MemLong(GetArrayPointer(Value) + PVDATA_OFFSET)
            End If
            
            E = RegSetValueEx(mHKey, Name, 0, REG_BINARY, ByVal lpData, Size)
        
        Case vbString Or vbArray
            Dim b() As Byte
            b = StrConv(Join(Value, vbNullChar) & (vbNullChar & vbNullChar), vbFromUnicode)
            E = RegSetValueEx(mHKey, Name, 0, REG_MULTI_SZ, b(0), cArray.GetLength(b))
            
        Case vbCurrency
            E = RegSetValueEx(mHKey, Name, 0, REG_QWORD, CCur(Value), 8)
        
        Case Else
            Throw Cor.NewArgumentException("Only integers, strings, byte arrays and string arrays can be saved to the Registry.", "Value")
    End Select
    If E <> NO_ERROR Then IOError E, "RegistryKey.SetValue"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call CloseKey
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
