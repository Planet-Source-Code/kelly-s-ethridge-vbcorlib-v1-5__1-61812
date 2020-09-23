VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKeyEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: RegistryKeyEnumerator
'

'
' This is the base implementation of an enumerator class.
'
' To use this class in a For..Each loop, it must be passed
' into the CreateEnumerator function. The result of that
' function is then returned through the NewEnum function.
'
' Public Function NewEnum() As IUnknown
'     Dim ret As New MyEnumerator
'     ret.Init Me                           ' Me would be the outer object, such as a list to be enumerated.
'     Set NewEnum = CreateEnumerator(ret)
' End Function
'
Option Explicit
Implements IObject
Implements IEnumerator
Implements IDictionaryEnumerator

Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Byte, ByRef lpcbData As Long) As Long

Private mKey            As RegistryKey
Private mIndex          As Long
Private mName           As String
Private mNameSize       As Long
Private mData()         As Byte
Private mFinished       As Boolean
Private mEnumType       As IDictionaryEnumType
Private mValueType      As RegistryValueKind
Private mValueSize      As Long
Private mDoNotExpand    As Boolean



''
' We place this function here to allos the caller to
' use the GetEnumerator function directly when starting
' a For..Each loop. The user may need to alter the enumeration
' with the parameters, so the NewEnum in the RegistryKey
' object will not get called. Instead, this one will be called.
'
' @return For..Each compatible enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Key As RegistryKey, ByVal DoNotExpand As Boolean, ByVal EnumType As IDictionaryEnumType)
    mEnumType = EnumType
    mDoNotExpand = DoNotExpand
    Set mKey = Key
    Call Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetValue() As Variant
    If (mValueSize > 0) And (mValueType <> REG_DWORD) And (mValueType <> REG_QWORD) Then
        If mValueSize <> UBound(mData) + 1 Then ReDim Preserve mData(0 To mValueSize - 1)
    End If
    
    Select Case mValueType
        Case REG_DWORD
            GetValue = AsLong(mData(0))
        
        Case REG_SZ
            If mValueSize > 0 Then
                GetValue = Left$(StrConv(mData, vbUnicode), mValueSize - 1)
            Else
                GetValue = ""
            End If
            
        Case REG_EXPAND_SZ
            If mValueSize > 0 Then
                GetValue = Left$(StrConv(mData, vbUnicode), mValueSize - 1)
                If mDoNotExpand Then Exit Function
                GetValue = Environment.ExpandEnvironmentVariables(GetValue)
            Else
                GetValue = ""
            End If
            
        Case REG_MULTI_SZ
            If mValueSize > 2 Then
                Dim s As String
                s = StrConv(mData, vbUnicode)
                s = Left$(s, Len(s) - 2)
                GetValue = Split(s, vbNullChar)
            Else
                GetValue = cArray.CreateInstance(ciString)
            End If
            
        Case REG_BINARY, REG_NONE
            If mValueSize > 0 Then
                GetValue = mData
            Else
                GetValue = cArray.CreateInstance(ciByte)
            End If
        
        Case REG_QWORD
            GetValue = AsCurr(mData(0))
            
        Case Else
            Throw Cor.NewInvalidCastException("Could not convert registry value to a supported datatype.")
    End Select
End Function

Private Sub VerifyEnum()
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mFinished Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mName = String$(MAX_PATH, 0)
    ReDim mData(0 To 127)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionaryEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    Call Helper.MoveVariant(IDictionaryEnumerator_Current, IEnumerator_Current)
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    Call VerifyEnum
    Set IDictionaryEnumerator_Entry = Cor.NewDictionaryEntry(mName, GetValue)
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    Call VerifyEnum
    IDictionaryEnumerator_Key = Left$(mName, mNameSize)
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    IDictionaryEnumerator_MoveNext = IEnumerator_MoveNext
End Function

Private Sub IDictionaryEnumerator_Reset()
    Call IEnumerator_Reset
End Sub

Private Property Get IDictionaryEnumerator_Value() As Variant
    Call VerifyEnum
    Call Helper.MoveVariant(IDictionaryEnumerator_Value, GetValue)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Call VerifyEnum
    
    Select Case mEnumType
        Case detEntries
            Set IEnumerator_Current = Cor.NewDictionaryEntry(Left$(mName, mNameSize), GetValue)
        Case detKeys
            Call VariantCopy(IEnumerator_Current, Left$(mName, mNameSize))
        Case detValues
            Call VariantCopy(IEnumerator_Current, GetValue)
    End Select
End Property

Private Function IEnumerator_MoveNext() As Boolean
    mIndex = mIndex + 1
    Call VerifyEnum
    
    mValueSize = UBound(mData) + 1
    Do
        Dim HKey As Long
        HKey = mKey.HKey
        If HKey = 0 Then _
            Throw Cor.NewInvalidOperationException("The Registry Key has been closed.")
        
        mNameSize = MAX_PATH
        
        Dim E As Long
        E = RegEnumValue(HKey, mIndex, mName, mNameSize, 0, mValueType, mData(0), mValueSize)
        
        If E = ERROR_MORE_DATA Then ReDim mData(0 To mValueSize - 1)
    Loop While (E = ERROR_MORE_DATA) And (HKey <> HKEY_PERFORMANCE_DATA)
    
    If E = ERROR_NO_MORE_ITEMS Then mFinished = True
    
    IEnumerator_MoveNext = (E = ERROR_SUCCESS)
End Function

Private Sub IEnumerator_Reset()
    mIndex = -1
    mFinished = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        IObject_Equals = (Value Is Me)
    End If
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = ObjPtr(CUnk(Me))
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Object.ToString(Me, App)
End Function

