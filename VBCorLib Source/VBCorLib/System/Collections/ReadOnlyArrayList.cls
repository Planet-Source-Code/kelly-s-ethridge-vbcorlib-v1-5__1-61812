VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: ReadOnlyArrayList
'

''
' Wraps an ArrayList object in a read-only wrapper.

Option Explicit
Implements IObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private mList As ArrayList


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal list As ArrayList)
    Set mList = list
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    ArrayList_BinarySearch = mList.BinarySearch(Value, StartIndex, Count, comparer)
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Property

Private Property Get ArrayList_Capacity() As Long
    ArrayList_Capacity = mList.Capacity
End Property

Private Sub ArrayList_Clear()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Function ArrayList_Clone() As Object
    Dim ret As New ReadOnlyArrayList
    Call ret.Init(mList.Clone)
    Set ArrayList_Clone = ret
End Function

Private Function ArrayList_Contains(Value As Variant, Optional ByVal comparer As IComparer) As Boolean
    ArrayList_Contains = mList.Contains(Value, comparer)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional arrayindex As Variant)
    Call mList.CopyTo(DstArray, arrayindex)
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal arrayindex As Long, ByVal Count As Long)
    Call mList.CopyToEx(Index, DstArray, arrayindex, Count)
End Sub

Private Property Get ArrayList_Count() As Long
    ArrayList_Count = mList.Count
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = IObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Object
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mList.Count, StartIndex, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))

    Dim ret As New ArrayListEnumerator
    Call ret.Init(Me, ElemIndex, ElemCount)
    Set ArrayList_GetEnumerator = ret
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = IObject_GetHashcode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    Dim result As Long
    result = VerifyListRange(mList.Count, Index, Count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, Index, "Index", Count, "Count")
    
    Dim ret As RangedArrayList
    Set ret = New RangedArrayList
    Call ret.Init(Me, Index, Count)
    Set ArrayList_GetRange = ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    ArrayList_IndexOf = mList.IndexOf(Value, StartIndex, Count, comparer)
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = True
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = True
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    Call Helper.MoveVariant(ArrayList_Item, mList(Index))
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    ArrayList_LastIndexOf = mList.LastIndexOf(Value, StartIndex, Count, comparer)
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(Value As Variant, Optional ByVal comparer As IComparer)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub ArrayList_Sort(Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As ciArrayTypes = 12&) As Variant
    ArrayList_ToArray = mList.ToArray(ArrayType)
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = mList.ToString
End Function

Private Sub ArrayList_TrimToSize()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Property Get ArrayList_Version() As Long
    ArrayList_Version = mList.Version
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = ArrayList_Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Object.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = ObjPtr(CUnk(Me))
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long)
    Call ArrayList_CopyTo(Arr, Index)
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = ArrayList_Count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = ArrayList_NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = ArrayList_NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(Value As Variant) As Long
    IList_Add = ArrayList_Add(Value)
End Function

Private Sub IList_Clear()
    Call ArrayList_Clear
End Sub

Private Function IList_Contains(Value As Variant, Optional ByVal comparer As IComparer) As Boolean
    IList_Contains = ArrayList_Contains(Value, comparer)
End Function

Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long)
    Call ArrayList_CopyTo(Arr, Index)
End Sub

Private Property Get IList_Count() As Long
    IList_Count = ArrayList_Count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IList_IndexOf(Value As Variant, Optional ByVal comparer As IComparer) As Long
    IList_IndexOf = ArrayList_IndexOf(Value, , , comparer)
End Function

Private Sub IList_Insert(ByVal Index As Long, Value As Variant)
    Call ArrayList_Insert(Index, Value)
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = ArrayList_IsFixedSize
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = ArrayList_IsReadOnly
End Property

Private Property Set IList_Item(ByVal Index As Long, RHS As Variant)
    Set ArrayList_Item(Index) = RHS
End Property

Private Property Let IList_Item(ByVal Index As Long, RHS As Variant)
    ArrayList_Item(Index) = RHS
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Call Helper.MoveVariant(IList_Item, ArrayList_Item(Index))
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = ArrayList_NewEnum
End Function

Private Sub IList_Remove(Value As Variant, Optional ByVal comparer As IComparer)
    Call ArrayList_Remove(Value, comparer)
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    Call ArrayList_RemoveAt(Index)
End Sub


