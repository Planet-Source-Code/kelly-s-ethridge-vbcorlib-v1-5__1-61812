VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangedArrayList"
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
'    Module: RangedArrayList
'

''
' This class is used to represent a sub-array view of the parent ArrayList object.
'

Option Explicit
Implements IObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private mList       As ArrayList
Private mVersion    As Long
Private mIndex      As Long
Private mCount      As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal list As ArrayList, ByVal Index As Long, ByVal Count As Long)
    Set mList = list
    mVersion = list.Version
    mIndex = Index
    mCount = Count
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyVersion()
    If mVersion <> mList.Version Then _
        Throw Cor.NewInvalidOperationException("The base ArrayList has been modified outside of the Ranged list.")
End Sub

Private Function GetCollectionSize(ByRef c As Variant) As Long
    If IsArray(c) Then
        If cArray.IsNull(c) Then _
            Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "c")
        
        GetCollectionSize = UBound(c) - LBound(c) + 1
    
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Collection), "c")
        
        If TypeOf c Is Collection Then
            Dim col As Collection
            Set col = c
            GetCollectionSize = col.Count
        ElseIf TypeOf c Is ICollection Then
            Dim icol As ICollection
            Set icol = c
            GetCollectionSize = icol.Count
        Else
            Throw Cor.NewInvalidCastException("An ICollection or VBA.Collection object is required.")
        End If
    Else
        Throw Cor.NewInvalidCastException("An ICollection object, VBA.Collection or an Array is required.")
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    Call VerifyVersion
    Call mList.Insert(mIndex + mCount, Value)
    ArrayList_Add = mCount
    mCount = mCount + 1
    mVersion = mVersion + 1
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Call VerifyVersion
    Call mList.InsertRange(mIndex + mCount, c)
    mCount = mCount + GetCollectionSize(c)
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mCount, Index, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim i As Long
    i = mList.BinarySearch(Value, mIndex + ElemIndex, ElemCount, comparer)
    If i >= 0 Then i = i - mIndex
    ArrayList_BinarySearch = i
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    Call VerifyVersion
    mList.Capacity = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Capacity() As Long
    Call VerifyVersion
    ArrayList_Capacity = mList.Capacity
End Property

Private Sub ArrayList_Clear()
    Call VerifyVersion
    If mCount = 0 Then Exit Sub
    Call mList.RemoveRange(mIndex, mCount)
    mCount = 0
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_Clone() As Object
    Call VerifyVersion
    Dim ret As New RangedArrayList
    Call ret.Init(mList.Clone, mIndex, mCount)
    Set ArrayList_Clone = ret
End Function

Private Function ArrayList_Contains(Value As Variant, Optional ByVal comparer As IComparer) As Boolean
    Call VerifyVersion
    ArrayList_Contains = (mList.IndexOf(Value, mIndex, mCount, comparer) >= 0)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional arrayindex As Variant)
    Call VerifyVersion
    
    Dim Index As Long
    If IsMissing(arrayindex) Then Index = LBound(DstArray) Else Index = arrayindex
    Call mList.CopyToEx(mIndex, DstArray, Index, mCount)
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal arrayindex As Long, ByVal Count As Long)
    Call VerifyVersion
    
    Dim result As Long
    result = VerifyListRange(mCount, Index, Count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, Index, "Index", Count, "Count")
    
    Call mList.CopyToEx(mIndex + Index, DstArray, arrayindex, Count)
End Sub

Private Property Get ArrayList_Count() As Long
    Call VerifyVersion
    ArrayList_Count = mCount
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = IObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Object
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mCount, StartIndex, ElemIndex, Count, ElemCount)
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
    result = VerifyListRange(mCount, Index, Count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, Index, "Index", Count, "Count")
    
    Dim ret As RangedArrayList
    Set ret = New RangedArrayList
    Call ret.Init(Me, Index, Count)
    Set ArrayList_GetRange = ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mCount, StartIndex, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))
    
    Dim Index As Long
    Index = mList.IndexOf(Value, mIndex + ElemIndex, ElemCount, comparer)
    If Index >= 0 Then Index = Index - mIndex
    ArrayList_IndexOf = Index
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    Call VerifyVersion
    
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    Call mList.Insert(mIndex + Index, Value)
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    Call VerifyVersion
    
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    Dim TotalElements As Long
    TotalElements = GetCollectionSize(c)
    Call mList.InsertRange(mIndex + Index, c)
    mCount = mCount + TotalElements
    mVersion = mVersion + 1
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = mList.IsFixedSize
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = mList.IsReadOnly
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Call VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    Set mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Call VerifyVersion
    
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    Call VerifyVersion
    
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    Helper.MoveVariant ArrayList_Item, mList(mIndex + Index)
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer) As Long
    Call VerifyVersion
    
    Dim ElemIndex As Long
    Dim ElemCount As Long
    If GetOptionalLongPair(StartIndex, mCount - 1, ElemIndex, Count, mCount, ElemCount) = Argument_ParamRequired Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ParamRequired), IIf(IsMissing(StartIndex), "StartIndex", "Count"))
    If ElemIndex >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_UBound), "StartIndex", ElemIndex)
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", ElemCount)
    If ElemIndex - ElemCount + 1 < 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")

    Dim Index As Long
    Index = mList.LastIndexOf(Value, mIndex + ElemIndex, ElemCount, comparer)
    If Index >= 0 Then Index = Index - mIndex
    ArrayList_LastIndexOf = Index
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    Call VerifyVersion
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(Value As Variant, Optional ByVal comparer As IComparer)
    Call VerifyVersion
    
    Dim Index As Long
    Index = mList.IndexOf(Value, mIndex, mCount, comparer)
    If Index >= 0 Then
        Call mList.RemoveAt(Index)
        mCount = mCount - 1
        mVersion = mVersion + 1
    End If
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    Call VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    Call mList.RemoveAt(mIndex + Index)
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Call VerifyVersion
    
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), IIf(Index < 0, "index", "count"), IIf(Index < 0, Index, Count))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "count")
    
    Call mList.RemoveRange(mIndex + Index, Count)
    mCount = mCount - Count
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mCount, Index, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Call mList.Reverse(mIndex + ElemIndex, ElemCount)
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    Call VerifyVersion
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Iindex", Index)
    
    Dim TotalElements As Long
    TotalElements = GetCollectionSize(c)
    If Index + TotalElements > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LargerThanCollection), "index", Index)
    
    Call mList.SetRange(mIndex + Index, c)
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Sort(Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal comparer As IComparer)
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mCount, StartIndex, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))

    Call mList.Sort(mIndex + ElemIndex, ElemCount, comparer)
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As ciArrayTypes = 12&) As Variant
    Call VerifyVersion
    
    Dim ret As Variant
    ret = cArray.CreateInstance(ArrayType, mCount)
    If mCount > 0 Then Call mList.CopyToEx(mIndex, ret, 0, mCount)
    ArrayList_ToArray = ret
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = IObject_ToString
End Function

Private Sub ArrayList_TrimToSize()
    Call VerifyVersion
    mList.TrimToSize
    mVersion = mVersion + 1
End Sub

Private Property Get ArrayList_Version() As Long
    ArrayList_Version = mVersion
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
    ArrayList_Clear
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


