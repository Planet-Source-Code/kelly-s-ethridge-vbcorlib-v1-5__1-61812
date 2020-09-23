VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListStatic"
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
'    Module: ArrayListStatic
'

''
' Represents the static methods of the ArrayList class.
'
' @remarks This class cannot be directly instantiated. In order to access
' the method of this class, use the <b>ArrayList.*</b> syntax.
' <pre>
' Set ReadOnlyList = ArrayList.ReadOnly(list)
' </pre>
' @see ArrayList
'
Option Explicit

''
' Returns an ArrayList object wrapped inside a readonly object.
'
' @param list The ArrayList to wrap
' @return The readonly ArrayList wrapper.
' @remarks Passing the object around provides access to the underlying
' ArrayList, however, the list cannot be changed in size or content.
'
Public Function ReadOnly(ByVal list As ArrayList) As ArrayList
    Dim ret As New ReadOnlyArrayList
    Call ret.Init(list)
    Set ReadOnly = ret
End Function

''
' Returns an ArrayList object that is wrapped in a fixedsize object.
'
' @param list The ArrayList to be wrapped.
' @return The wrapped ArrayList.
' @remarks Passing this object around provides access to the underlying
' ArrayList, however, prevents the list from changing size. The current
' values in the list can be modified, but the size of the list cannot change.
'
Public Function FixedSize(ByVal list As ArrayList) As ArrayList
    Dim ret As New FixedSizeArrayList
    Call ret.Init(list)
    Set FixedSize = ret
End Function

''
' Creates a new ArrayList and fills it with the supplied value the specified number of times.
'
' @param Value The value to fill the ArrayList with.
' @param Count The number of times the value is to be added to the list.
' @return A new ArrayList.
'
Public Function Repeat(ByRef Value As Variant, ByVal Count As Long) As ArrayList
    Set Repeat = New ArrayList
    Repeat.Capacity = Count
    
    Dim i As Long
    For i = 1 To Count
        Call Repeat.Add(Value)
    Next i
End Function
