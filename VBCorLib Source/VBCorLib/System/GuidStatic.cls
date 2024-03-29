VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuidStatic"
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
'    Module: GuidStatic
'

''
' Provides static methods used to create and manipulate <b>Guid</b> objects.
'
' @remarks The name Guid conflicts with a hidden type in VB. In order to
' specific the VBCorLib version of Guid, a variable must be defined as
' VBCorLib.Guid instead of just Guid.
' </p>This class cannot be directly created. To access the methods, use
' the variable name directly.
' <pre>
' Dim g As VBCorLib.Guid
' Set g = Guid.Parse("{533217B3-CDEC-40A2-B01C-1EA8593B850F}")
' </pre>
' </p>
'
' @see Guid
'
Option Explicit



''
' Returns a new <b>Guid</b> object with a random guid generated.
'
' @param LockGuid Locks the guids handle.
' @return A new <b>Guid</b> object.
' @see Guid
'
Public Function NewGuid() As Guid
    Set NewGuid = New Guid
    NewGuid.NewGuid
End Function

''
' Parses a string representation of a guid, returning a <b>Guid</b>
' containing the parsed value.
'
' @param strGuid The guid string to be parsed.
' @param LockGuid Locks the guids handle.
' @return A <b>Guid</b> set to the parsed value.
'
Public Function Parse(ByVal strGuid As String) As Guid
    Set Parse = New Guid
    Call Parse.Parse(strGuid)
End Function

''
' Converts an array of 16 bytes into a Guid.
'
' @param Bytes The bytes to convert to a guid.
' @param LockGuid Locks the guids handle.
' @return A new guid object set to the 16 bytes.
'
Public Function FromByteArray(ByRef Bytes() As Byte) As Guid
    Set FromByteArray = New Guid
    Call FromByteArray.FromByteArray(Bytes)
End Function

''
' Creates a guid using specified values and byte array.
'
' @param a A value representing 12345678-xxxx-xxxx-xxxx-xxxxxxxxxxxx
' @param b A value representing xxxxxxxx-1234-xxxx-xxxx-xxxxxxxxxxxx
' @param c A value representing xxxxxxxx-xxxx-1234-xxxx-xxxxxxxxxxxx
' @param d An array representing xxxxxxxx-xxxx-xxxx-1234-123456789012
' @param LockGuid Locks the guids handle.
' @return A new guid initialized to the values and byte array specified.
'
Public Function FromParts(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, ByRef d() As Byte) As Guid
    Set FromParts = New Guid
    Call FromParts.FromParts(a, b, c, d)
End Function

''
' Creates a new guid from the individual values and bytes.
'
' @param a A Long
' @param b An Integer
' @param c An Integer
' @param d A Byte
' @param e A Byte
' @param f A Byte
' @param g A Byte
' @param h A Byte
' @param i A Byte
' @param j A Byte
' @param k A Byte
' @param LockGuid Locks the guids handle.
' @return A new guid initialized to the values and bytes specified.
'
Public Function FromValues(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, ByVal d As Byte, ByVal E As Byte, ByVal f As Byte, ByVal g As Byte, ByVal h As Byte, ByVal i As Byte, ByVal j As Byte, ByVal k As Byte) As Guid
    Dim Bytes(7) As Byte
    Bytes(0) = d
    Bytes(1) = E
    Bytes(2) = f
    Bytes(3) = g
    Bytes(4) = h
    Bytes(5) = i
    Bytes(6) = j
    Bytes(7) = k
    
    Set FromValues = FromParts(a, b, c, Bytes)
End Function

''
' Creates a Guid based on a 16-byte block of memory.
'
' @param lpGuid A poiner to the memory location of the 16 bytes.
' @return A new guid object.
'
Public Function FromMemory(ByVal lpGuid As Long) As Guid
    Dim b(15) As Byte
    Call CopyMemory(b(0), ByVal lpGuid, SIZEOF_GUID)
    Set FromMemory = FromByteArray(b)
End Function

''
' Returns a Guid already set to the IUnknown interface.
'
' @return A preconfigured Guid object of type IUnknown.
'
Public Property Get IUnknownGuid() As Guid
    Set IUnknownGuid = Guid.Parse("{00000000-0000-0000-C000-000000000046}")
End Property

''
' Returns a Guid already set to the IDispatch (Object) interface.
'
' @return A preconfigured Guid object of type IDispatch (Object).
'
Public Property Get IDispatchGuid() As Guid
    Set IDispatchGuid = Guid.Parse("{00020400-0000-0000-C000-000000000046}")
End Property

''
' Returns a Read-Only wrapper for the supplied Guid object.
'
' @param g The Guid object to wrap as Read-Only.
' @return A Read-Only Guid object.
' @remarks This is to allow the passing of a Guid object around an
' application without allowing the guid to be modified through the
' Handle property, which allows direct access to the guid structure.
'
Public Function ReadOnly(ByVal g As Guid) As Guid
    Dim ret As New ReadOnlyGuid
    Call ret.Init(g)
    Set ReadOnly = ret
End Function
