VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: StopWatch
'

''
' Provides high resolution timing using the QueryPerformanceCount API.
'
Option Explicit
Implements IObject

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private mStartCount As Currency
Private mFrequency  As Currency


''
' Resets the StopWatch and begins counting.
'
' @remarks When the function is called, the <b>StopWatch</b>
' immediately begins the timing process. This function should
' be called immediately before the process that is to be timed.
'
Public Sub Reset()
    Call QueryPerformanceCounter(mStartCount)
End Sub

''
' Returns the time elapsed in milliseconds, including
' fractions of a millisecond (ticks).
'
' @return The time elapsed in milliseconds.
'
Public Property Get Milliseconds() As Currency
    Dim StopCount As Currency
    Call QueryPerformanceCounter(StopCount)
    Milliseconds = mFrequency * (StopCount - mStartCount)
End Property

''
' Returns a formatted version of the elapsed time.
'
' @return The formatted time elapsed.
' @remarks The format is "0.0000 ms".
'
Public Function ToString() As String
    ToString = Format$(Milliseconds, "0.0000 ms")
End Function

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to test for equality against this object instance.
' @return The result of the equality comparison.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr.
'
' @return The hashcode for this object instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init()
    If QueryPerformanceFrequency(mFrequency) = BOOL_FALSE Then _
        Throw Cor.NewNotSupportedException("Hardware does not support High Performance Counters.")
    mFrequency = 1000@ / mFrequency
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call Init
End Sub


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
