VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpanStatic"
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
'    Module: TimeSpanStatic
'

''
'   Provides severs static functions to help in creating TimeSpan objects.
'
' @see TimeSpan
'
Option Explicit
Private Const MILLISECONDS_PER_SECOND       As Long = 1000
Private Const SECONDS_PER_MINUTE            As Long = 60
Private Const MINUTES_PER_HOUR              As Long = 60
Private Const HOURS_PER_DAY                 As Long = 24

Private Const MILLISECONDS_PER_MINUTE       As Long = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR         As Long = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY          As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Private Const CHAR_MINUS                As Long = 45
Private Const CHAR_COLON                As Long = 58
Private Const CHAR_PERIOD               As Long = 46
Private Const CHAR_0                    As Long = 48
Private Const CHAR_9                    As Long = 57

Private mZero           As TimeSpan
Private mMaxValue       As TimeSpan
Private mMinValue       As TimeSpan
Private mStringBuilder  As StringBuilder
Private mMSPlaces()     As Currency


''
' Returns a TimeSpan with a value of zero.
'
' @return A zero timespan.
'
Public Property Get Zero() As TimeSpan
    Set Zero = mZero
End Property

''
' Returns a TimeSpan with a maximum value.
'
' @return A maximum value timespan.
'
Public Property Get MaxValue() As TimeSpan
    Set MaxValue = mMaxValue
End Property

''
' Returns a TimeSpan with a minimum value.
'
' @return A minimum value timespan.
'
Public Property Get MinValue() As TimeSpan
    Set MinValue = mMinValue
End Property

''
' Returns a TimeSpan created from milliseconds.
'
' @param Value The number of milliseconds the timespan will represent.
' @return The TimeSpan initialized to the specified number of milliseconds.
'
Public Function FromMilliseconds(ByVal Value As Currency) As TimeSpan
    Set FromMilliseconds = New TimeSpan
    Call FromMilliseconds.InitFromMilliseconds(Value)
End Function

''
' Returns a TimeSpan created from seconds.
'
' @param Value The number of seconds the timespan will represent.
' @return The TimeSpan initialized to the specified number of seconds.
'
Public Function FromSeconds(ByVal Value As Double) As TimeSpan
    Set FromSeconds = FromMilliseconds(Value * MILLISECONDS_PER_SECOND)
End Function

''
' Returns a TimeSpan created from minutes.
'
' @param Value The number of minutes the timespan will represent.
' @return The TimeSpan initialized to the specified number of minutes.
'
Public Function FromMinutes(ByVal Value As Double) As TimeSpan
    Set FromMinutes = FromMilliseconds(Value * MILLISECONDS_PER_MINUTE)
End Function

''
' Returns a TimeSpan created from hours.
'
' @param Value The number of hours the timespan will represent.
' @return The TimeSpan initialized to the specified number of hours.
'
Public Function FromHours(ByVal Value As Double) As TimeSpan
    Set FromHours = FromMilliseconds(Value * MILLISECONDS_PER_HOUR)
End Function

''
' Returns a TimeSpan created from days.
'
' @param Value The number of days the timespan will represent.
' @return The TimeSpan initialized to the specified number of days.
'
Public Function FromDays(ByVal Value As Double) As TimeSpan
    Set FromDays = FromMilliseconds(Value * MILLISECONDS_PER_DAY)
End Function

''
' Returns a TimeSpan created from ticks.
'
' @param Value The number of ticks the timespan will represent.
' @return The TimeSpan initialized to the specified number of ticks.
'
Public Function FromTicks(ByRef Value As Variant) As TimeSpan
    Set FromTicks = FromMilliseconds(CDec(Value) / 10000)
End Function

''
' Parses a string in the format of [d.]h:mm:ss[.ff] into a TimeSpan.
'
' @param s A string containing the date to be parsed.
' @return A TimeSpan created from the parsed string.
'
Public Function Parse(ByVal s As String) As TimeSpan
    Dim buf As WordBuffer
    Dim pos As Long
    Dim l As Long
    Dim hh As Long
    Dim mm As Long
    Dim ss As Long
    Dim ff As Currency
    Dim d As Long
    Dim i As Long
    Dim sign As Long
    
    s = cString.Trim(s)
    
    l = Len(s)
    Call InitWordBuffer(buf, StrPtr(s), l + 1)
    
    If buf.Data(0) = CHAR_MINUS Then
        sign = -1
        pos = 1
    Else
        sign = 1
    End If
    
    i = GetComponent(buf, pos, 1) * sign
    Select Case buf.Data(pos)
        Case CHAR_COLON
            hh = i
        Case CHAR_PERIOD
            d = i
            pos = pos + 1
            hh = GetComponent(buf, pos, 1) * sign
        Case Else
            Throw New FormatException
    End Select
    pos = pos + 1
    
    mm = GetComponent(buf, pos, 2) * sign
    If buf.Data(pos) <> CHAR_COLON Then Throw New FormatException
    pos = pos + 1
    ss = GetComponent(buf, pos, 2) * sign
    
    If pos < l Then
        If buf.Data(pos) <> CHAR_PERIOD Then Throw New FormatException
        pos = pos + 1
        ff = GetFraction(buf, pos) * sign
    End If
    If pos < l Then Throw New FormatException

    Set Parse = FromMilliseconds((d * MILLISECONDS_PER_DAY) + (hh * MILLISECONDS_PER_HOUR) + (mm * MILLISECONDS_PER_MINUTE) + (ss * MILLISECONDS_PER_SECOND) + ff)
End Function

''
' Returns a TimeSpan object representing the time of day of the date.
'
' @param Value The date to retrieve the time of day for.
' @return A TimeSpan representing the time of day.
'
Public Function FromDate(ByRef Value As Variant) As TimeSpan
    Dim dt As cDateTime
    Set dt = cDateTime.GetcDateTime(Value)
    Set FromDate = dt.TimeOfDay
End Function



' This is to prevent a TimeSpan from having to create and
' destroy an object. A TimeSpan may live very short lives when
' doing calculations. If for some reason a large number of
' TimeSpans is created and also converted to a string using
' ToString, the overhead of having to create and destroy a
' StringBuilder object will cost twice as much overhead as
' creating and destroying just the TimeSpan object.
'
' This is to be used only by TimeSpan objects in an immediate
' context since it is shared by all TimeSpan objects.
Friend Property Get StringBuilder() As StringBuilder
    mStringBuilder.Length = 0
    Set StringBuilder = mStringBuilder
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetFraction(ByRef buf As WordBuffer, ByRef pos As Long) As Currency
    Dim ch As Integer
    Dim ret As Currency
    Dim place As Long
    
    ch = buf.Data(pos)
    Do While ch <> 0
        If place >= 7 Then Throw New FormatException
        Select Case ch
            Case CHAR_0 To CHAR_9
                ret = ret + mMSPlaces(place) * (ch - CHAR_0)
                place = place + 1
            Case Else
                Throw New FormatException
        End Select
        pos = pos + 1
        ch = buf.Data(pos)
    Loop
    GetFraction = ret
End Function

Private Function GetComponent(ByRef buf As WordBuffer, ByRef pos As Long, ByVal MinChars As Long) As Long
    Dim ch As Integer
    Dim ret As Long
    
    ch = buf.Data(pos)
    Do While ch <> 0
        Select Case ch
            Case CHAR_0 To CHAR_9
                ret = ret * 10 + ch - CHAR_0
            Case Else
                If MinChars <= 0 Then Exit Do
                Throw New FormatException
        End Select
        MinChars = MinChars - 1
        pos = pos + 1
        ch = buf.Data(pos)
    Loop
    GetComponent = ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mZero = New TimeSpan
    Set mMaxValue = FromMilliseconds(922337203685477.5807@)
    Set mMinValue = FromMilliseconds(-922337203685477.5807@ - 0.0001@)
    Set mStringBuilder = New StringBuilder
    mMSPlaces = cArray.NewArray(ciCurrency, 100@, 10@, 1@, 0.1@, 0.01@, 0.001@, 0.0001@)
End Sub
