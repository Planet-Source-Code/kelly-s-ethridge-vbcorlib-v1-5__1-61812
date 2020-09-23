VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cString"
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
'    Module: cString
'

''
' Provides static functions to format and manipulate string values.
'
' @remarks This class cannot be directly instantiated. The methods are accessed
' through the class name itself.
' <pre>
' s = cString.Trim(s, "\")
' </pre>
'
Option Explicit

Private mBuilder        As StringBuilder
Private mString         As WordBuffer
Private mTrimChars      As WordBuffer
Private mJoinSA         As SafeArray1d
Private mJoinArray()    As String



''
' Returns if the string is null.
'
' @param s The string to be tested for null.
' @return Indication if the string is null.
' @remarks If the string is null it will return True. An empty string
' will return false. A null string is equal to StrPtr(s) = 0.
'
Public Function IsNull(ByRef s As String) As Boolean
    IsNull = (StrPtr(s) = 0)
End Function

''
' Returns a string from the characters in an integer array.
'
' @param Chars The array to create the string from.
' @param Index The starting index in the array.
' @param Count The number of characters to use.
' @return A string created from the chars in an integer array.
'
Public Function FromCharArray(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim result As Long
    
    result = GetOptionalArrayRange(SAPtr(Chars), Index, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowArrayRangeException(result, "Chars", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    FromCharArray = SysAllocStringLen(VarPtr(Chars(ElemIndex)), ElemCount)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\Includes\cString.Format.txt"
Public Function Format(ByRef fmt As String, ParamArray args() As Variant) As String
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16))
    Format = FormatArrayEx(Nothing, fmt, vArgs)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param provider A custom format provider that is used to format the argument instead of the default formatting.
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Public Function FormatEx(ByVal provider As IFormatProvider, ByRef fmt As String, ParamArray args() As Variant) As String
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(20))
    FormatEx = FormatArrayEx(provider, fmt, vArgs)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Public Function FormatArray(ByRef fmt As String, ByRef args() As Variant) As String
    FormatArray = FormatArrayEx(Nothing, fmt, args)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param provider A custom format provider that is used to format the argument instead of the default formatting.
' @param fmt The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Public Function FormatArrayEx(ByVal provider As IFormatProvider, ByRef fmt As String, ByRef args() As Variant) As String
    mBuilder.Length = 0
    Call mBuilder.InternalAppendFormat(provider, fmt, args)
    FormatArrayEx = mBuilder.ToString
End Function

''
' Creates and array of chars (Integers) from the specified string.
'
' @param s The string to create the chars from.
' @param startindex The start index in the string to begin converting to a char array. This is zero-based.
' @param length The number of characters to convert to chars.
' @return An array containing the converted characters from the string.
'
Public Function ToCharArray(ByRef s As String, Optional ByRef StartIndex As Variant, Optional ByRef Length As Variant) As Integer()
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim result      As Long
    result = GetOptionalListRange(Len(s), StartIndex, ElemIndex, Length, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Length", IsMissing(StartIndex))
    
    Dim ret() As Integer
    If ElemCount > 0 Then
        ReDim ret(0 To ElemCount - 1)
        Call CopyMemory(ret(0), ByVal StrPtr(s) + ElemIndex * 2, ElemCount * 2)
    Else
        ret = cArray.CreateInstance(ciInteger)
    End If
    ToCharArray = ret
End Function

''
' Pads a string's left side.
'
' @param s The string to be padded.
' @param totalWidth The total length of the final string.
' @param paddingChar The character to pad the left of the string with.
' @return The padded string.
' @remarks The paddingChar parameter can accept either a number or chr$ value.
'
Public Function PadLeft(ByRef s As String, ByVal totalWidth As Long, Optional ByRef paddingChar As Variant = 32) As String
    Dim l As Long
    l = Len(s)
    If totalWidth > l Then
        PadLeft = String$(totalWidth - l, paddingChar) & s
    Else
        PadLeft = s
    End If
End Function

''
' Pads a string's right side.
'
' @param s The string to be padded.
' @param totalWidth The total length of the final string.
' @param paddingChar The character to pad the right of the string with.
' @return The padded string.
' @remarks The paddingChar parameter can accept either a number or chr$ value.
'
Public Function PadRight(ByRef s As String, ByVal totalWidth As Long, Optional ByRef paddingChar As Variant = 32) As String
    Dim l As Long
    l = Len(s)
    If totalWidth > l Then
        PadRight = s & String$(totalWidth - l, paddingChar)
    Else
        PadRight = s
    End If
End Function

''
' Trims the beginning of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's LTrim function, this can trim multiple character types,
' not just spaces.
' <p>The <i>trimChars</i> can be either a string of the characters or an integer array
' of characters that will be trimmed from the beginning.</p>
'
Public Function TrimStart(ByRef s As String, Optional ByRef trimChars As Variant) As String
    Dim Chars() As Integer
    Dim StartIndex As Long
    
    Select Case VarType(trimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(trimChars)
            mTrimChars.SA.cElements = Len(trimChars)
            StartIndex = TrimStartIndex(s, mTrimChars.Data)
            
        Case INTEGER_ARRAY
            SAPtr(Chars) = GetValidArrayPointer(trimChars)
            StartIndex = TrimStartIndex(s, Chars)
            SAPtr(Chars) = 0
        
        Case vbError
            StartIndex = SzTrimStartIndex(s)
        
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
    
    If StartIndex = 0 Then Exit Function
    TrimStart = Mid$(s, StartIndex)
End Function

''
' Trims the end of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's RTrim function, this can trim multiple character types,
' not just spaces.
' <p>The <i>trimChars</i> can be either a string of the characters or an integer array
' of characters that will be trimmed from the end.</p>
'
Public Function TrimEnd(ByRef s As String, Optional ByRef trimChars As Variant) As String
    Dim Chars() As Integer
    Dim EndIndex As Long
    
    Select Case VarType(trimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(trimChars)
            mTrimChars.SA.cElements = Len(trimChars)
            EndIndex = TrimEndIndex(s, mTrimChars.Data)
            
        Case INTEGER_ARRAY
            SAPtr(Chars) = GetValidArrayPointer(trimChars)
            EndIndex = TrimEndIndex(s, Chars)
            SAPtr(Chars) = 0
        
        Case vbError
            EndIndex = SzTrimEndIndex(s)
        
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
    
    If EndIndex = 0 Then Exit Function
    TrimEnd = Left$(s, EndIndex)
End Function

''
' Trims both the beginning and end of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's Trim function, this can trim multiple character types,
' not just spaces.
' <p>The <i>trimChars</i> can be either a string of the characters or an integer array
' of characters that will be trimmed from both ends.</p>
'
Public Function Trim(ByRef s As String, Optional ByRef trimChars As Variant) As String
    Dim Chars() As Integer
    Dim StartIndex As Long
    Dim EndIndex As Long
    
    Select Case VarType(trimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(trimChars)
            mTrimChars.SA.cElements = Len(trimChars)
            StartIndex = TrimStartIndex(s, mTrimChars.Data)
            EndIndex = TrimEndIndex(s, mTrimChars.Data)
            
        Case INTEGER_ARRAY
            SAPtr(Chars) = GetValidArrayPointer(trimChars)
            StartIndex = TrimStartIndex(s, Chars)
            EndIndex = TrimEndIndex(s, Chars)
            SAPtr(Chars) = 0
            
        Case vbError    ' assume missing
            StartIndex = SzTrimStartIndex(s)
            EndIndex = SzTrimEndIndex(s)
        
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
    
    If StartIndex = 0 Or EndIndex = 0 Then Exit Function
    Trim = Mid$(s, StartIndex, EndIndex - StartIndex + 1)
End Function

''
' Joins an array or subarray of strings into a single string, separated by the <i>Delimiter</i>.
'
' @param Arr An array of strings or variants to join together.
' @param Delimiter A string inserted between each element when joining into a string.
' @param Index The starting element in <i>Arr</i> to begin joining from.
' @param Count The number of strings elements to join together.
' @return A string containing elements joined from the array separated by the delimiter.
'
Public Function Join(ByRef Arr As Variant, Optional ByRef Delimiter As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim pSA As Long
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim lb As Long
    Dim ub As Long
    Dim ElemSize As Long
    
    Select Case VarType(Arr)
        Case STRING_ARRAY:      ElemSize = 4
        Case VARIANT_ARRAY:     ElemSize = 16
        Case Else
            Throw Cor.NewArgumentException("A String or Variant array is required.", "Arr")
    End Select
    
    pSA = GetValidArrayPointer(Arr)
    lb = LBound(Arr)
    ub = UBound(Arr)
    
    ElemIndex = GetOptionalLong(Index, lb)
    ElemCount = GetOptionalLong(Count, ub - ElemIndex + 1)
    
    Dim result As Long
    result = VerifyArrayRange(pSA, ElemIndex, ElemCount)
    If result <> NO_ERROR Then Call ThrowArrayRangeException(result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    With mJoinSA
        .cElements = ElemCount
        .fFeatures = MemWord(pSA + FFEATURES_OFFSET)
        .pvData = MemLong(pSA + PVDATA_OFFSET) + (ElemIndex - lb) * ElemSize
    End With
    Join = VBA.Join(mJoinArray, Delimiter)
End Function

''
' Removes a substring from the string, returning a new string with the gap filled
' with the remaining characters.
'
' @param s The string to remove characters from.
' @param StartIndex The index of the first character to begin removing. This index is zero-based.
' @param Count The number of characters to be removed.
' @return A new string with the substring removed.
'
Public Function Remove(ByRef s As String, ByVal StartIndex As Long, ByVal Count As Long) As String
    Dim CharCount As Long
    CharCount = Len(s)
    
    Dim result As Long
    result = VerifyListRange(CharCount, StartIndex, Count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, StartIndex, "StartIndex", Count, "Count")
    
    Dim sp As Long
    sp = StrPtr(s)
    Remove = SysAllocStringLen(sp, CharCount - Count)
    Call CopyMemory(ByVal StrPtr(Remove) + StartIndex * 2, ByVal sp + (StartIndex + Count) * 2, (CharCount - (StartIndex + Count)) * 2)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TrimStartIndex(ByRef s As String, ByRef trimChars() As Integer) As Long
    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim ch As Integer
    Dim pos As Long
    Dim l As Long
    
    lb = LBound(trimChars)
    ub = UBound(trimChars)
    l = Len(s)
    With mString
        .SA.pvData = StrPtr(s)
        Do While pos < l
            ch = .Data(pos)
            For i = lb To ub
                If ch = trimChars(i) Then Exit For
            Next i
            If i > ub Then Exit Do
            pos = pos + 1
        Loop
        If pos = l Then Exit Function
        TrimStartIndex = pos + 1
    End With
End Function

' This method is optimized for the default set of whitespace.
Private Function SzTrimStartIndex(ByRef s As String) As Long
    Dim i As Long
    
    With mString
        .SA.pvData = StrPtr(s)
        For i = 0 To Len(s) - 1
            Select Case .Data(i)
                Case &H20, &HD, &H9, &HA, &HB, &HC, &HA0, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H200B, &H3000, &HFEFF
                Case Else
                    SzTrimStartIndex = i + 1
                    Exit Function
            End Select
        Next i
    End With
End Function

Private Function TrimEndIndex(ByRef s As String, ByRef trimChars() As Integer) As Long
    Dim i As Long
    Dim lb As Long
    Dim ub As Long
    Dim ch As Integer
    Dim pos As Long
    
    lb = LBound(trimChars)
    ub = UBound(trimChars)
    pos = Len(s) - 1
    With mString
        .SA.pvData = StrPtr(s)
        Do While pos >= 0
            ch = .Data(pos)
            For i = lb To ub
                If ch = trimChars(i) Then Exit For
            Next i
            If i > ub Then Exit Do
            pos = pos - 1
        Loop
        If pos < 0 Then Exit Function
        TrimEndIndex = pos + 1
    End With
End Function

' This method is optimized for the default set of whitespace.
Private Function SzTrimEndIndex(ByRef s As String) As Long
    Dim i As Long
    
    With mString
        .SA.pvData = StrPtr(s)
        For i = Len(s) - 1 To 0 Step -1
            Select Case .Data(i)
                Case &H20, &HD, &H9, &HA, &HB, &HC, &HA0, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H200B, &H3000, &HFEFF
                Case Else
                    SzTrimEndIndex = i + 1
                    Exit Function
            End Select
        Next i
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBuilder = New StringBuilder
    Call InitWordBuffer(mString, 0, &H7FFFFFFF)
    Call InitWordBuffer(mTrimChars, 0, 0)
    
    With mJoinSA
        .cbElements = 4
        .cDims = 1
    End With
    SAPtr(mJoinArray) = VarPtr(mJoinSA)
End Sub

Private Sub Class_Terminate()
    SAPtr(mJoinArray) = 0
End Sub
