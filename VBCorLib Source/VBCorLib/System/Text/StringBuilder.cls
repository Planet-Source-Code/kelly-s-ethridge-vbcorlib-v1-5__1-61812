VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: StringBuilder
'

''
' Allows easy manipulation, concatenation and removal of strings and characters.
'
' @remarks
' This class can be used to quickly minpulate strings. It allows for easy concatenation,
' removal and modification of the current underlying string value.
'
' @see Constructors
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Option Explicit
Implements IObject

Private Const DEF_FILLCHAR          As Integer = 32
Private Const DEF_CAPACITY          As Long = 512
Private Const COLON                 As Long = 58
Private Const OPEN_BRACE            As Long = 123
Private Const CLOSE_BRACE           As Long = 125
Private Const COMMA                 As Long = 44
Private Const MINUS                 As Long = 45
Private Const CHAR_COMMA            As Long = 44
Private Const CHAR_COLON            As Long = 58

' module wide scope
Private mStringPtr  As Long
Private mStringSA   As SafeArray1d
Private mString()   As Integer
Private mCapacity   As Long
Private mLength     As Long
Private mFillChar   As Integer

' used by Replace
Private OldStr()    As Integer
Private BufStr()    As Integer
Private OldStrSA    As SafeArray1d
Private BufStrSA    As SafeArray1d

' used by AppendFormat
Private mChars      As WordBuffer



''
' Appends a string to the current string value.
'
' @param Value Value to be appended.
' @param StartIndex A zero-based index of the first character to start appending from in the string.
' @param Count The number of characters from the string to append.
' @return This instance of StringBuilder.
' @remarks This version of Append attempts to convert to the value using the
' <b>Convert.ToString</b> method. This will allow for objects that impelement
' <b>IObject</b> or <b>IFormattable</b> to be converted to a string value.
'
Public Function Append(ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As StringBuilder
    Set Append = AppendString(Convert.ToString(Value), StartIndex, Count)
End Function

''
' Appends a character a specified number of times to the end of string value.
'
' @param char The AscW character code to append.
' @param count The number of times to append the character.
' @return This instance of StringBuilder.
' @remarks The character is a unicode value from -32767 to 65535. For values that
' are greater than 32767 (&H7FFF), 65536 (&H10000) is subtracted from it to map the
' character to a 16-bit Integer without overflowing.
'
Public Function AppendChar(ByVal Char As Long, Optional ByVal Count As Long = 1) As StringBuilder
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count cannot be negative.", "Count", Count)
    If Char < -&H7FFF& Or Char > &HFFFF& Then _
        Throw Cor.NewArgumentOutOfRangeException("Char must be between -32767 and 65535 inclusively.", "Char", Char)
    
    If Char > &H7FFF& Then Char = Char - &H10000
    
    If mLength + Count > mCapacity Then Call EnsureCapacity(mLength + Count)
    
    ' appending 1 character at a time happens 99%
    ' of the time, so optimize for it.
    If Count = 1 Then
        mString(mLength) = Char
    Else
        Call Fill(mLength, Char, Count)
    End If
    mLength = mLength + Count
    Set AppendChar = Me
End Function

''
' Appends an array of characters to the end of the string value.
'
' @param Chars The array of characters to be appended.
' @param startindex The index from Chars to start appending.
' @param count The number of characters to append.
' @return This instance of StringBuilder.
'
Public Function AppendChars(ByRef Chars() As Integer, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As StringBuilder
    Dim ElemCount As Long
    Dim ElemIndex As Long
    Dim result As Long
    
    result = GetOptionalArrayRange(SAPtr(Chars), StartIndex, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowArrayRangeException(result, "Chars", ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))
    
    If ElemCount = 0 Then Exit Function
    
    If mLength + ElemCount > mCapacity Then Call EnsureCapacity(mLength + ElemCount)
    Call CopyMemory(mString(mLength), Chars(ElemIndex), ElemCount * 2)
    mLength = mLength + ElemCount
    Set AppendChars = Me
End Function

''
' Appends a string with formatted arguments.
'
' @param format The string to append that contains formatting information.
' @param args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @remarks
' <p>The format string can embed indexes into the args parameter to indicate where
' a value should be placed. The way to indicate an index is through an argument index enclosed
' in braces {N}. N is the parameter index beginning with zero.</p>
' <br><p>Some arguments may handle additional formatting commands. Please refer to their
' documentation for more information.</p>
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\..\Includes\StringBuilder.AppendFormat.txt"
Public Function AppendFormat(ByRef Format As String, ParamArray args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16))
    Call InternalAppendFormat(Nothing, Format, vArgs)
    Set AppendFormat = Me
End Function

''
' Appends a string with formatted arguments using the supplied provider.
'
' @param provider A custom formatting object that formats the arguments.
' @param Format The string to append that contains formatting information.
' @param args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @remarks
' <p>The format string can embed indexes into the args parameter to indicate where
' a value should be placed. The way to indicate an index is through an argument index enclosed
' in braces {N}. N is the parameter index beginning with zero.</p>
' <br><p>Some arguments may handle additional formatting commands. Please refer to their
' documentation for more information.</p>
' <br><p>The supplied provider can format specific arguments unknown to the system. It can also
' delegate formatting of system-known arguments to another StringBuilder (usually cString.Format)
' to provide default formatting of known types and format only specific types.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Public Function AppendFormatEx(ByVal provider As IFormatProvider, ByRef Format As String, ParamArray args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(20))
    Call InternalAppendFormat(provider, Format, vArgs)
    Set AppendFormatEx = Me
End Function

''
' Appends a string to the current string value.
'
' @param s String to be appended.
' @param StartIndex A zero-based index of the first character to start appending from in the string.
' @param Count The number of characters from the string to append.
' @return A reference to this instance of StringBuilder.
' @remarks<p>This version of append is designed for fast concatentation
' of the supplied string value. The entire string is appended.
' <p>This function does not return a reference to the instance of StringBuilder
' in order to keep the overhead cost down on the function call.</p>
'
Public Function AppendString(ByRef s As String, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As StringBuilder
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim result      As Long
    result = GetOptionalListRange(Len(s), StartIndex, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))

    If mLength + ElemCount > mCapacity Then Call EnsureCapacity(mLength + ElemCount)
    Call CopyMemory(ByVal mStringPtr + mLength * 2, ByVal StrPtr(s) + ElemIndex * 2, ElemCount * 2)
    mLength = mLength + ElemCount
    Set AppendString = Me
End Function

''
' Appends the supplied string and a Carriage-Return/Linefeed to the end of the StringBuilder.
'
' @param s The string to append to the builder, followed by a Carraige-Return/Linefeed.
' @return A reference to this instance of StringBuilder.
'
Public Function AppendLine(Optional ByVal s As String) As StringBuilder
    Call AppendQuick(s & vbCrLf)
    Set AppendLine = Me
End Function

''
' This function is designed to append strings quickly.
'
' @param s The string value to be appended.
' @remarks This method appends the entire string to the end of the
' StringBuilder. There is no return reference to the StringBuilder.
'
Public Sub AppendQuick(ByRef s As String)
    Dim ElemCount As Long
    ElemCount = Len(s)
    If ElemCount = 0 Then Exit Sub
    If mLength + ElemCount > mCapacity Then Call EnsureCapacity(mLength + ElemCount)
    Call CopyMemory(ByVal mStringPtr + mLength * 2, ByVal StrPtr(s), ElemCount * 2)
    mLength = mLength + ElemCount
End Sub

''
' Copies a set of characters from the StringBuilder into a character array.
'
' @param sourceindex The starting index in the StringBuilder to be copying from (0-based).
' @param Destination The character array to copy to.
' @param destinationindex The start index to begin placing characters in to.
' @param Count The number of characters to copy.
'
Public Sub CopyTo(ByVal SourceIndex As Long, ByRef Destination() As Integer, ByVal DestinationIndex As Long, ByVal Count As Long)
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Destination), DestinationIndex, Count)
    If result <> NO_ERROR Then _
        ThrowArrayRangeException result, "Destination", DestinationIndex, "DestinationIndex", Count, "Count"
    
    result = VerifyListRange(mLength, SourceIndex, Count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, SourceIndex, "SourceIndex", Count, "Count")
    
    If Count > 0 Then Call CopyMemory(Destination(DestinationIndex), ByVal mStringPtr + SourceIndex * 2, Count * 2)
End Sub

''
' Returns the buffer size of the string value.
'
' @return Value indicating the amount of memory allocated in characters.
'
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the amount of memory allocated in characters.
'
' @param RHS The number of characters to allocate.
'
Public Property Let Capacity(ByVal RHS As Long)
    If RHS < mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Cannot set capacity less than the length of the current instance.", "Capacity", RHS)
    
    Dim Ptr As Long
    Ptr = CoTaskMemAlloc(RHS * 2)
    If Ptr = 0 Then Throw New OutOfMemoryException
    Call CopyMemory(ByVal Ptr, ByVal mStringPtr, mLength * 2)
    Call CoTaskMemFree(mStringPtr)
    mStringPtr = Ptr
    mStringSA.pvData = Ptr
    mCapacity = RHS
End Property

''
' Returns a character from the string value.
'
' @param Index The index from the string value to return the character.
' @return The character in the string at the specified index.
' @remarks Index is zero-base.
'
Public Property Get Chars(ByVal Index As Long) As Integer
Attribute Chars.VB_UserMemId = 0
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    Chars = mString(Index)
End Property

''
' Sets a character in the string value.
'
' @param index The index into the string value to set the character.
' @param char The character to set.
' @remarks index is zero-based.
'
Public Property Let Chars(ByVal Index As Long, ByVal Char As Integer)
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", Index)
    
    mString(Index) = Char
End Property

''
' Ensures that the internal buffer has atleast the requested amount.
'
' @param requiredCapacity The minimum amount of characters to be allocated.
' @return The capacity after the call is complete.
'
Public Function EnsureCapacity(ByVal RequiredCapacity As Long) As Long
    If RequiredCapacity <= mCapacity Then
        EnsureCapacity = mCapacity
        Exit Function
    End If
    
    Dim NewCapacity As Long
    NewCapacity = mCapacity * 2
    If RequiredCapacity > NewCapacity Then NewCapacity = RequiredCapacity
    Capacity = NewCapacity
    EnsureCapacity = NewCapacity
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns the fill character used to pad empty space.
'
' @return An Integer representing the AscW of the character.
'
Public Property Get FillCharacter() As Variant
    FillCharacter = mFillChar
End Property

''
' Sets the fill character used to pad empty space.
'
' @param RHS The character as either a String or Integer value.
' @remarks The new value can be either String, in which case the
' first characters is used, or it can be a number.
'
Public Property Let FillCharacter(ByVal RHS As Variant)
    Select Case VarType(RHS)
        Case vbLong, vbInteger, vbByte
            mFillChar = AscW(ChrW$(RHS))
        Case vbString
            mFillChar = AscW(RHS)
        Case Else
            Throw Cor.NewInvalidCastException("Fill Character can be either a string or number.")
    End Select
End Property

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function

''
' Inserts a string into the current string value.
'
' @param index The index in the string value to start inserting the new string.
' @param value The string to be inserted.
' @param count The number of times to insert the string.
' @return This instance of StringBuilder
'
Public Function Insert(ByVal Index As Long, ByVal Value As String, Optional ByVal Count As Long = 1) As StringBuilder
    If Index < 0 Or Index > mLength Then _
        Throw Cor.NewIndexOutOfRangeException("index must be non-negative and less than or equal to the length.")
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
        
    Dim l As Long
    l = Len(Value)
    If l > 0 And Count > 0 Then
        mLength = InsertSpace(Index, l * Count)
        l = l * 2
        Index = Index * 2
        Do While Count > 0
            Call CopyMemory(ByVal mStringPtr + Index, ByVal StrPtr(Value), l)
            Index = Index + l
            Count = Count - 1
        Loop
    End If
    Set Insert = Me
End Function

''
' Inserts an array of characters into the string value.
'
' @param index The index in the string value to start inserting.
' @param Chars The characters to be inserted.
' @param startindex The index in the Characters to start inserting from.
' @param count The number of characters to be inserted.
' @return This instance of StringBuilder.
'
Public Function InsertChars(ByVal Index As Long, ByRef Chars() As Integer, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As StringBuilder
    Dim ElemCount As Long
    Dim ElemIndex As Long
    Dim result As Long
    
    result = GetOptionalArrayRange(SAPtr(Chars), StartIndex, ElemIndex, Count, ElemCount)
    If result <> NO_ERROR Then _
        ThrowArrayRangeException result, "Chars", ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex)
    If Index < 0 Or Index > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Index must be between 0 and then length of the current instance.", "Index", Index)
    
    mLength = InsertSpace(Index, ElemCount)
    Call CopyMemory(mString(Index), Chars(ElemIndex), ElemCount * 2)
    Set InsertChars = Me
End Function

''
' Returns the length of the current string value.
'
' @return The length of the string value.
'
Public Property Get Length() As Long
    Length = mLength
End Property

''
' Sets the length of the current string value.
'
' @param RHS The length of the string value.
' @remarks If the new value is longer than the current length, then the
' new space is padded with FillCharacter (default 32).
'
Public Property Let Length(ByVal RHS As Long)
    If RHS < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Length cannot be set less than 0.", "Length", Length)
    
    If RHS > mLength Then
        If RHS > mCapacity Then Call EnsureCapacity(RHS)
        Call Fill(mLength, mFillChar, mCapacity - mLength)
    End If
    mLength = RHS
End Property

''
' Removes a number of characters from the string value.
'
' @param startindex The index in the string value to start removing from.
' @param count The number of characters to remove from the string value.
' @return This instance of StringBuilder.
' @remarks startindex is zero-based.
'
Public Function Remove(ByVal StartIndex As Long, ByVal Count As Long) As StringBuilder
    Dim result As Long
    result = VerifyListRange(mLength, StartIndex, Count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, StartIndex, "StartIndex", Count, "Count")
    
    If Count > 0 Then
        Call CopyMemory(ByVal mStringPtr + StartIndex * 2, ByVal mStringPtr + (StartIndex + Count) * 2, (mLength - (StartIndex + Count)) * 2)
        mLength = mLength - Count
    End If
    Set Remove = Me
End Function

''
' Replaces a substring in the string value with a new string.
'
' @param oldValue The substring to be replaced.
' @param newValue The string to replace the old string with.
' @param index The index of the start of the substring in the new string value.
' @param count The number of characters in the new string to include in the replacement.
' @return This instance of StringBuilder.
' @remarks Use index and count to replace the old value with a substring of the new value. index is zero-based.
'
Public Function Replace(ByRef oldValue As String, ByVal newValue As String, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As StringBuilder
    Dim oldLen As Long
    oldLen = Len(oldValue)
    If oldLen = 0 Then Exit Function
    
    OldStrSA.pvData = StrPtr(oldValue)
    
    Dim i As Long
    Dim j As Long
    Dim foundIndex As Long
    Dim occ As Long
    Dim occurrences() As Long
    
    ReDim occurrences(mLength - 1)
    foundIndex = -1
    For i = 0 To mLength - oldLen + 1
        If mString(i) = OldStr(j) Then
            j = j + 1
            If foundIndex = -1 Then foundIndex = i
            If j = oldLen Then
                occurrences(occ) = foundIndex
                occ = occ + 1
                j = 0
                foundIndex = -1
            End If
        Else
            foundIndex = -1
            j = 0
        End If
    Next i
    
    If occ > 0 Then
        Dim newLen As Long
        newLen = Len(newValue)
        
        Dim ElemCount As Long
        Dim ElemIndex As Long
        If GetOptionalLongPair(Index, 0, ElemIndex, Count, newLen, ElemCount) = Argument_ParamRequired Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ParamRequired), IIf(IsMissing(Index), "Index", "Count"))
        
        If Not IsMissing(Index) Then
            newValue = Mid$(newValue, ElemIndex + 1, ElemCount)
            newLen = Len(newValue)
        End If
        
        Dim NewPtr As Long
        NewPtr = StrPtr(newValue)
        
        mCapacity = mLength + newLen * occ
        BufStrSA.pvData = CoTaskMemAlloc(mCapacity * 2)
        
        Dim k As Long
        Dim c As Long
        
        j = 0: i = 0
        Do While c < occ
            k = occurrences(c) - i
            If k > 0 Then
                Call CopyMemory(ByVal BufStrSA.pvData + j * 2, ByVal mStringSA.pvData + i * 2, k * 2)
                j = j + k
                i = i + k
            End If
            If newLen > 0 Then
                Call CopyMemory(ByVal BufStrSA.pvData + j * 2, ByVal NewPtr, newLen * 2)
                j = j + newLen
            End If
            i = i + oldLen
            c = c + 1
        Loop
        If i < mLength Then
            Call CopyMemory(ByVal BufStrSA.pvData + j * 2, ByVal mStringSA.pvData + i * 2, (mLength - i) * 2)
        End If
        
        Call CoTaskMemFree(mStringPtr)
        mStringPtr = BufStrSA.pvData
        mStringSA.pvData = BufStrSA.pvData
        mLength = mLength + (newLen - oldLen) * occ
    End If

    Set Replace = Me
End Function

''
' Returns the current version of the string value.
'
' @param startindex The index to start the return substring from.
' @param length The number of characters to return in the string.
' @return A string or substring representing the internal string of the builder.
' @remarks startindex is zero-based.
'
Public Function ToString(Optional ByRef StartIndex As Variant, Optional ByRef Length As Variant) As String
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mLength, StartIndex, ElemIndex, Length, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Length", IsMissing(StartIndex))
    
    If mLength = 0 Then
        ToString = ""
    Else
        ToString = SysAllocStringLen(mStringPtr + ElemIndex * 2, ElemCount)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByRef StartIndex As Variant, ByRef Count As Variant, ByVal Capacity As Long)
    Call CoTaskMemFree(mStringPtr)
    
    Dim ElemCount As Long
    ElemCount = GetOptionalLong(Count, Len(s))
    
    Dim ElemIndex As Long
    ElemIndex = GetOptionalLong(StartIndex, 0)
    
    Dim result As Long
    result = VerifyListRange(Len(s), ElemIndex, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))
    
    If ElemCount > Capacity Then Capacity = ElemCount
    If Capacity < DEF_CAPACITY Then Capacity = DEF_CAPACITY
    
    mStringPtr = CoTaskMemAlloc(Capacity * 2)
    mCapacity = Capacity
    mLength = ElemCount
    If ElemCount > 0 Then Call CopyMemory(ByVal mStringPtr, ByVal StrPtr(s) + ElemIndex * 2, ElemCount * 2)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mFillChar = DEF_FILLCHAR
    mStringPtr = CoTaskMemAlloc(DEF_CAPACITY * 2)
    mCapacity = DEF_CAPACITY
    If mStringPtr = 0 Then Throw New OutOfMemoryException
    
    With OldStrSA
        .cbElements = 2
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    SAPtr(OldStr) = VarPtr(OldStrSA)
    BufStrSA = OldStrSA
    SAPtr(BufStr) = VarPtr(BufStrSA)
    mStringSA = OldStrSA
    SAPtr(mString) = VarPtr(mStringSA)
    mStringSA.pvData = mStringPtr
    
    Call InitWordBuffer(mChars, 0, &H7FFFFFFF)
End Sub

Private Sub Class_Terminate()
    Call CoTaskMemFree(mStringPtr)
    SAPtr(OldStr) = 0
    SAPtr(BufStr) = 0
    SAPtr(mString) = 0
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Fill(ByVal Index As Long, ByVal Char As Integer, ByVal Count As Long)
    Dim i As Long
    
    For i = Index To Index + Count - 1
        mString(i) = Char
    Next i
End Sub

Private Sub FormatError()
    Throw Cor.NewFormatException("Invalid format specified.")
End Sub

Private Function InsertSpace(ByVal Index As Long, ByVal Size As Long) As Long
    InsertSpace = mLength + Size
    If InsertSpace > mCapacity Then Call EnsureCapacity(InsertSpace)
    If Index < mLength Then
        ' Create space in the string. The new space is uninitialized.
        Call CopyMemory(ByVal mStringPtr + (Index + Size) * 2, ByVal mStringPtr + Index * 2, (mLength - Index) * 2)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend helper methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' We can supply a custom formatter through the IFormatProvider interface.
' This article lead me to believe that the StringBuilder.AppendFormat function
' can use an ICustomFormatter object, even though information about it is
' scarce in the MSDN.
'
' "http://www.codeproject.com/csharp/custstrformat.asp"
'
Friend Sub InternalAppendFormat(ByVal provider As IFormatProvider, ByRef Format As String, ByRef args() As Variant)
    Dim CustomFormatter As ICustomFormatter
    Dim Index As Long
    Dim Max As Long
    Dim SubStringStart As Long
    
    ' See if the user wants to format anything their own way.
    If Not provider Is Nothing Then
        Set CustomFormatter = provider.GetFormat("ICustomFormatter")
    End If
    
    Max = Len(Format)
    mChars.SA.pvData = StrPtr(Format)
    mChars.SA.cElements = Max
    
    ' Parse normal substring.
    Do While Index < Max
        Select Case mChars.Data(Index)
            Case OPEN_BRACE
                ' potential escape or format info.
                
                If Index > SubStringStart Then
                    ' If we have some characters, just append them now
                    ' even if the "{" may be escaped. Most times it won't be.
                    Call AppendChars(mChars.Data, SubStringStart, Index - SubStringStart)
                End If
                
                Index = Index + 1
                
                ' We test for escape by checking the next character for a "{".
                ' But if we're out of characters, then we ended with a "{",
                ' and that is an invalid format.
                If Index = Max Then Call FormatError
                
                ' We have enough characters to test for possible escape.
                If mChars.Data(Index) = OPEN_BRACE Then
                    ' escape it by setting the start of the substring to the second one.
                    SubStringStart = Index
                Else
                    Dim ArgIndex As Long
                    Dim ArgWidth As Long
                    Dim ArgFormatting As String
                    
                    ' must be format info within a "{}", so call the parser for it.
                    Index = ParseArgumentFormatting(Index, ArgIndex, ArgWidth, ArgFormatting)
                    If ArgIndex < 0 Or ArgIndex > UBound(args) Then _
                        Throw Cor.NewFormatException("The argument index must be from 0 to " & UBound(args) & ".")
                        
                    ' append formatted argument here.
                    Dim s As String
                    
                    ' Set this because we are in a loop and always
                    ' need to begin with a vbNullString.
                    s = vbNullString
                    
                    ' Give a custom formatter first shot at formatting.
                    If Not CustomFormatter Is Nothing Then
                        s = CustomFormatter.Format(ArgFormatting, args(ArgIndex), provider)
                    End If
                    
                    ' We distinguish between a null string and an empty string.
                    ' Null string indicates that the value was not formatted.
                    If StrPtr(s) = 0 Then
                        s = Convert.ToString(args(ArgIndex), ArgFormatting, provider)
                    End If
                    
                    ' Add leading padding.
                    If ArgWidth > Len(s) Then
                        Call AppendChar(CHAR_SPACE, ArgWidth - Len(s))
                    End If
                    
                    ' Finally append the formatted value to the string.
                    Call AppendQuick(s)
                    
                    ' Add trailing padding.
                    If -ArgWidth > Len(s) Then
                        Call AppendChar(CHAR_SPACE, (-ArgWidth) - Len(s))
                    End If
                    
                    SubStringStart = Index + 1
                    
                End If
            Case CLOSE_BRACE
                ' must be escaped or else it is a format error.
                
                Index = Index + 1
                
                ' Need enough characters to test for escape.
                If Index = Max Then Call FormatError
                
                If mChars.Data(Index) = CLOSE_BRACE Then
                    ' escape it and add the substring to this point.
                    ' The substring ends one character back because we
                    ' don't want to include both "}}" in the substring.
                    If Index - 1 > SubStringStart Then
                        ' Append the current substring.
                        Call AppendChars(mChars.Data, SubStringStart, Index - SubStringStart - 1)
                    End If
                    SubStringStart = Index
                Else
                    ' we failed to escape it, so there was no matching "{".
                    Call FormatError
                End If
        End Select
        Index = Index + 1
    Loop
    
    If Index > SubStringStart Then
        ' If we have some characters, just append them now.
        Call AppendChars(mChars.Data, SubStringStart, Index - SubStringStart)
    End If
End Sub

' We must be inside "{}". We must end with a matching "}" or it is a format error.
'
' @param Index The index of the first character after the "{".
' @param ArgIndex The index of the argument in the Args array.
' @param ArgWidth The width of the column the argument will be displayed in.
' @param ArgFormat The format information used to determine formatting for the argument.
' @return The index of the terminating "}".
'
Private Function ParseArgumentFormatting(ByVal Index As Long, ByRef ArgIndex As Long, ByRef ArgWidth As Long, ByRef ArgFormat As String) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    ' Get the argument index.
    ArgIndex = ParseNumber(Index)
    
    Index = ConsumeSpaces(Index)
    
    With mChars
        ' Check if a column size is being defined.
        If .Data(Index) = CHAR_COMMA Then
            Index = ConsumeSpaces(Index + 1)
            
            ' Get the column width and aligment. Remember
            ' negative values mean align left in the column.
            ArgWidth = ParseNumber(Index)
            
            Index = ConsumeSpaces(Index)
        Else
            ' Make sure they know there was no width specified.
            ArgWidth = 0
        End If
        
        ' Check if formatting is being specified.
        If .Data(Index) = CHAR_COLON Then
            Dim StartIndex As Long
            
            ' everything from here to the "}" is part of the argument formatting.
            Index = Index + 1
            StartIndex = Index
            Do
                ' can never reach the end until after a "}".
                If Index = Max Then FormatError
                If .Data(Index) = CLOSE_BRACE Then
                    ArgFormat = SysAllocStringLen(VarPtr(.Data(StartIndex)), Index - StartIndex)
                    Exit Do
                End If
                Index = Index + 1
            Loop
        Else
            ' Make sure we let them know that no formatting was supplied.
            ArgFormat = vbNullString
        End If
        
        ' if we get here we should have reached the closing "}".
        If .Data(Index) <> CLOSE_BRACE Then FormatError
    End With
    ParseArgumentFormatting = Index
End Function

''
' Advances the index to the next non-space character.
'
' @param The index to the first space to be consumed.
' @return The index of the first character after the contiguous
' stream of spaces have been consumed.
'
Private Function ConsumeSpaces(ByVal Index As Long) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    Do
        ' Can never hit the end consuming spaces.
        If Index = Max Then Call FormatError
        If mChars.Data(Index) <> CHAR_SPACE Then Exit Do
        Index = Index + 1
    Loop
    
    ConsumeSpaces = Index
End Function

''
' Parses out a positive or negative number.
'
' @param Index The index to the first digit or the "-" preceeding the first digit.
' @return The number that was parsed.
' @remarks The Index is passed by Reference and will be altered.
'
Private Function ParseNumber(ByRef Index As Long) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    Dim sign As Long
    sign = 1
    
    Dim FirstCharIndex As Long
    FirstCharIndex = Index
    
    Dim Number As Long
    Do
        ' We can never hit the end parsing a number.
        If Index = Max Then Call FormatError
        
        Dim ch As Integer
        ch = mChars.Data(Index)
        Select Case ch
            Case CHAR_0 To CHAR_9
                ' Keep adding digits to the number
                Number = Number * 10 + ch - CHAR_0
            Case CHAR_MINUS
                ' The "-" can only be the first character.
                If Index <> FirstCharIndex Then Call FormatError
                sign = -1
            Case Else
                ' we have reached the end of the number, so exit.
                Exit Do
        End Select
        Index = Index + 1
    Loop
    
    ParseNumber = Number * sign
End Function


