VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamReader"
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
'    Module: StreamReader
'

''
' Reads characters from a byte array in a particular encoding.
'
' @remarks The <b>StreamReader</b> object cannot be created directly. In order to
' create a new <b>StreamReader</b>, use the <b>Cor.NewStreamReader</b> method.
' <p>Unlike the FileStream and MemoryStream classes, the StreamReader
' reads bytes that are of a particular encoding to be decoded. Instead of reading
' bytes, the StreamReader returns characters or strings of the decoded bytes.</p>
' <p>The default encoding is UTF8Encoding.</p>
'
' @see Constructors
' @see StreamWriter
' @see TextReader
' @include "..\..\Includes\StreamReader.txt"
Option Explicit
Implements IObject
Implements TextReader

Private Const MIN_BUFFERSIZE As Long = 128


Private mStream             As Stream
Private mEncoding           As Encoding
Private mDecoder            As Decoder
Private mByteBuffer()       As Byte
Private mByteIndex          As Long
Private mByteLength         As Long
Private mCharBuffer()       As Integer
Private mCharIndex          As Long
Private mCharLength         As Long
Private mBufferSize         As Long
Private mDetermineEncoding  As Boolean
Private mIsOpen             As Boolean
Private mLineBuilder        As New StringBuilder


''
' Returns the base stream this reader is reading from.
'
' @return The base stream.
'
Public Property Get BaseStream() As Stream
    Set BaseStream = mStream
End Property

''
' Returns the current encoding used by the reader.
'
' @return The encoding being used by the reader currently.
'
Public Property Get CurrentEncoding() As Encoding
    Set CurrentEncoding = mEncoding
End Property

''
' Closes the reader.
'
Public Sub CloseReader()
    If mIsOpen Then
        mIsOpen = False
        Call mStream.CloseStream
    End If
End Sub

''
' Discards the currently buffered data to allow reading from
' a new position in the underlying stream.
'
' @remarks The StreamReader buffers more data than is usually
' read at a single time. If the position of the underlying
' stream is moved, then the data buffered in the StreamReader
' needs to be discarded so the reading can begin at the new
' Stream position.
'
Public Sub DiscardBufferedData()
    mByteIndex = 0
    mByteLength = 0
    mCharIndex = 0
    mCharLength = 0
    Set mDecoder = mEncoding.GetDecoder
End Sub

''
' Returns the the next char from the reader without consuming it.
'
' @return The next character to be read, or -1 if the end of the
' stream has been reached.
' @remarks The character that is returned is not removed from the stream.
' Any characters above &H7FFF will be returned as a positive value.
'
Public Function Peek() As Long
    Call VerifyIsOpen
    
    If HaveChars Then
        AsWord(Peek) = mCharBuffer(mCharIndex)
    Else
        Peek = -1
    End If
End Function

''
' Returns the next char from the reader.
'
' @return The next character to be read, or -1 if the end of the stream was reached.
' @remarks Any characters above &H7FFF will be returned as a positive value.
'
Public Function Read() As Long
    Read = Peek
    If Read <> -1 Then mCharIndex = mCharIndex + 1
End Function

''
' Reads a block of characters from the the stream.
'
' @param Buffer The array to read the characters in to.
' @param Index The starting index in <i>Buffer</i> to begin reading to.
' @param Count The number of characters to be read.
' @return The actual number of characters read.
'
Public Function ReadBlock(ByRef Buffer() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Call VerifyIsOpen
    
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Buffer), Index, Count)
    If result <> NO_ERROR Then ThrowArrayRangeException result, "Buffer", Index, "Index", Count, "Count"
    
    Dim AvailableChars As Long
    ' We will loop until we have read the
    ' requested number of characters, if possible.
    Do While Count > 0
        ' If we have no more characters to read then get out.
        If Not HaveChars Then Exit Function
        
        ' Calculate how many characters are available in the local buffer.
        AvailableChars = mCharLength - mCharIndex
        
        ' If there are more characters available than we actaully need,
        ' then trim back the number of characters to what we need.
        If AvailableChars > Count Then AvailableChars = Count
        
        ' Copy the characters we need from the local buffer to the destination buffer.
        Call CopyMemory(Buffer(Index), mCharBuffer(mCharIndex), AvailableChars * 2)
        
        ' Move the destination index out by the number of characters we copied.
        Index = Index + AvailableChars
        
        ' Move the internal index out by the number of characters we copied.
        ' This will be reset to 0 during the "HaveChars" function.
        mCharIndex = mCharIndex + AvailableChars
        
        ' Calculate the amount left to read.
        Count = Count - AvailableChars
        
        ' Calculate the amount read so far.
        ReadBlock = ReadBlock + AvailableChars
    Loop
End Function

''
' Returns a string of characters to the next new-line character.
'
' @return A string containing the characters up to the next new-line character.
' @remarks If there are no more bytes to in the stream, then a null string is returned.
' This can be checked for using the <b>cString.IsNull</b> function.
'
Public Function ReadLine() As String
    Dim Line As String
    Dim StartIndex As Long
    Dim TermChar As Integer
    
    Call VerifyIsOpen
    
    If Not HaveChars Then Exit Function
    
    Dim UsingBuilder As Boolean
    StartIndex = mCharIndex
    Do
        ' We have reached the end of the buffer without
        ' finding a new line separator, so add this buffer
        ' of characters to a StringBuilder for safe keeping.
        If mCharIndex = mCharLength Then
            ' If we aren't using the builder, init it and begin to.
            If Not UsingBuilder Then
                mLineBuilder.Length = 0
                UsingBuilder = True
            End If
            ' place the buffer into the builder and start over with a new full buffer.
            Call mLineBuilder.AppendChars(mCharBuffer, StartIndex, mCharIndex - StartIndex)
            StartIndex = 0
            If ReadBuffer = 0 Then Exit Do  ' there was nothing left to fill the buffer, so return what we have.
        End If
        
        ' Check for either a return or linefeed as new line separators.
        Select Case mCharBuffer(mCharIndex)
            Case CHAR_RETURN, CHAR_LINEFEED
                TermChar = mCharBuffer(mCharIndex)
                
                ' If we aren't using the builder and we found a new line separator
                ' then we can just create the return string from the current buffer.
                If Not UsingBuilder Then
                    Line = SysAllocStringLen(VarPtr(mCharBuffer(StartIndex)), mCharIndex - StartIndex) 'cString.FromCharArray(mCharBuffer, StartIndex, mCharIndex - StartIndex)
                    mCharIndex = mCharIndex + 1
                    Exit Do
                Else
                    ' we've been using the builder because we have reached the end of
                    ' atleast one buffer of characters. So add the subset of characters
                    ' from the current buffer to the builder for returning.
                    Call mLineBuilder.AppendChars(mCharBuffer, StartIndex, mCharIndex - StartIndex)
                    mCharIndex = mCharIndex + 1
                    Exit Do
                End If
        End Select
        mCharIndex = mCharIndex + 1
    Loop
    
    ' We found the entire line in the current buffer, so
    ' we can just return that line.
    If Not UsingBuilder Then
        ' or, might have been the end of Stream , so return a vbNullString.
        ' Use either cString.IsNull or StrPtr to determine this is
        ' what happened.
        If Len(Line) = 0 And mCharLength = 0 Then
            Exit Function
        Else
            ReadLine = Line
        End If
    Else
        ' We've been putting string fragments into a StringBuilder,
        ' so just return the contents.
        ReadLine = mLineBuilder.ToString
    End If
    
    ' We terminate a line with either a Return, LineFeed, or Return-LineFeed.
    ' So check if we had a Return, and if so, determine if the next character
    ' is a LineFeed. Remove them both if possible.
    If TermChar = CHAR_RETURN Then
        ' We still have characters in the buffer, so check if
        ' the next character is a linefeed and remove it if so.
        If mCharIndex < mCharLength Then
            If mCharBuffer(mCharIndex) = CHAR_LINEFEED Then mCharIndex = mCharIndex + 1
        Else
            ' The return was the last character in the buffer,
            ' so fill the buffer with the next set of character.
            ' If atleast 1 character was read, then check for
            ' a linefeed and remove it, if so.
            If ReadBuffer > 0 Then
                If mCharBuffer(0) = CHAR_LINEFEED Then mCharIndex = mCharIndex + 1
            End If
        End If
    End If
End Function

''
' Returns a string of all remaining characters in the stream.
'
' @return A string containing the remaining characters.
'
Public Function ReadToEnd() As String
    Call VerifyIsOpen
    
    If Not HaveChars Then Exit Function
    
    ' The byte buffer wasn't completey full, meaning there is no more to
    ' be read from the stream. So the last of the characters have been
    ' read, so just take them and make a string.
    If mByteLength < mBufferSize Then
        ReadToEnd = cString.FromCharArray(mCharBuffer, mCharIndex, mCharLength)
    Else
        ' Keep reading from the stream and placing the characters into the
        ' builder. Keep refilling the buffer until the end of the stream is reached.
        mLineBuilder.Length = 0
        Do
            Call mLineBuilder.AppendChars(mCharBuffer, mCharIndex, mCharLength - mCharIndex)
        Loop While ReadBuffer > 0
        ReadToEnd = mLineBuilder.ToString
    End If
    mCharIndex = mCharLength
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
' @param value The value to compare equality to.
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
Friend Sub Init(ByRef Source As Variant, ByVal oEncoding As Encoding, ByVal determineEncodingFromByteOrderMarks As Boolean, ByVal BufferSize As Long)
    If BufferSize < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "BufferSize", BufferSize)
    
    ' Get the source from the correct place.
    Select Case VarType(Source)
        Case vbObject
            ' We expect an open Stream object here.
            
            If Source Is Nothing Then _
                Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Stream))
            If Not TypeOf Source Is Stream Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_StreamRequired), "Source")
        
            Set mStream = Source
        
        Case vbString
            ' We expect a filename here.
            Set mStream = Cor.NewFileStream(Source, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare)
        
        Case Else
            Throw Cor.NewArgumentException("A file path or Stream object is required.", "Source")
    End Select
    
    ' Set up the encoding scheme.
    mDetermineEncoding = determineEncodingFromByteOrderMarks
    If Not oEncoding Is Nothing Then
        Set mEncoding = oEncoding
        Set mDecoder = mEncoding.GetDecoder
    End If
    
    ' Set up the read buffer.
    If BufferSize < MIN_BUFFERSIZE Then BufferSize = MIN_BUFFERSIZE
    mBufferSize = BufferSize
    ReDim mByteBuffer(0 To mBufferSize - 1)
    
    mIsOpen = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HaveChars() As Boolean
    ' This will be true when:
    ' - The buffer is empty
    ' - All the characters in the buffer have been read
    If mCharIndex = mCharLength Then
        ' The buffer needs filling.
        HaveChars = (ReadBuffer > 0)
    Else
        ' The buffer already has characters
        HaveChars = True
    End If
End Function

Private Function ReadBuffer() As Long
    ' Fill our local byte buffer from the stream.
    mByteLength = mStream.ReadBlock(mByteBuffer, 0, mBufferSize)
    
    ' If bytes have been read in then fill the character buffer.
    If mByteLength > 0 Then
        ' We may still need to determine encoding from the stream.
        If mDetermineEncoding Then DetermineEncoding
        
        ' We couldn't ReDim the character buffer earlier because we
        ' had no idea the type of encoding we'd end up with, causing
        ' a different number of characters to be generated.
        If SAPtr(mCharBuffer) = vbNullPtr Then ReDim mCharBuffer(0 To mEncoding.GetMaxCharCount(mBufferSize) - 1)
        
        ' Fill the character buffer using a Decoder to maintain
        ' state between buffer readings for multi-byte characters.
        mCharLength = mDecoder.GetChars(mByteBuffer, mByteIndex, mByteLength, mCharBuffer, 0)
    Else
        ' We didn't read anything from the Stream, so we
        ' can't have any characters in the character buffer.
        mCharLength = 0
    End If
    
    mByteIndex = 0
    mCharIndex = 0
    ReadBuffer = mCharLength
End Function

Private Sub DetermineEncoding()
    If mByteLength < 2 Then Exit Sub    ' don't have enough bytes to determine anything.
    
    Select Case AsWord(mByteBuffer(0))
        Case &HFEFF   ' Little-Endian UnicodeEncoding
            Set mEncoding = Cor.NewUnicodeEncoding(False, False)
            mByteIndex = 2
            mByteLength = mByteLength - 2
            
        Case &HFFFE   ' Big-Endian UnicodeEncoding
            Set mEncoding = Cor.NewUnicodeEncoding(True, False)
            mByteIndex = 2
            mByteLength = mByteLength - 2
        
        Case &HBBEF   ' possible UTF-8 encoding
            If mByteLength < 3 Then Exit Sub    ' not quite enough bytes to determine if UTF-8 encoding.
            If mByteBuffer(2) = &HBF Then Set mEncoding = Cor.NewUTF8Encoding(False, False)
            mByteIndex = 3
            mByteLength = mByteLength - 3
    End Select
    
    ' We have either determined the encoding from the BOM, or
    ' we are going to use the default encoding for the reader.
    
    ' Get the Decoder from the encoding now.
    If Not mEncoding Is Nothing Then Set mDecoder = mEncoding.GetDecoder
    mDetermineEncoding = False
End Sub

Private Sub VerifyIsOpen()
    If Not mIsOpen Then Throw Cor.NewObjectDisposedException("StreamReader", "Cannot read from a closed StreamReader.")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mEncoding = Encoding.UTF8
    Set mDecoder = mEncoding.GetDecoder
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
'   TextReader Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextReader_CloseReader()
    Call CloseReader
End Sub

Private Function TextReader_Equals(Value As Variant) As Boolean
    TextReader_Equals = Equals(Value)
End Function

Private Function TextReader_GetHashCode() As Long
    TextReader_GetHashCode = GetHashCode
End Function

Private Function TextReader_Peek() As Long
    TextReader_Peek = Peek
End Function

Private Function TextReader_Read() As Long
    TextReader_Read = Read
End Function

Private Function TextReader_ReadBlock(Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    TextReader_ReadBlock = ReadBlock(Chars, Index, Count)
End Function

Private Function TextReader_ReadLine() As String
    TextReader_ReadLine = ReadLine
End Function

Private Function TextReader_ReadToEnd() As String
    TextReader_ReadToEnd = ReadToEnd
End Function

Private Function TextReader_ToString() As String
    TextReader_ToString = ToString
End Function
