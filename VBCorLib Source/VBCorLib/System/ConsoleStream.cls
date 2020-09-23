VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsoleStream"
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
'    Module: ConsoleStream
'

''
' This is a <b>Stream</b> wrapper around the console input and output methods.
'
Option Explicit
Implements Stream

Private mHandle     As Long
Private mCanRead    As Boolean
Private mCanWrite   As Boolean



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Handle As Long, ByVal Access As FileAccess)
    mHandle = Handle
    mCanRead = (Access And ReadAccess)
    mCanWrite = (Access And WriteAccess)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Stream Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Dim ret As StreamAsyncResult
    
    Set ret = Cor.NewStreamAsyncResult(State)
    On Error GoTo errTrap
    With ret
        .BytesRead = Stream_ReadBlock(Buffer, Offset, Count)
        .IsCompleted = True
        .IsReadType = True
    End With
    
    If Not Callback Is Nothing Then Call Callback.Execute(ret)

errTrap:
    Dim ex As Exception
    If Catch(ex, Err) Then Set ret.Exception = ex
    Set Stream_BeginRead = ret
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Dim ret As StreamAsyncResult
    
    Set ret = Cor.NewStreamAsyncResult(State)
    On Error GoTo errTrap:
    With ret
        .CompletedSynchronously = True
        .IsCompleted = True
    End With
    Call Stream_WriteBlock(Buffer, Offset, Count)
    
    If Not Callback Is Nothing Then Call Callback.Execute(ret)
    
errTrap:
    Dim ex As Exception
    If Catch(ex, Err) Then Set ret.Exception = ex
    Set Stream_BeginWrite = ret
End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = True
End Property

Private Property Get Stream_CanSeek() As Boolean
    Stream_CanSeek = False
End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = True
End Property

Private Sub Stream_CloseStream()
    mCanRead = False
    mCanWrite = False
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Dim result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult object is required.", "AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    Set result = AsyncResult
    If Not result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")
    
    If Not result.Exception Is Nothing Then Throw result.Exception
    
    Stream_EndRead = result.BytesRead
End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    Dim result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult object is required.", "AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    Set result = AsyncResult
    If result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")

    If Not result.Exception Is Nothing Then Throw result.Exception
End Sub

Private Function Stream_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        Stream_Equals = (Value Is Me)
    End If
End Function

Private Sub Stream_Flush()
    ' do nothing
End Sub

Private Function Stream_GetHashCode() As Long
    Stream_GetHashCode = ObjPtr(CUnk(Me))
End Function

Private Property Get Stream_Length() As Currency
    Throw New NotSupportedException
End Property

Private Property Let Stream_Position(ByVal RHS As Currency)
    Throw New NotSupportedException
End Property

Private Property Get Stream_Position() As Currency
    Throw New NotSupportedException
End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Buffer), Offset, Count)
    If result <> NO_ERROR Then _
        Call ThrowArrayRangeException(result, "Buffer", Offset, "Offset", Count, "Count")
    If Not mCanRead Then _
        Throw Cor.NewNotSupportedException("Stream does not support reading")
    
    If ReadFile(mHandle, VarPtr(Buffer(Offset)), Count, Stream_ReadBlock, ByVal 0&) = BOOL_FALSE Then IOError Err.LastDllError
End Function

Private Function Stream_ReadByte() As Long
    Dim b(0) As Byte
    
    If Stream_ReadBlock(b, 0, 1) > 0 Then
        Stream_ReadByte = b(0)
    Else
        Stream_ReadByte = -1
    End If
End Function

Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    Throw New NotSupportedException
End Function

Private Sub Stream_SetLength(ByVal Value As Currency)
    Throw New NotSupportedException
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = Object.ToString(Me, App)
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Buffer), Offset, Count)
    If result <> NO_ERROR Then _
        Call ThrowArrayRangeException(result, "Buffer", Offset, "Offset", Count, "Count")
    If Not mCanWrite Then _
        Throw Cor.NewNotSupportedException("The stream does not support writing.")
    
    If WriteFile(mHandle, VarPtr(Buffer(Offset)), Count, 0, ByVal 0&) = BOOL_FALSE Then IOError Err.LastDllError
End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)
    Dim b(0) As Byte
    b(0) = Value
    Call Stream_WriteBlock(b, 0, 1)
End Sub
