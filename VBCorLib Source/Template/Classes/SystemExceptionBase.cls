VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemExceptionBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This is the base implementation of a SystemException class.
'
'
Option Explicit
Implements IObject
Implements Exception
Implements SystemException


' Default Values
Private Const DEF_HRESULT   As Long = 5   ' Invalid procedure call or argument
Private Const DEF_MESSAGE   As String = vbNullString

' Property names used with the PropertyBag
Private Const PROP_BASE As String = "Base"

' The ExceptionBase class handles everything this class will perform.
Private mBase As VBCorLib.ExceptionBase




''
' Gets a link to a help file associated with the exception.
'
' @return The Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Get HelpLink() As String
    HelpLink = mBase.HelpLink
End Property

''
' Sets a link to a help file associated with the exception.
'
' @param RHS Set the Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Let HelpLink(ByVal RHS As String)
    mBase.HelpLink = RHS
End Property

''
' Gets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @return The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned.
' @see Exception
'
Public Property Get HResult() As Long
    HResult = mBase.HResult
End Property

''
' Sets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @param RHS The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned. This is the same as <b>Err.Number</b>.
' @see Exception
'
Public Property Let HResult(ByVal RHS As Long)
    mBase.HResult = RHS
End Property

''
' Gets a description of the source of the exception.
'
' @return A description of the source of the exception.
' @remarks The source of an exception generally will contain the name of
' the function being called when the exception was thrown. This is to help
' narrow down exactly where the exception had occurred.
'
Public Property Get Source() As String
    Source = mBase.Source
End Property

''
' Sets a description of the source of the exception.
'
' @param RHS A description of the source of the exception.
' @remarks The source of an exception generally will contain the name of
' the function being called when the exception was thrown. This is to help
' narrow down exactly where the exception had occurred.
'
Public Property Let Source(ByVal RHS As String)
    mBase.Source = RHS
End Property

''
' Gets the error message associated with the Subclass exception.
'
' @return A custom message set by the Subclass, or a default
' message of "An Error has occurred."
' @remarks Generally this property is set to a meaningful message that
' is related to the exception that is being thrown. The message should
' be human readable.
' <p>This property can be set in the constructor <b>NewException</b>.
'
Public Property Get Message() As String
    Message = mBase.Message
End Property

''
' Gets the exception that caused the Subclassed exception to be thrown.
'
' @return The inner exception that caused the current exception to be thrown.
' @remarks when an exception is thrown and that exception causes another
' exception to be thrown, then the <b>InnerException</b> of the new <b>Exception</b>
' object should contain the exception that caused it to be thrown.
'
Public Property Get InnerException() As Exception
    Set InnerException = mBase.InnerException
End Property

''
' Gets the original exception that caused the chain of exceptions to occur.
'
' @return The <b>Exception</b> that caused the chain of exceptions to occur.
' @remarks If exceptions set their <b>InnerException</b> to a previously thrown
' exception, then a chain of exceptions can be created. Using this function will
' traverse that chain of exceptions until the original exception is reached. That
' exception with then be returned to the caller.
' <p>When an <b>InnerException</b> of Nothing is reached, then then the exception object is returned
' as the base exception because it did not have an inner exception, so it is assumed that
' the exception object is the last in the chain and therefore the cause of the
' chain of exceptions being iterated.
'
Public Function GetBaseException() As Exception
    Set GetBaseException = mBase.GetBaseException
End Function

''
' Returns the exception message prepended with the type name of the Subclass Exception.
'
' @return A formatted message containing the original message and possible type of exception.
' @remarks A general format might look like this:<br>
' VBCorLib.SystemException: An Error has occurred.
' <p>A listing of all inner exceptions will be included in the return value.</p>
'
Public Function ToString() As String
    ToString = mBase.ToString(Message)
End Function

''
' Provides a basic implementation of the Equals function of the <b>IObject</b> interface.
'
' @param Value The value to determine if is the same object instance as the exception.
' @return Returns True if Value is the same instance as the exception object, False otherwise.
'
Public Function Equals(ByRef value As Variant) As Boolean
    Equals = mBase.Equals(value)
End Function

''
' Provides a basic implementation of the GetHashcode function of the <b>IObject</b> interface.
'
' @return A 32-bit value used to help identify the exception object.
' @remarks The value generated is not unique across all hashcodes. Additional
' measures must be taken to find a unique value that happens to have the same
' hashcode as the exception object.
'
Public Function GetHashCode() As Long
    GetHashCode = mBase.GetHashCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' Use this method to initialize a new instance of the class.
'
Friend Sub Init(ByRef Message As String, ByRef InnerException As Exception)
    Set mBase = Nothing
    Set mBase = Cor.NewExceptionBase(Me, App.Title, Message, InnerException, DEF_HRESULT)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    Call Init(DEF_MESSAGE, Nothing)
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Set mBase = PropBag.ReadProperty(PROP_BASE)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(PROP_BASE, mBase)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Exception Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Exception_Equals(value As Variant) As Boolean
    Exception_Equals = Equals(value)
End Function

Private Function Exception_GetBaseException() As VBCorLib.Exception
    Set Exception_GetBaseException = GetBaseException
End Function

Private Function Exception_GetHashCode() As Long
    Exception_GetHashCode = GetHashCode
End Function

Private Property Let Exception_HelpLink(ByVal RHS As String)
    HelpLink = RHS
End Property

Private Property Get Exception_HelpLink() As String
    Exception_HelpLink = HelpLink
End Property

Private Property Let Exception_HResult(ByVal RHS As Long)
    HResult = RHS
End Property

Private Property Get Exception_HResult() As Long
    Exception_HResult = HResult
End Property

Private Property Get Exception_InnerException() As VBCorLib.Exception
    Set Exception_InnerException = InnerException
End Property

Private Property Get Exception_Message() As String
    Exception_Message = Message
End Property

Private Property Let Exception_Source(ByVal RHS As String)
    Source = RHS
End Property

Private Property Get Exception_Source() As String
    Exception_Source = Source
End Property

Private Function Exception_ToString() As String
    Exception_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(value As Variant) As Boolean
    IObject_Equals = Equals(value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SystemException Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SystemException_Equals(value As Variant) As Boolean
    SystemException_Equals = Equals(value)
End Function

Private Function SystemException_GetBaseException() As VBCorLib.Exception
    Set SystemException_GetBaseException = GetBaseException
End Function

Private Function SystemException_GetHashCode() As Long
    SystemException_GetHashCode = GetHashCode
End Function

Private Property Let SystemException_HelpLink(ByVal RHS As String)
    HelpLink = RHS
End Property

Private Property Get SystemException_HelpLink() As String
    SystemException_HelpLink = HelpLink
End Property

Private Property Let SystemException_HResult(ByVal RHS As Long)
    HResult = RHS
End Property

Private Property Get SystemException_HResult() As Long
    SystemException_HResult = HResult
End Property

Private Property Get SystemException_InnerException() As VBCorLib.Exception
    Set SystemException_InnerException = InnerException
End Property

Private Property Get SystemException_Message() As String
    SystemException_Message = Message
End Property

Private Property Let SystemException_Source(ByVal RHS As String)
    Source = RHS
End Property

Private Property Get SystemException_Source() As String
    SystemException_Source = Source
End Property

Private Function SystemException_ToString() As String
    SystemException_ToString = ToString
End Function

