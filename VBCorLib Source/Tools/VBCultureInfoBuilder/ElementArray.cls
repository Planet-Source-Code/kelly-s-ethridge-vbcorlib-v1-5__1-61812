VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElementArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ElementArray
'
Option Explicit

Private mName As String
Private mElements As Collection



Public Sub Parse(ByVal node As IXMLDOMNode)
    Dim childnode As IXMLDOMNode
    Dim Index As IXMLDOMAttribute
    
    mName = node.nodeName
    Set mElements = New Collection
    
    For Each childnode In node.childNodes
        Set Index = childnode.Attributes(0)
        mElements.Add InsertUnicodeCharacters(childnode.Text), CStr(Index.Value)
    Next childnode
End Sub

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
    Item = mElements(CStr(Index))
End Property

Public Property Get Count() As Long
    Count = mElements.Count
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Function Equals(ByVal arr As ElementArray) As Boolean
    If arr.Name <> mName Then Exit Function
    If arr.Count <> mElements.Count Then Exit Function
    
    Dim i As Long
    For i = 0 To mElements.Count - 1
        If arr(i) <> Me(i) Then Exit Function
    Next i
    Equals = True
End Function

' Return the length in bytes.
Public Property Get Length() As Long
    Dim ret As Long
    Dim v As Variant
    
    For Each v In mElements
        ret = ret + LenB(v) + 2     ' 2 bytes for terminating null
    Next
    Length = ret + 2    ' 2 bytes for terminating null of array. This will cause 4 bytes of 0.
End Property

Public Sub WriteTo(ByVal file As FileWriter)
    Dim i As Long
    
    file.WriteInteger mElements.Count
    For i = 0 To mElements.Count - 1
        file.WriteString mElements(CStr(i))
    Next i
End Sub



Private Function InsertUnicodeCharacters(ByVal s As String) As String
    Dim i As Long
    Dim Code As String
    
    i = InStr(1, s, "\U", vbTextCompare)
    Do While i > 0
        Code = Mid$(s, i + 2, 4)
        s = Replace$(s, "\U" & Code, ChrW$(Val("&H" & Code)), , , vbTextCompare)
        i = InStr(i, s, "\U", vbTextCompare)
    Loop
    InsertUnicodeCharacters = s
End Function

