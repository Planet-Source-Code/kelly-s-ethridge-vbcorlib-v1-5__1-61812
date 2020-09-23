VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncodingData
'
Option Explicit

Private mLongIndexes As Collection
Private mEncodingName As String
Private mPool As StringPool



Public Sub Parse(ByVal EncodingInfoNode As IXMLDOMNode, ByVal EncodingName As String, ByVal pool As StringPool)
    Dim node As IXMLDOMNode
    
    mEncodingName = EncodingName
    Set mPool = pool
    
    For Each node In EncodingInfoNode.childNodes
        If node.Attributes(0).Text = "String" Then
            pool.Add mEncodingName, node.nodeName, node.Text
        Else
            mLongIndexes.Add CLng(node.Text), node.nodeName
        End If
    Next node
End Sub

Public Sub WriteLongsTo(ByVal file As FileWriter)
    With file
        .WriteLong mLongIndexes("CodePage")
        .WriteLong mLongIndexes("IsBrowserDisplay")
        .WriteLong mLongIndexes("IsBrowserSave")
        .WriteLong mLongIndexes("IsMailNewsDisplay")
        .WriteLong mLongIndexes("IsMailNewsSave")
        .WriteLong mLongIndexes("WindowsCodePage")
    End With
End Sub

Public Sub WriteStringIndexesTo(ByVal file As FileWriter)
    With file
        '.WriteLong mPool.GetIndex(mEncodingName, "EnglishName")
        .WriteLong mPool.GetIndex(mEncodingName, "BodyName")
        .WriteLong mPool.GetIndex(mEncodingName, "EncodingName")
        .WriteLong mPool.GetIndex(mEncodingName, "HeaderName")
        .WriteLong mPool.GetIndex(mEncodingName, "WebName")
    End With
End Sub

Public Property Get Length() As Long
    Length = 40     ' 4 bytes per string index, 4 bytes per long
End Property

Private Sub Class_Initialize()
    Set mLongIndexes = New Collection
End Sub




