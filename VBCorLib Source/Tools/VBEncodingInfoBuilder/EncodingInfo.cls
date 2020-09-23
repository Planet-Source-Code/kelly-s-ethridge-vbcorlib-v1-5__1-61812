VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncodingInfo
'
Option Explicit

Private mEncodingData As EncodingData
Private mEncodingName As String



Public Property Get EncodingName() As String
    EncodingName = mEncodingName
End Property

Public Sub Parse(ByVal node As IXMLDOMNode, ByVal EncodingName As String, ByVal pool As StringPool)
    Dim childnode As IXMLDOMNode
    
    mEncodingName = EncodingName
    
    For Each childnode In node.childNodes
        If LCase$(childnode.nodeName) = "encodinginfo" Then
            mEncodingData.Parse childnode, EncodingName, pool
        End If
    Next childnode
End Sub

Public Sub WriteLongsTo(ByVal file As FileWriter)
    mEncodingData.WriteLongsTo file
End Sub

Public Sub WriteStringIndexesTo(ByVal file As FileWriter)
    mEncodingData.WriteStringIndexesTo file
End Sub

Public Property Get Length() As Long
    Dim ret As Long
    
    ret = mEncodingData.Length
    Length = ret
End Property

Private Sub Class_Initialize()
    Set mEncodingData = New EncodingData
End Sub

