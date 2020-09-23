VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElementData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Element
'
Option Explicit

Private mData As Collection


Public Sub Parse(ByVal CultureInfoNode As IXMLDOMNode)
    Dim node As IXMLDOMNode
    Dim att As IXMLDOMNode
    Dim arr As ElementArray
    
    For Each node In CultureInfoNode.childNodes
        Set att = node.Attributes(0)
        If Not att Is Nothing Then
            If LCase$(att.nodeName) = "isarray" And LCase$(att.Text) = "true" Then
                Set arr = New ElementArray
                arr.Parse node
                mData.Add arr, node.nodeName
            End If
        Else
            mData.Add node.Text, node.nodeName
        End If
    Next node
End Sub


Private Sub Class_Initialize()
    Set mData = New Collection
End Sub

