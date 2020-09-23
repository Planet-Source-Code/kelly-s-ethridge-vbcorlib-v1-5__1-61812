VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CultureData
'
Option Explicit

Private mIntegerIndexes As Collection
Private mCultureName As String
Private mPool As StringPool



Public Sub Parse(ByVal CultureInfoNode As IXMLDOMNode, ByVal CultureName As String, ByVal pool As StringPool)
    Dim node As IXMLDOMNode
    
    mCultureName = CultureName
    Set mPool = pool
    
    For Each node In CultureInfoNode.childNodes
        Select Case node.Attributes(0).Text
            Case "String"
                pool.Add mCultureName, node.nodeName, node.Text
            Case "Integer"
                mIntegerIndexes.Add CInt(node.Text), node.nodeName
            Case "Array"
                pool.Add mCultureName, node.nodeName, ParseArray(node)
        End Select

'        If node.Attributes(0).Text = "String" Then
'            pool.Add mCultureName, node.nodeName, node.Text
'        Else
'            mIntegerIndexes.Add CInt(node.Text), node.nodeName
'        End If
    Next node
End Sub

Public Function WriteIntegersTo(ByVal file As FileWriter) As Long
    With file
        .WriteInteger mIntegerIndexes("LCID")
        .WriteInteger mIntegerIndexes("ParentLCID")
        .WriteInteger mIntegerIndexes("Calendar")
    End With
    WriteIntegersTo = 3
End Function

Public Sub WriteStringIndexesTo(ByVal file As FileWriter)
    With file
        .WriteLong mPool.GetIndex(mCultureName, "EnglishName")
        .WriteLong mPool.GetIndex(mCultureName, "DisplayName")
        .WriteLong mPool.GetIndex(mCultureName, "Name")
        .WriteLong mPool.GetIndex(mCultureName, "NativeName")
        .WriteLong mPool.GetIndex(mCultureName, "ThreeLetterISOLanguageName")
        .WriteLong mPool.GetIndex(mCultureName, "TwoLetterISOLanguageName")
        .WriteLong mPool.GetIndex(mCultureName, "ThreeLetterWindowsLanguageName")
        .WriteLong mPool.GetIndex(mCultureName, "OptionalCalendars")
    End With
End Sub

Public Property Get Length() As Long
    Length = 38     ' 4 bytes per string index, 2 bytes per integer
End Property

Private Function ParseArray(ByVal node As IXMLDOMNode) As ElementArray
    Set ParseArray = New ElementArray
    ParseArray.Parse node
End Function

Private Sub Class_Initialize()
    Set mIntegerIndexes = New Collection
End Sub



