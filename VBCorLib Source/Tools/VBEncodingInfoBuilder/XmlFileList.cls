VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' XmlFileList
'
Option Explicit

Private Const FILTER_XML As String = "*.xml"

Private mFiles As Collection



Public Property Get Count() As Long
    Count = mFiles.Count
End Property

Public Property Get Item(ByVal index As Long)
Attribute Item.VB_UserMemId = 0
    Item = mFiles(index + 1)
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mFiles.[_NewEnum]
End Function

Public Sub Load(ByVal Path As String)
    Dim file As String
    
    Clear
    Path = BuildProperPath(Path)
    file = Dir$(BuildFilteredPath(Path))
    Do While Len(file) > 0
        mFiles.Add Path & file
        file = Dir$
    Loop
End Sub

Public Sub Clear()
    Set mFiles = New Collection
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildFilteredPath(ByVal Path As String) As String
    BuildFilteredPath = Path & FILTER_XML
End Function

Private Function BuildProperPath(ByVal Path As String)
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    BuildProperPath = Path
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Clear
End Sub
