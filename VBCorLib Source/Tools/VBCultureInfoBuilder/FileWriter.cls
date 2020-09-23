VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FileWriter
'
Option Explicit

Private mFileNumber As Long


Public Sub OpenFile(ByVal FileName As String)
    mFileNumber = FreeFile
    Open FileName For Binary As #mFileNumber
End Sub

Public Sub CloseFile()
    Close #mFileNumber
End Sub

Public Sub WriteInteger(ByVal Value As Integer)
    Put #mFileNumber, , Value
End Sub

Public Sub WriteLong(ByVal Value As Long)
    Put #mFileNumber, , Value
End Sub

Public Sub WriteString(ByVal Value As String)
    Dim i As Long
    
    For i = 1 To Len(Value)
        WriteInteger Asc(Mid$(Value, i, 1))
    Next i
    WriteNull
End Sub

Public Sub WriteNull()
    WriteInteger 0
End Sub

