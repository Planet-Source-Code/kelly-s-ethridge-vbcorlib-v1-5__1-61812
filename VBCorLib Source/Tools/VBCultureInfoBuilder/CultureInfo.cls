VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CultureInfo
'
Option Explicit

Private mCultureData As CultureData
Private mNumberFormatData As NumberFormatData
Private mDateTimeFormatData As DateTimeFormatData
Private mCultureName As String



Public Property Get CultureName() As String
    CultureName = mCultureName
End Property

Public Sub Parse(ByVal node As IXMLDOMNode, ByVal CultureName As String, ByVal pool As StringPool)
    Dim childnode As IXMLDOMNode
    
    mCultureName = CultureName
    
    For Each childnode In node.childNodes
        Select Case LCase$(childnode.nodeName)
            Case "cultureinfo"
                mCultureData.Parse childnode, CultureName, pool
            Case "datetimeformatinfo"
                mDateTimeFormatData.Parse childnode, CultureName, pool
            Case "numberformatinfo"
                mNumberFormatData.Parse childnode, CultureName, pool
        End Select
    Next childnode
End Sub

Public Sub WriteIntegersTo(ByVal file As FileWriter)
    Dim size As Long
    
    size = mCultureData.WriteIntegersTo(file)
    size = size + mDateTimeFormatData.WriteIntegersTo(file)
    size = size + mNumberFormatData.WriteIntegersTo(file)
    If (size Mod 4) <> 0 Then file.WriteInteger 0   ' 4-byte align the file
End Sub

Public Sub WriteStringIndexesTo(ByVal file As FileWriter)
    mCultureData.WriteStringIndexesTo file
    mDateTimeFormatData.WriteStringIndexesTo file
    mNumberFormatData.WriteStringIndexesTo file
End Sub

Public Property Get Length() As Long
    Dim ret As Long
    
    ret = mCultureData.Length
    ret = ret + mNumberFormatData.Length
    ret = ret + mDateTimeFormatData.Length
    If (ret Mod 4) <> 0 Then ret = ret + 2  ' 4-byte align the culture layout
    Length = ret
End Property

Private Sub Class_Initialize()
    Set mCultureData = New CultureData
    Set mDateTimeFormatData = New DateTimeFormatData
    Set mNumberFormatData = New NumberFormatData
End Sub
