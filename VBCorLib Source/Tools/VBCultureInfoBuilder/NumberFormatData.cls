VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberFormatData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' NumberFormatData
'
Option Explicit

Private mIntegers As Collection
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
                mIntegers.Add CInt(node.Text), node.nodeName
            Case "Array"
                pool.Add mCultureName, node.nodeName, ParseArray(node)
        End Select
    Next node
End Sub

Public Function WriteIntegersTo(ByVal file As FileWriter) As Long
    On Error GoTo writeZero
    With file
        .WriteInteger mIntegers("CurrencyDecimalDigits")
        '.WriteInteger mPool.GetIndex(mCultureName, "CurrencyGroupSizes")
        .WriteInteger mIntegers("CurrencyNegativePattern")
        .WriteInteger mIntegers("CurrencyPositivePattern")
        .WriteInteger mIntegers("NumberDecimalDigits")
        '.WriteInteger mPool.GetIndex(mCultureName, "NumberGroupSizes")
        .WriteInteger mIntegers("NumberNegativePattern")
        .WriteInteger mIntegers("PercentDecimalDigits")
        '.WriteInteger mPool.GetIndex(mCultureName, "PercentGroupSizes")
        .WriteInteger mIntegers("PercentNegativePattern")
        .WriteInteger mIntegers("PercentPositivePattern")
    End With
    
    WriteIntegersTo = 8
    Exit Function
    
writeZero:
    file.WriteInteger 0
    Resume Next
End Function

Public Sub WriteStringIndexesTo(ByVal file As FileWriter)
    With file
        WriteStringIndex file, "CurrencyGroupSizes"
        WriteStringIndex file, "NumberGroupSizes"
        WriteStringIndex file, "PercentGroupSizes"
        WriteStringIndex file, "CurrencyDecimalSeparator"
        WriteStringIndex file, "CurrencyGroupSeparator"
        WriteStringIndex file, "CurrencySymbol"
        WriteStringIndex file, "NaNSymbol"
        WriteStringIndex file, "NegativeInfinitySymbol"
        WriteStringIndex file, "NegativeSign"
        WriteStringIndex file, "NumberDecimalSeparator"
        WriteStringIndex file, "NumberGroupSeparator"
        WriteStringIndex file, "PercentDecimalSeparator"
        WriteStringIndex file, "PercentGroupSeparator"
        WriteStringIndex file, "PercentSymbol"
        WriteStringIndex file, "PerMilleSymbol"
        WriteStringIndex file, "PositiveInfinitySymbol"
        WriteStringIndex file, "PositiveSign"
    End With
End Sub

Public Property Get Length() As Long
    Length = 84     ' 4 bytes per string index, 2 bytes per integer
End Property

Private Sub WriteStringIndex(ByVal file As FileWriter, ByVal Name As String)
    On Error Resume Next
    file.WriteLong mPool.GetIndex(mCultureName, Name)
    If Err.Number <> 0 Then file.WriteLong 0
End Sub

Private Function ParseArray(ByVal node As IXMLDOMNode) As ElementArray
    Set ParseArray = New ElementArray
    ParseArray.Parse node
End Function

Private Sub Class_Initialize()
    Set mIntegers = New Collection
End Sub

