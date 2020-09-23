VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeFormatData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DateTimeFormatData
'
Option Explicit

Private mIntegerIndexes As Collection
Private mCultureName As String
Private mPool As StringPool



Public Sub Parse(ByVal CultureInfoNode As IXMLDOMNode, ByVal CultureName As String, ByVal pool As StringPool)
    Dim node As IXMLDOMNode
    
    mCultureName = CultureName
    Set mPool = pool
    Set mIntegerIndexes = New Collection
    
    For Each node In CultureInfoNode.childNodes
        Select Case node.Attributes(0).Text
            Case "String"
                pool.Add mCultureName, node.nodeName, node.Text
            Case "Integer"
                mIntegerIndexes.Add CInt(node.Text), node.nodeName
            Case "Array"
                pool.Add mCultureName, node.nodeName, ParseArray(node)
        End Select
    Next node
End Sub

Public Function WriteIntegersTo(ByVal file As FileWriter) As Long
    On Error GoTo writeZero
    With file
        .WriteInteger mIntegerIndexes("CalendarWeekRule")
        .WriteInteger mIntegerIndexes("FirstDayOfWeek")
    End With
    
    WriteIntegersTo = 2
    Exit Function
    
writeZero:
    file.WriteInteger 0
    Resume Next
End Function

Public Sub WriteStringIndexesTo(ByVal file As FileWriter)

    WriteStringIndex file, "AbbreviatedDayNames"
    WriteStringIndex file, "AbbreviatedMonthNames"
    WriteStringIndex file, "AMDesignator"
    WriteStringIndex file, "DateSeparator"
    WriteStringIndex file, "DayNames"
    WriteStringIndex file, "LongDatePattern"
    WriteStringIndex file, "LongTimePattern"
    WriteStringIndex file, "MonthDayPattern"
    WriteStringIndex file, "MonthNames"
    WriteStringIndex file, "PMDesignator"
    WriteStringIndex file, "ShortDatePattern"
    WriteStringIndex file, "ShortTimePattern"
    WriteStringIndex file, "TimeSeparator"
    WriteStringIndex file, "YearMonthPattern"
    WriteStringIndex file, "AllLongDatePatterns"
    WriteStringIndex file, "AllShortDatePatterns"
    WriteStringIndex file, "AllLongTimePatterns"
    WriteStringIndex file, "AllShortTimePatterns"
    WriteStringIndex file, "AllMonthDayPatterns"
End Sub

Public Property Get Length() As Long
    Length = 80     ' 4 bytes per string index, 2 bytes per integer
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


