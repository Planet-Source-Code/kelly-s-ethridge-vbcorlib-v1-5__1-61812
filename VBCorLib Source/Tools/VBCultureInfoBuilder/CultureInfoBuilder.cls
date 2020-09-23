VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfoBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CultureInfoBuilder
'
Option Explicit


Public Event ProcessingFile(ByVal Name As String)
Public Event WritingCulture(ByVal Name As String)


Private mCultures As Collection
Private mPool As StringPool


Public Sub Build(ByVal files As XmlFileList)
    Set mPool = New StringPool
    ParseXmlFiles files
    CreateBinaryFile
End Sub

Private Sub CreateBinaryFile()
    Dim writer As FileWriter
    Dim offsets() As Integer
    Dim i As Long
    Dim c As CultureInfo
    Dim cultureLength As Long
    
    Set c = mCultures(1)
    cultureLength = c.Length
    
    Set writer = New FileWriter
    writer.OpenFile App.Path & "\VBCultures.nlp"
    
    ' 4 bytes for number of cultures
    writer.WriteLong mCultures.Count
    
    ' 4 byte pointer to start of string pool
    writer.WriteLong cultureLength * mCultures.Count + 12 ' 12 bytes for header info
    
    ' 4 bytes for length of a culture in bytes
    writer.WriteLong cultureLength
    
'    offsets = CalculateCultureOffsets
'    For i = 0 To UBound(offsets)
'        writer.WriteInteger offsets(i)
'    Next i
    
    For Each c In mCultures
        RaiseEvent WritingCulture(c.CultureName)
        c.WriteIntegersTo writer
        c.WriteStringIndexesTo writer
    Next c
    
    mPool.WriteTo writer
    
    writer.CloseFile
End Sub

' offsets are calculated from zero, not from the first byte
' after the Header.
Private Function CalculateCultureOffsets() As Integer()
    Dim ret() As Integer
    Dim c As CultureInfo
    Dim prev As CultureInfo
    Dim i As Long
    
    ReDim ret(0 To mCultures.Count - 1)
    For Each c In mCultures
        If i = 0 Then
            ret(0) = mCultures.Count * 2 + 8
        Else
            ret(i) = ret(i - 1) + prev.Length
        End If
        Set prev = c
        i = i + 1
    Next c
    CalculateCultureOffsets = ret
End Function

Private Function SumCultureLengths() As Long
    Dim c As CultureInfo
    Dim ret As Long
    
    For Each c In mCultures
        ret = ret + c.Length
    Next c
    SumCultureLengths = ret
End Function

Private Sub ParseXmlFiles(ByVal files As XmlFileList)
    Dim file As Variant
    Dim culture As CultureInfo
    Dim doc As DOMDocument
    Dim node As IXMLDOMNode
    
    Set mCultures = New Collection
    
    For Each file In files
        Set doc = New DOMDocument
        If doc.Load(file) Then
            RaiseEvent ProcessingFile(ParseFileName(file))
            Set node = doc.selectSingleNode("Data")
            Set culture = New CultureInfo
            culture.Parse node, ParseFileName(file), mPool
            mCultures.Add culture
        End If
    Next file
End Sub

Private Function ParseFileName(ByVal file As String) As String
    Dim i As Long
    
    i = InStrRev(file, "\")
    If i > 0 Then
        ParseFileName = Mid$(file, i + 1)
    Else
        ParseFileName = file
    End If
End Function
