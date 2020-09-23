VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingInfoBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncodingsBuilder
'
Option Explicit

Public Event ProcessingFile(ByVal Name As String)
Public Event WritingEncoding(ByVal Name As String)


Private mEncodings As Collection
Private mPool As StringPool


Public Sub Build(ByVal files As XmlFileList)
    Set mPool = New StringPool
    ParseXmlFiles files
    CreateBinaryFile
End Sub

Private Sub CreateBinaryFile()
    Dim writer As FileWriter
    Dim Offsets() As Integer
    Dim i As Long
    Dim c As EncodingInfo
    Dim EncodingLength As Long
    
    Set c = mEncodings(1)
    EncodingLength = c.Length
    
    Set writer = New FileWriter
    writer.OpenFile App.Path & "\VBEncodings.nlp"
    
    ' 4 bytes for number of Encodings
    writer.WriteLong mEncodings.Count
    
    ' 4 byte pointer to start of string pool
    writer.WriteLong EncodingLength * mEncodings.Count + 12 ' 12 bytes for header info
    
    ' 4 bytes for length of a Encoding in bytes
    writer.WriteLong EncodingLength
    
'    offsets = CalculateEncodingOffsets
'    For i = 0 To UBound(offsets)
'        writer.WriteInteger offsets(i)
'    Next i
    
    For Each c In mEncodings
        RaiseEvent WritingEncoding(c.EncodingName)
        c.WriteLongsTo writer
        c.WriteStringIndexesTo writer
    Next c
    
    mPool.WriteTo writer
    
    writer.CloseFile
End Sub

' offsets are calculated from zero, not from the first byte
' after the Header.
Private Function CalculateEncodingOffsets() As Integer()
    Dim ret() As Integer
    Dim c As EncodingInfo
    Dim prev As EncodingInfo
    Dim i As Long
    
    ReDim ret(0 To mEncodings.Count - 1)
    For Each c In mEncodings
        If i = 0 Then
            ret(0) = mEncodings.Count * 2 + 8
        Else
            ret(i) = ret(i - 1) + prev.Length
        End If
        Set prev = c
        i = i + 1
    Next c
    CalculateEncodingOffsets = ret
End Function

Private Function SumEncodingLengths() As Long
    Dim c As EncodingInfo
    Dim ret As Long
    
    For Each c In mEncodings
        ret = ret + c.Length
    Next c
    SumEncodingLengths = ret
End Function

Private Sub ParseXmlFiles(ByVal files As XmlFileList)
    Dim file As Variant
    Dim Encoding As EncodingInfo
    Dim doc As DOMDocument
    Dim node As IXMLDOMNode
    
    Set mEncodings = New Collection
    
    For Each file In files
        Set doc = New DOMDocument
        If doc.Load(file) Then
            RaiseEvent ProcessingFile(ParseFileName(file))
            Set node = doc.selectSingleNode("Data")
            Set Encoding = New EncodingInfo
            Encoding.Parse node, ParseFileName(file), mPool
            mEncodings.Add Encoding
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

