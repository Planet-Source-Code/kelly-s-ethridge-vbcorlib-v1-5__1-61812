VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StringPool
'
Option Explicit

' A string could be a string array, so we need to declare this as variant.
Private mStrings() As Variant
Private mIndexes As Collection
Private mCount As Long
Private mDuplicates As Long



Public Sub Add(ByVal EncodingName As String, ByVal ValueName As String, ByVal Value As Variant)
    Dim Index As Long
    
    If VarType(Value) = vbString Then Value = InsertUnicodeCharacters(Value)
    
    Index = FindDuplicate(Value)
    If Index < 0 Then
        Index = mCount
        AppendValue Value
    Else
        mDuplicates = mDuplicates + 1
    End If
    mIndexes.Add Index, CreateKey(EncodingName, ValueName)
End Sub

Public Property Get Duplicates() As Long
    Duplicates = mDuplicates
End Property

Public Function GetIndex(ByVal EncodingName As String, ByVal ValueName As String) As Long
    Dim i As Long
    Dim ret As Long
    Dim max As Long
    
    max = mIndexes(CreateKey(EncodingName, ValueName)) - 1
    
    For i = 0 To max
'        If IsObject(mStrings(i)) Then
'            Set a = mStrings(i)
'            ret = ret + a.Length    ' get length for array of strings
'        Else
            ret = ret + LenB(mStrings(i)) + 2   ' 2 bytes for terminating null
'        End If
    Next i
    GetIndex = ret
End Function

Public Sub WriteTo(ByVal file As FileWriter)
    Dim i As Long
'    Dim a As ElementArray
    
    For i = 0 To UBound(mStrings)
'        If IsObject(mStrings(i)) Then
'            Set a = mStrings(i)
'            a.WriteTo file
'        Else
            file.WriteString mStrings(i)
'        End If
    Next i
End Sub

Public Property Get Length() As Long
    Dim i As Long
'    Dim a As ElementArray
    Dim ret As Long
    
    For i = 0 To mCount - 1
'        If IsObject(mStrings(i)) Then
'            Set a = mStrings(i)
'            ret = ret + a.Length
'        Else
            ret = ret + Len(mStrings(i)) + 1
'        End If
    Next i
    Length = ret
End Property

Private Function InsertUnicodeCharacters(ByVal s As String) As String
    Dim i As Long
    Dim Code As String
    
    i = InStr(1, s, "\U", vbTextCompare)
    Do While i > 0
        Code = Mid$(s, i + 2, 4)
        s = Replace$(s, "\U" & Code, ChrW$(Val("&H" & Code)), , , vbTextCompare)
        i = InStr(i, s, "\U", vbTextCompare)
    Loop
    InsertUnicodeCharacters = s
End Function

Private Sub AppendValue(ByRef Value As Variant)
    EnsureCapacity mCount + 1
'    If IsObject(Value) Then
'        Set mStrings(mCount) = Value
'    Else
        mStrings(mCount) = Value
'    End If
    mCount = mCount + 1
End Sub

Private Function CreateKey(ByVal EncodingName As String, ByVal ValueName As String) As String
    CreateKey = EncodingName & "_" & ValueName
End Function

Private Function FindDuplicate(ByRef Value As Variant) As Long
'    If IsObject(Value) Then
'        FindDuplicate = FindArrayDuplicate(Value)
'    Else
        FindDuplicate = FindStringDuplicate(Value)
'    End If
End Function

Private Function FindStringDuplicate(ByVal s As String) As Long
    Dim i As Long
    
    For i = 0 To mCount - 1
        If Not IsObject(mStrings(i)) Then
            If mStrings(i) = s Then
                FindStringDuplicate = i
                Exit Function
            End If
        End If
    Next i
    FindStringDuplicate = -1
End Function

'Private Function FindArrayDuplicate(ByVal arr As ElementArray) As Long
'    Dim i As Long
'    Dim a As ElementArray
'
'    For i = 0 To mCount - 1
'        If IsObject(mStrings(i)) Then
'            If arr.Equals(mStrings(i)) Then
'                FindArrayDuplicate = i
'                Exit Function
'            End If
'        End If
'    Next i
'    FindArrayDuplicate = -1
'End Function

Private Property Get Capacity() As Long
    Capacity = UBound(mStrings) + 1
End Property


Private Sub EnsureCapacity(ByVal requiredCapacity As Long)
    Dim newCapacity As Long
    
    If requiredCapacity <= Capacity Then Exit Sub
    newCapacity = Capacity * 2
    If newCapacity < requiredCapacity Then newCapacity = requiredCapacity * 2
    ReDim Preserve mStrings(0 To newCapacity - 1)
End Sub

Private Sub Class_Initialize()
    ReDim mStrings(31)
    Set mIndexes = New Collection
End Sub
