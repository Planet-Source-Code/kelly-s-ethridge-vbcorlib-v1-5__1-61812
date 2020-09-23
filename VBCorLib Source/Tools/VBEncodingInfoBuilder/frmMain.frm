VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Encoding Info Builder"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFolderBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pbrFilesParsed 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrEncodingsWritten 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblXmlDataFilesLocation 
      Caption         =   "Xml Data Files Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label lblParsingStatus 
      Caption         =   "Parsing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblFilesParsed 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblEncodingsWritten 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblWritingStatus 
      Caption         =   "Writing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' frmMain
'
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub SHFree Lib "shell32.dll" (ByVal pv As Long)

Private Const BIF_RETURNONLYFSDIRS      As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN     As Long = &H2
Private Const BIF_NEWDIALOGSTYLE        As Long = &H40
Private Const BIF_NONEWFOLDERBUTTON     As Long = &H200
Private Const MAX_PATH                  As Long = 260

Private Type BrowseInfo
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Private WithEvents mBuilder As EncodingInfoBuilder
Attribute mBuilder.VB_VarHelpID = -1
Private mFilesParsed As Long
Private mTotalFiles As Long
Private mEncodingsWritten As Long



Private Sub cmdBuild_Click()
    Dim files As XmlFileList
    
    mFilesParsed = 0
    mEncodingsWritten = 0
    
    Set files = New XmlFileList
    files.Load txtPath.Text
    
    If files.Count = 0 Then Exit Sub
    
    mTotalFiles = files.Count
    pbrFilesParsed.max = mTotalFiles
    pbrEncodingsWritten.max = mTotalFiles
    
    Set mBuilder = New EncodingInfoBuilder
    mBuilder.Build files
End Sub

Private Sub cmdFolderBrowse_Click()
    BrowseForFolder
    SaveSetting "VBEncodingInfoBuilder", "Settings", "Path", txtPath.Text
End Sub

Private Sub Form_Load()
    txtPath.Text = GetSetting("VBEncodingInfoBuilder", "Settings", "Path", App.Path)
End Sub

Private Sub mBuilder_ProcessingFile(ByVal Name As String)
    mFilesParsed = mFilesParsed + 1
    lblFilesParsed.Caption = mFilesParsed & " of " & mTotalFiles & " (" & Name & ")"
    lblFilesParsed.Refresh
    pbrFilesParsed.Value = mFilesParsed
End Sub

Private Sub mBuilder_WritingEncoding(ByVal Name As String)
    mEncodingsWritten = mEncodingsWritten + 1
    lblEncodingsWritten.Caption = mEncodingsWritten & " of " & mTotalFiles & " (" & Name & ")"
    lblEncodingsWritten.Refresh
    pbrEncodingsWritten.Value = mEncodingsWritten
End Sub

Private Sub BrowseForFolder()
    Dim ID As Long
    Dim buf As String
    Dim info As BrowseInfo
    
    With info
        .hOwner = Me.hWnd
        .lpszTitle = "Select the folder containing the XML Encoding data files."
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE Or BIF_NONEWFOLDERBUTTON Or BIF_DONTGOBELOWDOMAIN
    End With
    
    ID = SHBrowseForFolder(info)
    
    If ID <> 0 Then
        buf = String$(MAX_PATH, 0)
        SHGetPathFromIDList ID, buf
        txtPath.Text = Left$(buf, InStr(buf, vbNullChar) - 1)
        SHFree ID
    End If
End Sub
        



