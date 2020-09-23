VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Compile It!"
   ClientHeight    =   2835
   ClientLeft      =   4185
   ClientTop       =   1635
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOptions 
      Caption         =   "Options"
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6435
      Begin VB.CommandButton BtnCompile 
         Caption         =   "Compile"
         Height          =   495
         Left            =   4800
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox cOutput 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Text            =   "C:\My Documents\Visual Studio Projects"
         Top             =   300
         Width           =   5115
      End
      Begin VB.TextBox cProjectList 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Text            =   "C:\ProjectList.Txt"
         Top             =   660
         Width           =   5115
      End
      Begin VB.TextBox cVBPath 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Text            =   "C:\Program Files\Microsoft Visual Studio\VB98\Vb6.exe"
         Top             =   1020
         Width           =   5115
      End
      Begin VB.Label lStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting..."
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1740
         Width           =   4515
      End
      Begin VB.Label lCompiling 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiling project:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lOutput 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lProjectList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project List:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lVBPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB Path:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   1080
         Width           =   930
      End
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   60
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "http://developers.sytes.net"
      Top             =   2280
      Width           =   2985
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub BtnCompile_Click()
    Dim Str As String
    
    Open cProjectList.Text For Input As #1

    BtnCompile.Enabled = False
    While Not EOF(1)
        Input #1, Str
        lStatus.Caption = Str
        Form1.Refresh
        DoEvents
        CompileProject (Str)
    Wend
    BtnCompile.Enabled = True
    lStatus.Caption = ""
    Close #1
End Sub


Public Sub CompileProject(project As String)
    Dim cmd As String
    
    cmd = """" + cVBPath.Text + """ /m """ + project + """ /outdir """ + cOutput.Text + """"
    Call Spawn(cmd, True)
End Sub


Private Sub Form_Load()
   cProjectList.Text = App.Path + "\ProjectList.txt"
End Sub

Private Sub imgLogo_Click()
    Call ShellExecute(Me.hwnd, vbNullString, "http://developers.sytes.net", vbNullString, "c:\", SW_SHOWNORMAL)
End Sub
