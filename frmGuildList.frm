VERSION 5.00
Begin VB.Form frmGuildList 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "서버 열기"
   ClientHeight    =   3180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmGuildList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer timTimer 
      Interval        =   1
      Left            =   5040
      Top             =   2760
   End
   Begin VB.Frame fGuildInfo 
      Caption         =   "서버 요약 정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtPermissions 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtGuildID 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "권한수(&R):"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "서버 &ID:"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "서버 목록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
      Begin VB.ListBox lvGuilds 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Object
Dim Http As New WinHttp.WinHttpRequest

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    EnableTLS Http
    lvGuilds.AddItem "(불러오는 중...)"
End Sub

Private Sub lvGuilds_Click()
    On Error Resume Next
    For I% = 1 To p("guilds").count
        If p("guilds")(I)("name") = lvGuilds.Text Then
            txtGuildID.Text = p("guilds")(I)("id")
            txtPermissions.Text = p("guilds")(I)("permissions")
            Exit Sub
        End If
    Next I
End Sub

Private Sub lvGuilds_DblClick()
    OKButton_Click
End Sub

Private Sub OKButton_Click()
    OKButton.Enabled = 0
    CancelButton.Enabled = 0
    frmMain.OpenGuild txtGuildID.Text
    Unload Me
End Sub

Private Sub timTimer_Timer()
    timTimer.Enabled = 0
    Http.Open "GET", "https://discord.com/api/v8/users/@me/guilds", False
    Http.SetRequestHeader "Content-Type", "application/json"
    Http.SetRequestHeader "Authorization", Token
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    Http.Send
    
    Set p = JSON.parse("{""guilds"":" & CStr(Http.ResponseText) & "}")
    If Http.Status >= 400 Then
        MsgBox "디스코드 API 오류입니다. (오류 코드 " & p("guilds")("code") & ")" & vbCrLf & "  " & CStr(p("guilds")("message")), 16, "문제가 발생했습니다!"
        Unload Me
        Exit Sub
    End If
    
    lvGuilds.Clear
    For I% = 1 To p("guilds").count
        lvGuilds.AddItem p("guilds")(I)("name")
    Next I
End Sub
