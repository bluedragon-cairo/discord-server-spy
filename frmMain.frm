VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.3#0"; "TabExT01.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  '���� ����
   Caption         =   "���ڵ� ���� ������"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows �⺻��
   Begin TabExCtl.SSTabEx ssTabs 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8493
      Tabs            =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   12
      Style           =   2
      TabHeight       =   536
      TabMinWidth     =   1323
      TabSelHighlight =   -1  'True
      TabWidthStyle   =   1
      TabAppearance   =   1
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "�Ϲ�"
      Tab(0).ControlCount=   15
      Tab(0).Control(0)=   "txtGuildID"
      Tab(0).Control(1)=   "txtGuildName"
      Tab(0).Control(2)=   "txtGuildDescription"
      Tab(0).Control(3)=   "pbBoostProgress"
      Tab(0).Control(4)=   "cmdSaveIcon"
      Tab(0).Control(5)=   "lvFeatures"
      Tab(0).Control(6)=   "imgGuildIcon"
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(9)=   "Label16"
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(11)=   "lblBoostCount"
      Tab(0).Control(12)=   "Label18"
      Tab(0).Control(13)=   "lblMemberCount"
      Tab(0).Control(14)=   "Label19"
      TabCaption(1)   =   "ä��"
      Tab(1).ControlCount=   3
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "fChannelInfo"
      Tab(1).Control(2)=   "Frame2"
      TabCaption(2)   =   "����"
      Tab(2).ControlCount=   5
      Tab(2).Control(0)=   "fPermissions"
      Tab(2).Control(1)=   "chkHoistRole"
      Tab(2).Control(2)=   "chkMentionableRole"
      Tab(2).Control(3)=   "fRoleInfo"
      Tab(2).Control(4)=   "Frame1"
      TabCaption(3)   =   "���"
      Tab(3).ControlCount=   2
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame10"
      TabCaption(4)   =   "�̸���"
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "�ʴ�"
      Tab(5).ControlCount=   7
      Tab(5).Control(0)=   "Command3"
      Tab(5).Control(1)=   "cmdInfiniteUses"
      Tab(5).Control(2)=   "cmdInfiniteAge"
      Tab(5).Control(3)=   "chkTemporary"
      Tab(5).Control(4)=   "Frame4"
      Tab(5).Control(5)=   "fInviteInfo"
      Tab(5).Control(6)=   "Frame5"
      TabCaption(6)   =   "�� ���"
      Tab(6).ControlCount=   3
      Tab(6).Control(0)=   "Frame6"
      Tab(6).Control(1)=   "fBanInfo"
      Tab(6).Control(2)=   "cmdUnban"
      TabCaption(7)   =   "���� ����"
      Tab(7).ControlCount=   5
      Tab(7).Control(0)=   "Frame7"
      Tab(7).Control(1)=   "cmdSaveAuditLog"
      Tab(7).Control(2)=   "Frame11"
      Tab(7).Control(3)=   "fAuditLogChangeInfo"
      Tab(7).Control(4)=   "ssAuditLogTabs"
      TabCaption(8)   =   "������"
      Tab(8).ControlCount=   14
      Tab(8).Control(0)=   "txtGuildRegion"
      Tab(8).Control(1)=   "cbVerificationLevel"
      Tab(8).Control(2)=   "cbNotificationLevel"
      Tab(8).Control(3)=   "cbFilterLevel"
      Tab(8).Control(4)=   "chk2FARequired"
      Tab(8).Control(5)=   "lblAFKInfo"
      Tab(8).Control(6)=   "lblWidgetInfo"
      Tab(8).Control(7)=   "Label3"
      Tab(8).Control(8)=   "Label20"
      Tab(8).Control(9)=   "Label21"
      Tab(8).Control(10)=   "Label22"
      Tab(8).Control(11)=   "Label23"
      Tab(8).Control(12)=   "Label24"
      Tab(8).Control(13)=   "Label25"
      Begin VB.Frame fPermissions 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   -72120
         TabIndex        =   39
         Top             =   2520
         Width           =   5295
         Begin VB.ListBox lvPermissions 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtPermissionDescription 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   2400
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   40
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�����(&C)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68160
         TabIndex        =   63
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdInfiniteUses 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -67800
         TabIndex        =   65
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdInfiniteAge 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -67800
         TabIndex        =   64
         Top             =   3600
         Width           =   975
      End
      Begin VB.CheckBox chkTemporary 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72240
         TabIndex        =   72
         Top             =   2280
         Width           =   255
      End
      Begin VB.Frame Frame7 
         Caption         =   "���� ����"
         Height          =   1935
         Left            =   -68520
         TabIndex        =   112
         Top             =   480
         Width           =   1935
         Begin VB.ListBox lvAuditLogChanges 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdSaveAuditLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����(&S)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70320
         TabIndex        =   85
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Frame Frame11 
         Caption         =   "���� ��� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   110
         Top             =   480
         Width           =   6255
         Begin ComctlLib.ListView lvAuditLogs 
            Height          =   1935
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame fAuditLogChangeInfo 
         Caption         =   "�ڼ��� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   107
         Top             =   2880
         Width           =   8295
         Begin VB.TextBox txtNewValue 
            Height          =   1335
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   109
            Top             =   240
            Width           =   3735
         End
         Begin VB.TextBox txtOldValue 
            Height          =   1335
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   240
            Width           =   3735
         End
         Begin VB.Line Line7 
            X1              =   4200
            X2              =   4320
            Y1              =   1075
            Y2              =   885
         End
         Begin VB.Line Line6 
            X1              =   4200
            X2              =   4320
            Y1              =   720
            Y2              =   900
         End
         Begin VB.Line Line5 
            X1              =   4200
            X2              =   4200
            Y1              =   960
            Y2              =   1080
         End
         Begin VB.Line Line4 
            X1              =   4200
            X2              =   4200
            Y1              =   840
            Y2              =   720
         End
         Begin VB.Line Line3 
            X1              =   3960
            X2              =   4200
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line2 
            X1              =   3960
            X2              =   4200
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            X1              =   3960
            X2              =   3960
            Y1              =   840
            Y2              =   960
         End
      End
      Begin VB.TextBox txtGuildRegion 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   480
         Width           =   6615
      End
      Begin VB.ComboBox cbVerificationLevel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   96
         Top             =   2160
         Width           =   6615
      End
      Begin VB.ComboBox cbNotificationLevel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   95
         Top             =   2640
         Width           =   6615
      End
      Begin VB.ComboBox cbFilterLevel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   94
         Top             =   3120
         Width           =   6615
      End
      Begin VB.CheckBox chk2FARequired 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74880
         TabIndex        =   93
         Top             =   3720
         Width           =   255
      End
      Begin VB.Frame Frame6 
         Caption         =   "��� ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   91
         Top             =   480
         Width           =   2415
         Begin VB.ListBox lvBans 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3660
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame fBanInfo 
         Caption         =   "�� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   -72360
         TabIndex        =   88
         Top             =   480
         Width           =   5775
         Begin VB.TextBox txtBanReason 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   89
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '����
            Caption         =   "���� ����(&R):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdUnban 
         Caption         =   "�� ����(&U)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -68040
         TabIndex        =   87
         Top             =   2400
         Width           =   1335
      End
      Begin ComctlLib.TabStrip ssAuditLogTabs 
         Height          =   375
         Left            =   -68280
         TabIndex        =   86
         Top             =   2595
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "���� ��"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "���� ��"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Caption         =   "�ʴ��� ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   83
         Top             =   480
         Width           =   2415
         Begin VB.ListBox lvInvites 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3660
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame fInviteInfo 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   -72360
         TabIndex        =   73
         Top             =   480
         Width           =   5775
         Begin VB.Label Label26 
            BackStyle       =   0  '����
            Caption         =   "�ʴ���:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblInviter 
            BackStyle       =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   81
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label27 
            BackStyle       =   0  '����
            Caption         =   "������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblInviteChannel 
            BackStyle       =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   79
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label28 
            BackStyle       =   0  '����
            Caption         =   "��� Ƚ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblInviteUses 
            BackStyle       =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   77
            Top             =   960
            Width           =   4575
         End
         Begin VB.Label Label29 
            BackStyle       =   0  '����
            Caption         =   "������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblExpiration 
            BackStyle       =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   75
            Top             =   1320
            Width           =   4575
         End
         Begin VB.Label Label30 
            BackStyle       =   0  '����
            Caption         =   "�ӽ� ��� �ڰ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   74
            Top             =   1800
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "�� �ʴ� �����"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   -70560
         TabIndex        =   66
         Top             =   3000
         Width           =   3855
         Begin VB.TextBox txtMaxUses 
            Alignment       =   1  '������ ����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1800
            TabIndex        =   68
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtMaxAge 
            Alignment       =   1  '������ ����
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1800
            TabIndex        =   67
            Text            =   "0"
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label31 
            BackStyle       =   0  '����
            Caption         =   "�ִ� ��� Ƚ��(&M):"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label32 
            BackStyle       =   0  '����
            Caption         =   "�ð� ����(&T):"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label33 
            BackStyle       =   0  '����
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   69
            Top             =   660
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "��� ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   61
         Top             =   360
         Width           =   2415
         Begin ComctlLib.TreeView tvMembers 
            Height          =   3735
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6588
            _Version        =   327682
            Indentation     =   542
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "��� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -72360
         TabIndex        =   53
         Top             =   360
         Width           =   5775
         Begin VB.TextBox txtUserTag 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtUserID 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label34 
            BackStyle       =   0  '����
            Caption         =   "�±�(&T):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   280
            Width           =   855
         End
         Begin VB.Label Label35 
            BackStyle       =   0  '����
            Caption         =   "����� &ID:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label36 
            BackStyle       =   0  '����
            Caption         =   "�ֻ��� �и� ����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label lblMemberRole 
            BackStyle       =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   57
            Top             =   1005
            Width           =   3855
         End
         Begin VB.Label Label37 
            BackStyle       =   0  '����
            Caption         =   "Label37"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   56
            Top             =   2280
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkHoistRole 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   220
         Left            =   -72240
         TabIndex        =   52
         Top             =   1680
         Width           =   200
      End
      Begin VB.CheckBox chkMentionableRole 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   220
         Left            =   -72240
         TabIndex        =   51
         Top             =   2040
         Width           =   200
      End
      Begin VB.Frame fRoleInfo 
         Caption         =   "���� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -72360
         TabIndex        =   42
         Top             =   480
         Width           =   5775
         Begin VB.TextBox txtRoleID 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '����
            Caption         =   "���� ����� ����� �� �ֵ��� ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   1605
            Width           =   3375
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '����
            Caption         =   "������ �¶��� ����� �и��Ͽ� ǥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   47
            Top             =   1245
            Width           =   3735
         End
         Begin VB.Label lblRoleColor 
            BackStyle       =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   46
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '����
            Caption         =   "��(&C):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '����
            Caption         =   "���� &ID: "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "���� ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   49
         Top             =   480
         Width           =   2415
         Begin VB.ListBox lvRoles 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3660
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ACL"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   -72120
         TabIndex        =   33
         Top             =   2400
         Width           =   5295
         Begin VB.ListBox lvPermissionOverwrites 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtAllow 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   35
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtDeny 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   34
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '����
            Caption         =   "��� ����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   38
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  '����
            Caption         =   "�ź� ����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   37
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame fChannelInfo 
         Caption         =   "ä�� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -72360
         TabIndex        =   19
         Top             =   360
         Width           =   5775
         Begin VB.CheckBox chkSystemChannel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   1275
            Width           =   255
         End
         Begin VB.CheckBox chkNSFW 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1275
            Width           =   255
         End
         Begin VB.TextBox txtTopic 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox txtPosition 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   4575
         End
         Begin VB.TextBox txtChannelID 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label15 
            BackStyle       =   0  '����
            Caption         =   "�ý��� ä��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   32
            Top             =   1290
            Width           =   1335
         End
         Begin VB.Label Label14 
            BackStyle       =   0  '����
            Caption         =   "�Ĺ����� ä��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   1290
            Width           =   2100
         End
         Begin VB.Label Label13 
            BackStyle       =   0  '����
            Caption         =   "����(&T):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '����
            Caption         =   "��ġ(&P):"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   "ä�� &ID: "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ä�� ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   2415
         Begin VB.ListBox lvChannels 
            Height          =   780
            Left            =   480
            TabIndex        =   22
            Top             =   2640
            Visible         =   0   'False
            Width           =   1575
         End
         Begin ComctlLib.TreeView tvChannels 
            Height          =   3735
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6588
            _Version        =   327682
            HideSelection   =   0   'False
            Indentation     =   542
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.TextBox txtGuildID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtGuildName 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtGuildDescription 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   8
         Top             =   1320
         Width           =   5415
      End
      Begin ComctlLib.ProgressBar pbBoostProgress 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   2400
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   0
         Max             =   14
      End
      Begin VB.CommandButton cmdSaveIcon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���� ������ ����(&V)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   4080
         Width           =   1935
      End
      Begin VB.ListBox lvFeatures 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   3000
         Style           =   1  'Ȯ�ζ�
         TabIndex        =   5
         Top             =   3240
         Width           =   5415
      End
      Begin VB.Label lblAFKInfo 
         BackStyle       =   0  '����
         Height          =   375
         Left            =   -73320
         TabIndex        =   106
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Label lblWidgetInfo 
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   105
         Top             =   960
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "���� ����(&R):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   104
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  '����
         Caption         =   "���� ä��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   103
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  '����
         Caption         =   "��ȭ�� ��� ���:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  '����
         Caption         =   "���� ����(&V):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   101
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackStyle       =   0  '����
         Caption         =   "�ʱ� �˸� ����(&N):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   100
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label24 
         BackStyle       =   0  '����
         Caption         =   "�޽��� �˿�(&L):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   99
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  '����
         Caption         =   "������ 2�ܰ� ���� �ʿ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   98
         Top             =   3735
         Width           =   4245
      End
      Begin VB.Image imgGuildIcon 
         Height          =   735
         Left            =   360
         Stretch         =   -1  'True
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "���� &ID:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "���� �̸�(&S):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  '����
         Caption         =   "����(&D):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  '����
         Caption         =   "�ν�Ʈ ��Ȳ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Label lblBoostCount 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "(0��)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   14
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  '����
         Caption         =   "��� ��Ȳ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblMemberCount 
         BackStyle       =   0  '����
         Caption         =   "0�� �� 0�� �¶���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  '����
         Caption         =   "���(&F):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   3240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6240
      Tag             =   "1"
      Top             =   5040
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�(&C)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "����(&O)..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetToken 
      Caption         =   "����(&T)..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Object
Dim guild As Object
Dim roles As Object
Dim channels As Object
Dim channelMap As Object
Dim roleMap As Object
Dim permissions()
Dim RolePermissions(32) As String
Dim permissionOverwrites As Object
Dim Http As New WinHttp.WinHttpRequest
Dim loadingFeatures As Boolean
Dim invitesLoaded As Boolean
Dim invites As Object
Dim inviteMap As Object
Dim bansLoaded As Boolean
Dim banMap As Object
Dim bans As Object
Dim membersLoaded As Boolean
Dim members As Object
Dim memberMap As Object
Dim auditLog As Object
Dim auditLogs As Object
Dim auditLogMap As Object
Dim auditLogLoaded As Boolean
Dim auditLogChanges As Object
Dim AuditKeys As New Dictionary
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal _
        wParam As Long, lParam As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, _
    ByVal nBkMode As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableThemeDialogTexture Lib "uxtheme" _
(ByVal hWnd As Long, _
ByVal flags As Long) As Long

Const LVS_EX_FULLROWSELECT = &H20
Const LVM_FIRST = &H1000
Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H37
Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H36

Function AuditLogType(ByVal typ As Integer) As String
    If typ = 10 Then
        AuditLogType = "ä�� ����"
    ElseIf typ = 24 Then
        AuditLogType = "��� ����"
    ElseIf typ = 20 Then
        AuditLogType = "��� �߹�"
    ElseIf typ = 11 Then
        AuditLogType = "ä�� ����"
    ElseIf typ = 40 Then
        AuditLogType = "�ʴ��� ����"
    ElseIf typ = 42 Then
        AuditLogType = "�ʴ��� ����"
    ElseIf typ = 82 Then
        AuditLogType = "���� ����"
    ElseIf typ = 1 Then
        AuditLogType = "���� ���� ����"
    Else
        AuditLogType = "#" & typ
    End If
End Function

         
Function DownloadFile(ByRef v_strURL As String, ByRef v_strPath As String) As String
    On Error Resume Next
    
    Dim iFileNo As Integer
    Dim aryData() As Byte
    
    Http.Open "GET", v_strURL, False
    Http.SetRequestHeader "Content-Type", "image/png"
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    Http.Send
    aryData = Http.ResponseBody

    iFileNo = FreeFile
    Open v_strPath For Binary As #iFileNo
        Put #iFileNo, , aryData
    Close #iFileNo
    
    DownloadFile = v_strPath
End Function

Function toBinary(ByVal x) As String
    Dim ret$
    Do While x
        ret = ret & (CInt(Right$(CStr(x), 1)) Mod 2)
        x = Fix(x / 2)
    Loop
    toBinary = StrReverse(ret)
End Function

Function ZeroFill(ByVal x As String, ByVal Length) As String
    For I% = 1 To Length - Len(x) Step 1
        x = "0" & x
    Next I
    ZeroFill = x
End Function

Function BitAnd(ByVal xi, ByVal yi)
    Dim lx%, ly%, Length
    Dim x$, y$
    Dim ret$
    x = toBinary(xi)
    y = toBinary(yi)
    
    lx = Len(x)
    ly = Len(y)
    If (lx > ly) Then
        Length = lx
    Else
        Length = ly
    End If
    
    x = ZeroFill(x, Length)
    y = ZeroFill(y, Length)
    
    ret = ""
    For I% = 1 To Length
        ret = ret & CStr(CInt(Mid$(x, I, 1)) And CInt(Mid$(y, I, 1)))
    Next I
    
    BitAnd = ret
End Function

Sub OpenGuild(ByVal GuildID As String)
    'Dim GuildID As String
    'GuildID = InputBox("���� ID: ", "����", "918102050812862514")

    Me.Caption = "(�ҷ����� ��...) - ���ڵ� ���� ������"
    Http.Open "GET", "https://discord.com/api/v8/guilds/" & GuildID & "?with_counts=true", True
    Http.SetRequestHeader "Content-Type", "application/json"
    Http.SetRequestHeader "Authorization", Token
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    Http.Send
    Http.WaitForResponse 60
    
    Set p = JSON.parse(CStr(Http.ResponseText))
    If Http.Status >= 400 Then
        MsgBox CStr(p("message")), 16, "������ �߻��߽��ϴ�!"
        Exit Sub
    End If
    
    Set guild = p
    Me.Caption = guild("name") & " - ���ڵ� ���� ������"
    txtGuildName.Text = guild("name")
    txtGuildID.Text = GuildID
    txtGuildRegion.Text = guild("region")
    If Not IsNull(guild("description")) Then _
        txtGuildDescription.Text = guild("description")
    lblMemberCount.Caption = guild("approximate_member_count") & "�� �� " & guild("approximate_presence_count") & "�� ���� ��"
    ssTabs.TabCaption(3) = "��� (" & guild("approximate_member_count") & ")"
    
    If Not IsNull(guild("premium_subscription_count")) Then
        pbBoostProgress.Value = CInt(guild("premium_subscription_count"))
        lblBoostCount.Caption = "(" & guild("premium_subscription_count") & "��)"
    Else
        pbBoostProgress.Value = 0
        lblBoostCount.Caption = "(0��)"
    End If
    
    loadingFeatures = -1
    For k% = 0 To lvFeatures.ListCount - 1
        lvFeatures.Selected(k) = 0
    Next k
    For I% = 1 To guild("features").count
        For k% = 0 To lvFeatures.ListCount - 1
            If lvFeatures.List(k) = guild("features")(I) Then _
                lvFeatures.Selected(k) = -1
        Next k
    Next I
    loadingFeatures = 0
    
    Set roles = guild("roles")
    Set roleMap = New Dictionary
    ssTabs.TabCaption(2) = "���� (" & roles.count & ")"
    lvRoles.Clear
    For I% = 1 To roles.count
        lvRoles.AddItem roles(I)("name")
        roleMap.Add roles(I)("id"), roles(I)
    Next I
    
    Dim iconPath$
    iconPath = Environ$("TEMP") & "\DISCORD_GUILD_ICON_" & GuildID & ".PNG"
    DownloadFile "https://cdn.discordapp.com/icons/" & GuildID & "/" & guild("icon") & ".png", iconPath
    Set imgGuildIcon.Picture = StdPictureEx.LoadPicture(iconPath)
    Kill iconPath
    
    Http.Open "GET", "https://discord.com/api/v8/guilds/" & GuildID & "/channels", True
    Http.SetRequestHeader "Content-Type", "application/json"
    Http.SetRequestHeader "Authorization", Token
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    Http.Send
    Http.WaitForResponse 60
    
    Set p = JSON.parse("{""channels"":" & CStr(Http.ResponseText) & "}")
    Set channelMap = New Dictionary
    Set channels = p("channels")
    ssTabs.TabCaption(1) = "ä�� (" & channels.count & ")"
    lvChannels.Clear
    tvChannels.Nodes.Clear
    For I% = 1 To channels.count
        If IsNull(channels(I)("parent_id")) Then
            If channels(I)("type") = 4 Then
                tvChannels.Nodes.Add , , CStr(channels(I)("id")), "[ " & channels(I)("name") & " ]"
            Else
                tvChannels.Nodes.Add , , CStr(channels(I)("id")), channels(I)("name")
            End If
        End If
        lvChannels.AddItem channels(I)("name")
        channelMap.Add CStr(channels(I)("id")), channels(I)
    Next I
    For I% = 1 To channels.count
        If channels(I)("parent_id") Then
            tvChannels.Nodes.Add CStr(channels(I)("parent_id")), tvwChild, channels(I)("id"), channels(I)("name")
        End If
    Next I
    
    If guild("widget_enabled") Then
        If IsNull(guild("widget_channel_id")) Then
            lblWidgetInfo.Caption = "�ʴ� �Ұ����� ���� ��� ��"
        Else
            lblWidgetInfo.Caption = "#" & channelMap(guild("widget_channel_id"))("name") & "(��)�� �ʴ�Ǵ� ���� ��� ��"
        End If
    Else
        lblWidgetInfo.Caption = "(������ Ȱ��ȭ�Ǿ� ���� �ʽ��ϴ�.)"
    End If
    
    If Not IsNull(guild("afk_channel_id")) Then
        lblAFKInfo.Caption = (CInt(guild("afk_timeout")) / 60) & "�� �� " & channelMap(guild("afk_channel_id"))("name") & " ä�η� �ڵ� �̵�"
    Else
        lblAFKInfo.Caption = (CInt(guild("afk_timeout")) / 60) & "��"
    End If
    
    cbVerificationLevel.ListIndex = guild("verification_level")
    cbNotificationLevel.ListIndex = guild("default_message_notifications")
    cbFilterLevel.ListIndex = guild("explicit_content_filter")
    
    chk2FARequired.Value = CInt(guild("mfa_level"))
    
    invitesLoaded = False
    bansLoaded = False
    auditLogLoaded = False
    membersLoaded = False
End Sub

Private Sub cmdInfiniteAge_Click()
    txtMaxAge.Text = "0"
End Sub

Private Sub cmdInfiniteUses_Click()
    txtMaxUses.Text = "0"
End Sub

Private Sub cmdOpen_Click()
    frmGuildList.Show 1, Me
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSaveAuditLog_Click()
    MsgBox "�̱���"
End Sub

Private Sub cmdSaveIcon_Click()
    Dim iconPath$
    iconPath = Environ$("USERPROFILE") & "\My Documents\My Pictures\DISCORD_GUILD_ICON_" & guild("id") & ".PNG"
    DownloadFile "https://cdn.discordapp.com/icons/" & guild("id") & "/" & guild("icon") & ".png", iconPath
End Sub

Private Sub cmdSetToken_Click()
    frmAuthentication.Show 1, Me
    Exit Sub
    
    Dim userInput As String
    userInput = InputBox("��ū ����: ", "��ū")
    If userInput <> "" Then
        Token = userInput
        SaveSetting "Discord API Explorer 2", "Authorization", "Token", Token
    End If
End Sub

Private Sub Form_Load()
    EnableTLS Http
    Http.Open "GET", "https://discord.com/api/v6/WINXP_FOREVER", False
    Http.SetRequestHeader "Content-Type", "application/json"
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    On Error GoTo e
    Http.Send

    On Error Resume Next
    SetFont Me
    Set guild = Nothing
    Token = GetSetting("Discord API Explorer 2", "Authorization", "Token", "")
    Me.Height = 5970
    Me.Show
    'ssTabs_Click
    
    permissions = Array( _
            Array(8, "��� ���� ������", ""), Array(128, "���� ���� ��ȸ", ""), Array(524288, "���� �λ���Ʈ ��ȸ", ""), _
            Array(32, "���� ���� ����", ""), Array(268435456, "���� ����", ""), Array(16, "ä�� ����", ""), _
            Array(2, "�߹�", ""), Array(4, "����", ""), Array(1, "�ʴ�", ""), _
            Array(67108864, "���� ����", ""), Array(134217728, "���� ����", ""), Array(1073741824, "�̸��� ����", ""), _
            Array(536870912, "����ũ ����", ""), Array(1024, "�޽��� �б�", ""), Array(2048, "�޽��� ������", ""), _
            Array(4096, "TTS �޽��� ������", ""), Array(8192, "�޽��� ���� �� ����", ""), Array(16384, "��ũ ����", ""), _
            Array(32768, "ȭ�� �ø���", ""), Array(65536, "���� �޽��� �б�", ""), Array(131072, "��� ���ϱ�", ""), _
            Array(262144, "�缳 �̸��� ���", ""), Array(64, "���� �߰�", ""), Array(1048576, "���� ä�� ����", ""), _
            Array(2097152, "���ϱ�", ""), Array(512, "ī�޶� �� ȭ�� ����", ""), Array(4194304, "���� ����ũ ���Ұ�", ""), _
            Array(8388608, "���� ����Ŀ ���Ұ�", ""), Array(16777216, "���� ä�� ���� �̵�", ""), Array(33554432, "�ڵ� ����ũ ���", ""), _
            Array(256, "�켱 �߾���", "") _
        )
    
    lvFeatures.AddItem "ANIMATED_BANNER"
    lvFeatures.AddItem "ANIMATED_ICON"
    lvFeatures.AddItem "APPLICATION_COMMAND_PERMISSIONS_V2"
    lvFeatures.AddItem "AUTO_MODERATION"
    lvFeatures.AddItem "BANNER"
    lvFeatures.AddItem "COMMUNITY"
    lvFeatures.AddItem "DEVELOPER_SUPPORT_SERVER"
    lvFeatures.AddItem "DISCOVERABLE"
    lvFeatures.AddItem "FEATURABLE"
    lvFeatures.AddItem "INVITES_DISABLED"
    lvFeatures.AddItem "INVITE_SPLASH"
    lvFeatures.AddItem "MEMBER_VERIFICATION_GATE_ENABLED"
    lvFeatures.AddItem "MONETIZATION_ENABLED"
    lvFeatures.AddItem "MORE_STICKERS"
    lvFeatures.AddItem "NEWS"
    lvFeatures.AddItem "PARTNERED"
    lvFeatures.AddItem "PREVIEW_ENABLED"
    lvFeatures.AddItem "ROLE_ICONS"
    lvFeatures.AddItem "TICKETED_EVENTS_ENABLED"
    lvFeatures.AddItem "VANITY_URL"
    lvFeatures.AddItem "VERIFIED"
    lvFeatures.AddItem "VIP_REGIONS"
    lvFeatures.AddItem "WELCOME_SCREEN_ENABLED"
    lvFeatures.AddItem "COMMUNITY"
    lvFeatures.AddItem "INVITES_DISABLED"
    lvFeatures.AddItem "DISCOVERABLE"
    
    cbVerificationLevel.AddItem "0 - ���� ����"
    cbVerificationLevel.AddItem "1 - ���ڿ��� ���� �ʿ�"
    cbVerificationLevel.AddItem "2 - ȸ������ �� 5�� ���"
    cbVerificationLevel.AddItem "3 - ���� ���� �� 10�� ���"
    cbVerificationLevel.AddItem "4 - ��ȭ��ȣ ���� �ʿ�"
    
    cbNotificationLevel.AddItem "0 - ��� �޽���"
    cbNotificationLevel.AddItem "1 - �� �޽�����"
    
    cbFilterLevel.AddItem "0 - ����"
    cbFilterLevel.AddItem "1 - ���� ���� ��� �˿�"
    cbFilterLevel.AddItem "2 - ��� ��� �˿�"
    
    '2295
    lvAuditLogs.ColumnHeaders.Add , "id", "ID", 0, 0
    lvAuditLogs.ColumnHeaders.Add , "num", "#", 245, 1
    lvAuditLogs.ColumnHeaders.Add , "action", "����", 1500, 2
    lvAuditLogs.ColumnHeaders.Add , "executor", "������", 1325, 2
    lvAuditLogs.ColumnHeaders.Add , "target", "���", 1325, 2
    
    
    AuditKeys.Add "nick", "����"
    AuditKeys.Add "communication_disabled_until", "Ÿ�Ӿƿ�"
    AuditKeys.Add "code", "�ʴ� �ڵ�"
    AuditKeys.Add "channel_id", "ä�� ID"
    AuditKeys.Add "inviter_id", "�ʴ��� ID"
    AuditKeys.Add "uses", "��� ��"
    AuditKeys.Add "max_uses", "�ִ� ��� ��"
    AuditKeys.Add "max_age", "�ִ� ����"
    AuditKeys.Add "temporary", "�ӽ� ���"
    AuditKeys.Add "name", "�̸�"
    AuditKeys.Add "type", "����"
    AuditKeys.Add "bitrate", "��Ʈ ���۷� ����"
    AuditKeys.Add "user_limit", "���� ������ ����"
    AuditKeys.Add "permission_overwrites", "ACL"
    AuditKeys.Add "nsfw", "����"
    AuditKeys.Add "rate_limit_per_user", "���ο� ���"
    AuditKeys.Add "flags", "�÷���"
    
    Dim lStyle As Long
    lStyle = SendMessage(lvAuditLogs.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lStyle = lStyle Or LVS_EX_FULLROWSELECT
    Call SendMessage(lvAuditLogs.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)
    
    Exit Sub
    
e:
    If InStr(1, LCase(Err.Description), "security") > 0 Or InStr(1, LCase(Err.Description), "����") > 0 Then
        If MsgBox("���ڵ� API ������ ������ �� �����ϴ�. TLS 1.2�� Ȱ��ȭ�ߴ��� Ȯ���Ͻʽÿ�. Windows XP�� ����ϴ� ���, TLS 1.2 ��ġ�� ��ġ�Ͻʽÿ�. Windows XP�� TLS 1.2 Ȱ��ȭ�� ���õ� �ڼ��� ������ ������ [���]�� �����ʽÿ�.", 16 + vbOKCancel, "���� ���� ����") = vbCancel Then
            Shell "explorer.exe http://web.archive.org/web/20221213130046if_/https://www.emailarchitect.net/easendmail/sdk/html/object_tls12_a.htm"
        End If
        End
    Else
        MsgBox "���ڵ� API ������ ������ �� �����ϴ�. �ٽ� �õ��� �ֽʽÿ�." & vbCrLf & "  " & Err.Number & ": " & Err.Description, 16, "���� ���� ����"
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Http = Nothing
    Set p = Nothing
    Set roles = Nothing
End Sub

Private Sub lvAuditLogChanges_Click()
    On Error GoTo e
    Dim change As Object
    Set change = auditLogChanges(lvAuditLogChanges.ListIndex + 1)
    fAuditLogChangeInfo.Caption = lvAuditLogChanges.Text
    Dim log As Object
    Set log = auditLogMap(lvAuditLogs.SelectedItem.Text)
'    Dim newold As String
'    If ssAuditLogTabs.SelectedItem.Index = 2 Then
'        newold = "new_value"
'    Else
'        newold = "old_value"
'    End If
    
    If change.Exists("old_value") Then
        If IsObject(change("old_value")) Then
            txtOldValue.Text = JSON.toString(change("old_value"))
        Else
            txtOldValue.Text = change("old_value")
        End If
    Else
        txtOldValue.Text = "(�ش� ����)"
    End If
    
    If change.Exists("new_value") Then
        If IsObject(change("new_value")) Then
            txtNewValue.Text = JSON.toString(change("new_value"))
        Else
            txtNewValue.Text = change("new_value")
        End If
    Else
        txtNewValue.Text = "(�ش� ����)"
    End If
    Exit Sub
    
e:
End Sub

Private Sub lvAuditLogs_Click()
    On Error GoTo e
    Dim log As Object
    Set log = auditLogMap(lvAuditLogs.SelectedItem.Text)
    lvAuditLogChanges.Clear
    fAuditLogChangeInfo.Caption = "�ڼ��� ����"
    txtOldValue.Text = ""
    txtNewValue.Text = ""
    Set auditLogChanges = New Dictionary
    Dim keydesc As String
    If Not IsNull(log("changes")) Then
        For I% = 1 To log("changes").count
            If AuditKeys.Exists(log("changes")(I)("key")) Then
                keydesc = AuditKeys(log("changes")(I)("key"))
            Else
                keydesc = log("changes")(I)("key")
            End If
            lvAuditLogChanges.AddItem "#" & I & " - " & keydesc
            auditLogChanges.Add I, log("changes")(I)
        Next I
    End If
    Exit Sub
e:
End Sub

Private Sub lvBans_Click()
    On Error Resume Next
    Dim ban As Object
    Set ban = banMap(lvBans.Text)
    fBanInfo.Caption = ban("user")("username") & "#" & ban("user")("discriminator")
    txtBanReason.Text = ban("reason")
End Sub

Private Sub lvFeatures_ItemCheck(Item As Integer)
    If Not loadingFeatures Then _
        lvFeatures.Selected(Item) = Not lvFeatures.Selected(Item)
End Sub

Private Sub lvInvites_Click()
    On Error Resume Next
    Dim invite As Object
    Set invite = inviteMap(lvInvites.Text)
    lblInviter.Caption = invite("inviter")("username") & "#" & invite("inviter")("discriminator")
    lblInviteChannel.Caption = "#" & invite("channel")("name")
    lblInviteUses.Caption = invite("uses") & "ȸ"
    If invite("max_uses") > 0 Then
        lblInviteUses.Caption = invite("max_uses") & "ȸ �� " & lblInviteUses.Caption & " ���"
    End If
    If IsNull(invite("expires_at")) Then
        lblExpiration.Caption = "����"
    Else
        lblExpiration.Caption = invite("expires_at")
    End If
    chkTemporary.Value = -invite("temporary")
End Sub

Private Sub ssAuditLogTabs_Click()
    lvAuditLogChanges_Click
End Sub

Private Sub ssTabs_Click(PreviousTab As Integer)
    On Error Resume Next
    If (Not invitesLoaded) And ssTabs.TabSel = 5 And (Not (guild Is Nothing)) Then
        invitesLoaded = True
        lvInvites.Clear
        lvInvites.AddItem "(�ҷ����� ��...)"
        Http.Open "GET", "https://discord.com/api/v8/guilds/" & guild("id") & "/invites", True
        Http.SetRequestHeader "Content-Type", "application/json"
        Http.SetRequestHeader "Authorization", Token
        Http.SetRequestHeader "User-Agent", "My XML App V1.0"
        Http.Send
        Http.WaitForResponse 60
        
        If Http.Status >= 400 Then
            MsgBox "������ �����Ͽ� �ʴ��� ����� �ҷ��� �� �����ϴ�.", 16, "����"
            Exit Sub
        End If
        
        Set p = JSON.parse("{""invites"":" & CStr(Http.ResponseText) & "}")
        Set invites = p("invites")
        Set inviteMap = New Dictionary
        lvInvites.Clear
        For I% = 1 To invites.count
            lvInvites.AddItem invites(I)("code")
            inviteMap.Add invites(I)("code"), invites(I)
        Next I
    End If
    
    If (Not bansLoaded) And ssTabs.TabSel = 6 And (Not (guild Is Nothing)) Then
        bansLoaded = True
        lvBans.Clear
        lvBans.AddItem "(�ҷ����� ��...)"
        Http.Open "GET", "https://discord.com/api/v8/guilds/" & guild("id") & "/bans", True
        Http.SetRequestHeader "Content-Type", "application/json"
        Http.SetRequestHeader "Authorization", Token
        Http.SetRequestHeader "User-Agent", "My XML App V1.0"
        Http.Send
        Http.WaitForResponse 60
        
        If Http.Status >= 400 Then
            MsgBox "������ �����Ͽ� �� ����� �ҷ��� �� �����ϴ�. ��� ���� ������ �־�� �մϴ�.", 16, "����"
            Exit Sub
        End If
        
        Set p = JSON.parse("{""bans"":" & CStr(Http.ResponseText) & "}")
        Set bans = p("bans")
        Set banMap = New Dictionary
        lvBans.Clear
        For I% = 1 To bans.count
            lvBans.AddItem bans(I)("user")("username")
            banMap.Add bans(I)("user")("username"), bans(I)
        Next I
    End If
    
    If (Not membersLoaded) And ssTabs.TabSel = 3 And (Not (guild Is Nothing)) Then
        membersLoaded = True
        tvMembers.Nodes.Clear
        tvMembers.Nodes.Add , , "X", "(�ҷ����� ��...)"
        
        If Left$(Token, 4) <> "Bot " Then
            MsgBox "��� ��� ��ȸ�� �� �������θ� �����մϴ�.", 16, "����"
            Exit Sub
        End If
        
        Http.Open "GET", "https://discord.com/api/v8/guilds/" & guild("id") & "/members?limit=50", True
        Http.SetRequestHeader "Content-Type", "application/json"
        Http.SetRequestHeader "Authorization", Token
        Http.SetRequestHeader "User-Agent", "My XML App V1.0"
        Http.Send
        Http.WaitForResponse 60
        
        If Http.Status >= 400 Then
            MsgBox "���ڵ� ������ ��Ż���� ���ø����̼��� ��� ��� ����Ʈ�� Ȱ��ȭ�ϰ� �ٽ� �õ��Ͻʽÿ�.", 16, "����"
            Exit Sub
        End If
        
        Set p = JSON.parse("{""members"":" & CStr(Http.ResponseText) & "}")
        Set members = p("members")
        Set memberMap = New Dictionary
        tvMembers.Nodes.Clear
        Dim rpos%
        Dim roleID$
        Dim rolename$
        For I% = 1 To members.count
            rolename = "[ ���� ���� ]"
            roleID = "0"
            rpos = 0
            For k% = 1 To members(I)("roles").count
                If members(I)("roles")(k) <> guild("id") And roleMap(members(I)("roles")(k))("hoist") And roleMap(members(I)("roles")(k))("position") >= rpos Then
                    rolename = "[ " & roleMap(members(I)("roles")(k))("name") & " ]"
                    roleID = members(I)("roles")(k)
                    rpos = roleMap(members(I)("roles")(k))("position")
                End If
            Next k
            If roleID <> "0" Then
                tvMembers.Nodes.Add , , "R_" & roleID, rolename
                tvMembers.Nodes("R_" & roleID).Expanded = True
            End If
            memberMap.Add CStr(members(I)("user")("id")), members(I)
        Next I
        tvMembers.Nodes.Add , , "R_-1", "[ ���� ���� ]"
        tvMembers.Nodes("X").Expanded = True
        For I% = 1 To members.count
            rolename = "[ ���� ���� ]"
            roleID = "R_-1"
            rpos = 0
            For k% = 1 To members(I)("roles").count
                If members(I)("roles")(k) <> guild("id") And roleMap(members(I)("roles")(k))("hoist") And roleMap(members(I)("roles")(k))("position") >= rpos Then
                    roleID = members(I)("roles")(k)
                    rpos = roleMap(members(I)("roles")(k))("position")
                End If
            Next k
            tvMembers.Nodes.Add "R_" & roleID, tvwChild, CStr(members(I)("user")("id")), members(I)("user")("username")
        Next I
    End If
    
    If (Not auditLogLoaded) And ssTabs.TabSel = 7 And (Not (guild Is Nothing)) Then
        auditLogLoaded = True
        lvAuditLogs.ListItems.Clear
        lvAuditLogs.ListItems.Add , , "0"
        lvAuditLogs.ListItems(1).SubItems(2) = "(�ҷ����� ��...)"
        Http.Open "GET", "https://discord.com/api/v8/guilds/" & guild("id") & "/audit-logs", True
        Http.SetRequestHeader "Content-Type", "application/json"
        Http.SetRequestHeader "Authorization", Token
        Http.SetRequestHeader "User-Agent", "My XML App V1.0"
        Http.Send
        Http.WaitForResponse 60
        
        If Http.Status >= 400 Then
            MsgBox "������ �����Ͽ� ���� �α׸� ��ȸ�� �� �����ϴ�. ���� �α� ���� ������ �־�� �մϴ�.", 16, "����"
            Exit Sub
        End If
        
        Set auditLog = JSON.parse(CStr(Http.ResponseText))
        lvAuditLogs.ListItems.Clear
        Set auditLogs = auditLog("audit_log_entries")
        Set auditLogMap = New Dictionary
        Dim userName As String
        Dim user As Object
        For I% = 1 To auditLogs.count
            lvAuditLogs.ListItems.Add , , auditLogs(I)("id")
            lvAuditLogs.ListItems(I).SubItems(1) = "#" & I
            lvAuditLogs.ListItems(I).SubItems(2) = AuditLogType(auditLogs(I)("action_type"))
            
            userName = ""
            For k% = 1 To auditLog("users").count
                Set user = auditLog("users")(k)
                If user("id") = auditLogs(I)("user_id") Then
                    userName = auditLog("users")(k)("username")
                    Exit For
                End If
            Next k
            If Len(userName) Then
                lvAuditLogs.ListItems(I).SubItems(3) = userName
                auditLogs(I).Add "user", user
            Else
                lvAuditLogs.ListItems(I).SubItems(3) = "#" & auditLogs(I)("user_id")
            End If
            
            If auditLogs(I)("action_type") >= 20 And auditLogs(I)("action_type") < 30 Then
                userName = ""
                For k% = 1 To auditLog("users").count
                    Set user = auditLog("users")(k)
                    If user("id") = auditLogs(I)("target_id") Then
                        userName = auditLog("users")(k)("username")
                        Exit For
                    End If
                Next k
                If Len(userName) Then
                    lvAuditLogs.ListItems(I).SubItems(4) = userName
                    auditLogs(I).Add "target", user
                Else
                    lvAuditLogs.ListItems(I).SubItems(4) = "#" & auditLogs(I)("target_id")
                End If
            ElseIf auditLogs(I)("action_type") >= 10 And auditLogs(I)("action_type") < 20 Then
                If channelMap.Exists(auditLogs(I)("target_id")) Then
                    lvAuditLogs.ListItems(I).SubItems(4) = channelMap(auditLogs(I)("target_id"))("name")
                Else
                    lvAuditLogs.ListItems(I).SubItems(4) = "#" & auditLogs(I)("target_id")
                End If
                auditLogs(I).Add "target", channelMap(auditLogs(I)("target_id"))
            Else
                If auditLogs(I).Exists("target_id") And Len(auditLogs(I)("target_id")) Then
                    lvAuditLogs.ListItems(I).SubItems(4) = "#" & auditLogs(I)("target_id")
                Else
                    lvAuditLogs.ListItems(I).SubItems(4) = "-"
                End If
            End If
            
            auditLogMap.Add auditLogs(I)("id"), auditLogs(I)
        Next I
    End If
End Sub

Private Sub Timer1_Timer()
    On Error GoTo e
    Timer1.Enabled = 0
    Dim pageIdx%
    pageIdx = Timer1.Tag
    ssTabs.Tabs(pageIdx).Selected = -1
    ssTabs_Click
    fTabContents(pageIdx).AutoRedraw = True
    fTabContents(pageIdx).BackColor = RGB(255, 255, 255)
    SetTransparent fTabContents(pageIdx)
    Timer1.Tag = pageIdx + 1
    Exit Sub
e:
    Timer1.Enabled = 0
End Sub

Private Sub tvChannels_Click()
    On Error GoTo e
    Dim user As Object
    
    Dim channel As Object
    Set channel = channelMap(CStr(tvChannels.SelectedItem.key))
    
    Select Case CInt(channel("type"))
        Case 4
            fChannelInfo.Caption = "[ī�װ�] "
        Case 2
            fChannelInfo.Caption = "[��ȭ��] "
        Case 0
            fChannelInfo.Caption = "[ä�ù�] "
        Case 13
            fChannelInfo.Caption = "[��������] "
        Case 5
            fChannelInfo.Caption = "[����] "
        Case Else
            fChannelInfo.Caption = "[ä�� ���� #" & channel("type") & "] "
    End Select
    
    fChannelInfo.Caption = fChannelInfo.Caption & channel("name")
    txtChannelID.Text = channel("id")
    txtPosition.Text = channel("position")
'    If Not IsNull(channel("parent_id")) Then
'        txtParent.Text = channelMap(channel("parent_id"))("name")
'    Else
'        txtParent.Text = "(�ֻ���)"
'    End If
    
    If Not IsNull(channel("topic")) Then
        txtTopic.Text = channel("topic")
    Else
        txtTopic.Text = "(����)"
    End If
    
    chkNSFW.Value = channel("nsfw")
    chkSystemChannel.Value = -((Not IsNull(guild("system_channel_id"))) And (guild("system_channel_id") = channel("id")))
    
    lvPermissionOverwrites.Clear
    txtAllow.Text = ""
    txtDeny.Text = ""
    Set permissionOverwrites = New Dictionary
    On Error Resume Next
    For k% = 1 To channel("permission_overwrites").count
        If channel("permission_overwrites")(k)("type") = 0 Then
            lvPermissionOverwrites.AddItem roleMap(channel("permission_overwrites")(k)("id"))("name")
            permissionOverwrites.Add roleMap(channel("permission_overwrites")(k)("id"))("name"), channel("permission_overwrites")(k)
        Else
            Http.Open "GET", "https://discord.com/api/v8/users/" & channel("permission_overwrites")(k)("id"), True
            Http.SetRequestHeader "Content-Type", "application/json"
            Http.SetRequestHeader "Authorization", Token
            Http.SetRequestHeader "User-Agent", "My XML App V1.0"
            Http.Send
            Http.WaitForResponse 60
            
            Set user = JSON.parse(Http.ResponseText)
            lvPermissionOverwrites.AddItem user("username")
            permissionOverwrites.Add user("username"), channel("permission_overwrites")(k)
            Set user = Nothing
        End If
    Next k
    
    Set channel = Nothing
    Exit Sub
    
e:
    'MsgBox "ä�� ������ �ҷ��� �� �����ϴ�.", 16, "ä��"
End Sub

Private Sub lvPermissionOverwrites_Click()
    On Error GoTo e
    Dim po As Object
    Set po = permissionOverwrites(lvPermissionOverwrites.Text)
    txtAllow.Text = ""
    txtDeny.Text = ""
    For k% = LBound(permissions) To UBound(permissions)
        If ZeroFill(BitAnd(CDec(po("allow")), CDec(permissions(k)(0))), 48) = ZeroFill(toBinary(CDec(permissions(k)(0))), 48) Then
            txtAllow.Text = txtAllow.Text & permissions(k)(1) & " / "
        End If
        
        If ZeroFill(BitAnd(CDec(po("deny")), CDec(permissions(k)(0))), 48) = ZeroFill(toBinary(CDec(permissions(k)(0))), 48) Then
            txtDeny.Text = txtDeny.Text & permissions(k)(1) & " / "
        End If
    Next k
    
    Exit Sub
e:
End Sub

Private Sub lvPermissions_Click()
    On Error Resume Next
    txtPermissionDescription.Text = RolePermissions(lvPermissions.ListIndex)
End Sub

Private Sub lvRoles_Click()
    'On Error GoTo e
    Dim role As Object
    Dim hexColor$
    For I% = 1 To roles.count
        Set role = roles(I)
        If role("name") = lvRoles.Text Then
            fRoleInfo.Caption = "#" & I & " " & lvRoles.Text
            txtRoleID.Text = role("id")
            
            If role("color") Then
                hexColor = ZeroFill(CStr(Hex$(role("color"))), 6)
                lblRoleColor.Caption = "�� #" & hexColor
                lblRoleColor.ForeColor = RGB(CLng("&H" & Mid$(hexColor, 1, 2)), CLng("&H" & Mid$(hexColor, 3, 2)), CLng("&H" & Mid$(hexColor, 5, 2)))
            Else
                lblRoleColor.Caption = "����"
                lblRoleColor.ForeColor = RGB(0, 0, 0)
            End If
            
            chkHoistRole.Value = -role("hoist")
            chkMentionableRole.Value = -role("mentionable")
            fPermissions.Caption = "���� (" & role("permissions") & ")"
            lvPermissions.Clear
            txtPermissionDescription.Text = ""
            For k% = LBound(permissions) To UBound(permissions)
                If ZeroFill(BitAnd(CDec(role("permissions")), CDec(permissions(k)(0))), 48) = ZeroFill(toBinary(CDec(permissions(k)(0))), 48) Then
                    lvPermissions.AddItem permissions(k)(1)
                    RolePermissions(lvPermissions.ListCount - 1) = "���Ѽ�: " & CDec(permissions(k)(0)) & vbCrLf & vbCrLf & permissions(k)(2)
                End If
            Next k
            
            Set role = Nothing
            Exit Sub
        End If
    Next I
    
    'Exit Sub
    
e:
    MsgBox "���� ������ �ҷ��� �� �����ϴ�.", 16, "����"
End Sub

Private Sub tvChannels_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub tvMembers_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub tvMembers_Click()
    If Left$(tvMembers.SelectedItem.key, 2) = "R_" Then Exit Sub
    Dim member As Object
    Dim TopRoleID$
    Dim hexColor$
    Dim role As Object
    Set member = memberMap(tvMembers.SelectedItem.key)
    txtUserTag.Text = member("user")("username") & "#" & member("user")("discriminator")
    txtUserID.Text = member("user")("id")
    
    TopRoleID = Right$(tvMembers.SelectedItem.Parent.key, Len(tvMembers.SelectedItem.Parent.key) - 2)
    If IsNumeric(TopRoleID) Then
        Set role = roleMap(TopRoleID)
        If role("color") Then
            hexColor = ZeroFill(CStr(Hex$(role("color"))), 6)
            lblMemberRole.Caption = "�� " & role("name")
            lblMemberRole.ForeColor = RGB(CLng("&H" & Mid$(hexColor, 1, 2)), CLng("&H" & Mid$(hexColor, 3, 2)), CLng("&H" & Mid$(hexColor, 5, 2)))
        Else
            lblMemberRole.Caption = "����"
            lblMemberRole.ForeColor = RGB(0, 0, 0)
        End If
    End If
    Label37.Caption = member("permissions")
End Sub
