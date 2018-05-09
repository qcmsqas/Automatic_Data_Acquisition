VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "MPC08控制程序"
   ClientHeight    =   14790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   14790
   ScaleWidth      =   14760
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame11 
      Caption         =   "谱线设定"
      Height          =   1815
      Left            =   360
      TabIndex        =   183
      Top             =   11520
      Width           =   5055
      Begin VB.CommandButton CmmdSetSpecFreshA 
         Caption         =   "刷新"
         Height          =   495
         Left            =   3360
         TabIndex        =   189
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton CmmdSetSpecSureA 
         Caption         =   "应用"
         Height          =   495
         Left            =   3360
         TabIndex        =   188
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtStopFreqA 
         Height          =   300
         Left            =   1320
         TabIndex        =   187
         Text            =   "0"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtNumOfPointsA 
         Height          =   300
         Left            =   1320
         TabIndex        =   186
         Text            =   "0"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtStartFreqA 
         Height          =   300
         Left            =   1320
         TabIndex        =   185
         Text            =   "0"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtRateFreqA 
         Height          =   300
         Left            =   1320
         TabIndex        =   184
         Text            =   "0"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "取样数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   193
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "终止频率"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   192
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "起始频率"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   191
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "合并度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   190
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame fileinfos 
      Caption         =   "文件信息"
      Height          =   1335
      Left            =   360
      TabIndex        =   179
      Top             =   9960
      Width           =   5055
      Begin VB.TextBox TxtFileLoad 
         Height          =   375
         Left            =   1080
         TabIndex        =   180
         Text            =   "Text1"
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label30 
         Caption         =   "文件位置"
         Height          =   375
         Left            =   240
         TabIndex        =   181
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox PicErrPrinter 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Index           =   0
      Left            =   360
      ScaleHeight     =   1395
      ScaleWidth      =   4995
      TabIndex        =   120
      Top             =   8520
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      Caption         =   "状态显示"
      Height          =   2295
      Left            =   360
      TabIndex        =   66
      Top             =   6000
      Width           =   5055
      Begin VB.Timer TimerPoint 
         Interval        =   100
         Left            =   0
         Top             =   1920
      End
      Begin VB.Label LblSpeedUnitZ 
         Caption         =   "mm/s"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   206
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label LblSpeedNumZ 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   205
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label LblSpeedZ 
         Caption         =   "Z速度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   204
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label LblPosiUnitZ 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   203
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label LblPosiNumZ 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   202
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label LblPosiZ 
         Alignment       =   2  'Center
         Caption         =   "Z位置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   201
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label LblPosiX 
         Caption         =   "X位置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   78
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblPosiY 
         Caption         =   "Y位置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   77
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblPosiNumX 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   76
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblPosiNumY 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   75
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label C 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   74
         Top             =   480
         Width           =   255
      End
      Begin VB.Label LblPosiUnitY 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   73
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label LblSpeedX 
         Caption         =   "X速度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   72
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblSpeedY 
         Caption         =   "Y速度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   71
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblSpeedNumX 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   70
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblSpeedNumY 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   69
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblSpeedUnitX 
         Caption         =   "mm/s"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   68
         Top             =   480
         Width           =   495
      End
      Begin VB.Label LblSpeedUnitY 
         Caption         =   "mm/s"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   67
         Top             =   1080
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   25000
      Left            =   120
      TabIndex        =   144
      Top             =   120
      Width           =   14455
      _ExtentX        =   25506
      _ExtentY        =   44106
      _Version        =   393216
      Tabs            =   7
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "设定及测试"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameSetPoint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmmdStart"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmmdFend"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmmdResetZero"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmmdLend"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrameSpeedSet"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FramePosiSet"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "手动扫Marker"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CmmdLendM"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "CmmdResetZeroM"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmmdFendM"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CmmdStartM"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "FrmScanM"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "自动扫Marker"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "TimerA"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "TimerAvgA"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "TimerDelayA"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "手动扫谱"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TimerAvgB"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "TimerDelayB"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame10"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame7"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Command5"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Command4"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "CmmdFendS"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "CmmdStartS"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "自动扫谱"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TimerB"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame8"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame6"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "二维扫谱"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame12"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame13"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "TimerSB"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "三维扫谱"
      TabPicture(6)   =   "frmMain.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame14"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame15"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame15 
         Caption         =   "数据采集"
         Height          =   11535
         Left            =   -69480
         TabIndex        =   142
         Top             =   1080
         Width           =   8775
         Begin RichTextLib.RichTextBox RichTextBox2 
            Height          =   10215
            Left            =   360
            TabIndex        =   155
            Top             =   1080
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   18018
            _Version        =   393217
            Enabled         =   -1  'True
            RightMargin     =   1e7
            OLEDropMode     =   0
            TextRTF         =   $"frmMain.frx":00C4
         End
         Begin VB.CommandButton Command13 
            Caption         =   "保存数据"
            Height          =   615
            Left            =   5880
            TabIndex        =   146
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command11 
            Caption         =   "清空数据"
            Height          =   615
            Left            =   4440
            TabIndex        =   147
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command10 
            Caption         =   "中止扫描"
            Height          =   615
            Left            =   3000
            TabIndex        =   148
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "开始扫描"
            Height          =   615
            Left            =   1560
            TabIndex        =   149
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "连接网分"
            Height          =   615
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "运动设定"
         Height          =   2895
         Left            =   -74800
         TabIndex        =   90
         Top             =   1080
         Width           =   5055
         Begin VB.TextBox Text7 
            Height          =   300
            Left            =   1800
            TabIndex        =   177
            Text            =   "Text7"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox Text6 
            Height          =   300
            Left            =   840
            TabIndex        =   163
            Text            =   "Text6"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   3960
            TabIndex        =   160
            Text            =   "Combo1"
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Height          =   270
            Left            =   3600
            TabIndex        =   194
            Text            =   "Text5"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox Text4 
            Height          =   300
            Left            =   2880
            TabIndex        =   167
            Text            =   "Text4"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Left            =   2400
            TabIndex        =   165
            Text            =   "Text3"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Text2 
            Height          =   300
            Left            =   1800
            TabIndex        =   175
            Text            =   "Text2"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   840
            TabIndex        =   158
            Text            =   "Text1"
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "Y方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   172
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "X方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   170
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Timer TimerSB 
         Interval        =   200
         Left            =   -70080
         Top             =   4320
      End
      Begin VB.Frame Frame13 
         Caption         =   "运动设定"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   157
         Top             =   1080
         Width           =   5055
         Begin VB.TextBox TxtIBSB 
            Height          =   300
            Left            =   3480
            TabIndex        =   182
            Text            =   "0"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox TxtStepAYSB 
            Height          =   300
            Left            =   1800
            TabIndex        =   178
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox TxtStepAXSB 
            Height          =   300
            Left            =   1800
            TabIndex        =   176
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtStepBYSB 
            Height          =   300
            Left            =   2880
            TabIndex        =   169
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox TxtStepBXSB 
            Height          =   300
            Left            =   2880
            TabIndex        =   168
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtIASB 
            Height          =   300
            Left            =   2400
            TabIndex        =   166
            Text            =   "0"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox TxtStartYSB 
            Height          =   300
            Left            =   840
            TabIndex        =   164
            Text            =   "0"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox CmblUnitYSB 
            Height          =   300
            ItemData        =   "frmMain.frx":0153
            Left            =   3960
            List            =   "frmMain.frx":0160
            Style           =   2  'Dropdown List
            TabIndex        =   162
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox CmblUnitXSB 
            Height          =   300
            ItemData        =   "frmMain.frx":0173
            Left            =   3960
            List            =   "frmMain.frx":0180
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox TxtStartXSB 
            Height          =   300
            Left            =   840
            TabIndex        =   159
            Text            =   "0"
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            Caption         =   "起点 步长1步数1 步长2步数2  单位"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            TabIndex        =   174
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            Caption         =   "Y方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   173
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            Caption         =   "X方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   171
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "数据采集"
         Height          =   11535
         Left            =   -69480
         TabIndex        =   143
         Top             =   1080
         Width           =   8775
         Begin VB.CommandButton CmmdStopScanSB 
            Caption         =   "中止扫描"
            Height          =   615
            Left            =   3000
            TabIndex        =   154
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command9 
            Caption         =   "清空数据"
            Height          =   615
            Left            =   4440
            TabIndex        =   153
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command8 
            Caption         =   "保存数据"
            Height          =   615
            Left            =   5880
            TabIndex        =   152
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdConnectPNASB 
            Caption         =   "连接网分"
            Height          =   615
            Left            =   120
            TabIndex        =   151
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdScanSB 
            Caption         =   "开始扫描"
            Height          =   615
            Left            =   1560
            TabIndex        =   145
            Top             =   360
            Width           =   1335
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   10215
            Left            =   120
            TabIndex        =   156
            Top             =   1080
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   18018
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   1e7
            TextRTF         =   $"frmMain.frx":0193
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   855
         Left            =   5400
         TabIndex        =   139
         Top             =   8520
         Width           =   1455
      End
      Begin VB.Timer TimerB 
         Interval        =   200
         Left            =   -70080
         Top             =   6720
      End
      Begin VB.Timer TimerAvgB 
         Left            =   -70080
         Top             =   6720
      End
      Begin VB.Timer TimerDelayB 
         Left            =   -70440
         Top             =   6720
      End
      Begin VB.Frame Frame10 
         Caption         =   "谱线设定"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   130
         Top             =   3360
         Width           =   5055
         Begin VB.TextBox TxtRateFreq 
            Height          =   300
            Left            =   1320
            TabIndex        =   141
            Text            =   "0"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton CmmdSetSpecFresh 
            Caption         =   "刷新"
            Height          =   495
            Left            =   3360
            TabIndex        =   138
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton CmmdSetSpecSure 
            Caption         =   "应用"
            Height          =   495
            Left            =   3360
            TabIndex        =   137
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtStopFreq 
            Height          =   300
            Left            =   1320
            TabIndex        =   133
            Text            =   "0"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox TxtNumOfPoints 
            Height          =   300
            Left            =   1320
            TabIndex        =   132
            Text            =   "0"
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox TxtStartFreq 
            Height          =   300
            Left            =   1320
            TabIndex        =   131
            Text            =   "0"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label28 
            Caption         =   "合并度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "取样数"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label23 
            Caption         =   "终止频率"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "起始频率"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "测量设定"
         Height          =   1935
         Left            =   5520
         TabIndex        =   116
         Top             =   3960
         Width           =   5055
         Begin VB.CommandButton CmmdGetDataCancel 
            Caption         =   "取消"
            Height          =   495
            Left            =   3840
            TabIndex        =   125
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton CmmdGetDataSure 
            Caption         =   "应用"
            Height          =   495
            Left            =   3840
            TabIndex        =   124
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TxtGetDataAvgDelayTime 
            Height          =   300
            Left            =   2040
            TabIndex        =   119
            Text            =   "0"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox TxtGetDataAvgNum 
            Height          =   300
            Left            =   2040
            TabIndex        =   118
            Text            =   "0"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox TxtGetDataDelayTime 
            Height          =   300
            Left            =   2040
            TabIndex        =   117
            Text            =   "0"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "ms"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   129
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label20 
            Caption         =   "ms"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   128
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "ms"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   126
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "采样间隔时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   123
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "求均值采样次数"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   122
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "网分稳定时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   121
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "数据采集"
         Height          =   11535
         Left            =   -69480
         TabIndex        =   109
         Top             =   1080
         Width           =   8775
         Begin VB.CommandButton CmmdScanSA 
            Caption         =   "开始扫描"
            Height          =   615
            Left            =   1560
            TabIndex        =   114
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdConnectPNASA 
            Caption         =   "连接网分"
            Height          =   615
            Left            =   120
            TabIndex        =   113
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command12 
            Caption         =   "保存数据"
            Height          =   615
            Left            =   5880
            TabIndex        =   112
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdResetSA 
            Caption         =   "清空数据"
            Height          =   615
            Left            =   4440
            TabIndex        =   111
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdStopScanSA 
            Caption         =   "中止扫描"
            Height          =   615
            Left            =   3000
            TabIndex        =   110
            Top             =   360
            Width           =   1335
         End
         Begin RichTextLib.RichTextBox RichDataSA 
            Height          =   10215
            Left            =   120
            TabIndex        =   115
            Top             =   1200
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   18018
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   1e7
            TextRTF         =   $"frmMain.frx":0222
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "数据采集"
         Height          =   11535
         Left            =   -69480
         TabIndex        =   103
         Top             =   1080
         Width           =   8775
         Begin VB.CommandButton CmmdGetDataS 
            Caption         =   "读取数据"
            Height          =   615
            Left            =   1920
            TabIndex        =   107
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton CmmdDeletDataS 
            Caption         =   "清空数据"
            Height          =   615
            Left            =   3600
            TabIndex        =   106
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   "保存数据"
            Height          =   615
            Left            =   5280
            TabIndex        =   105
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton CmmdConnectPNAS 
            Caption         =   "连接网分"
            Height          =   615
            Left            =   240
            TabIndex        =   104
            Top             =   360
            Width           =   1455
         End
         Begin RichTextLib.RichTextBox RichDataS 
            Height          =   10215
            Left            =   120
            TabIndex        =   108
            Top             =   1200
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   18018
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   1e7
            TextRTF         =   $"frmMain.frx":02B1
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "运动设定"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   91
         Top             =   1080
         Width           =   5055
         Begin VB.TextBox TxtStartXSA 
            Height          =   300
            Left            =   840
            TabIndex        =   99
            Text            =   "0"
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox CmblUnitXSA 
            Height          =   300
            ItemData        =   "frmMain.frx":0340
            Left            =   3960
            List            =   "frmMain.frx":034D
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox CmblUnitYSA 
            Height          =   300
            ItemData        =   "frmMain.frx":0360
            Left            =   3960
            List            =   "frmMain.frx":036D
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox TxtStartYSA 
            Height          =   300
            Left            =   840
            TabIndex        =   96
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TxtendXSA 
            Height          =   300
            Left            =   2040
            TabIndex        =   95
            Text            =   "0"
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TxtendYSA 
            Height          =   300
            Left            =   2040
            TabIndex        =   94
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TxtStepXSA 
            Height          =   300
            Left            =   3240
            TabIndex        =   93
            Text            =   "0"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TxtStepYSA 
            Height          =   300
            Left            =   3240
            TabIndex        =   92
            Text            =   "0"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "X方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   102
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Y方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   101
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "起点      终点     步长   单位"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            TabIndex        =   100
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "运动设定"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   83
         Top             =   1080
         Width           =   5055
         Begin VB.TextBox TxtLenYS 
            Height          =   300
            Left            =   1560
            TabIndex        =   87
            Text            =   "0"
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox CmblUnitYS 
            Height          =   300
            ItemData        =   "frmMain.frx":0380
            Left            =   3360
            List            =   "frmMain.frx":038D
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox CmblUnitXS 
            Height          =   300
            ItemData        =   "frmMain.frx":03A0
            Left            =   3360
            List            =   "frmMain.frx":03AD
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxtLenXS 
            Height          =   300
            Left            =   1560
            TabIndex        =   84
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Y方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   89
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "X方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   88
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "缓停"
         Height          =   495
         Left            =   -72120
         TabIndex        =   82
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "全部置零"
         Height          =   495
         Left            =   -70800
         TabIndex        =   81
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton CmmdFendS 
         Caption         =   "急停"
         Height          =   495
         Left            =   -73440
         TabIndex        =   80
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton CmmdStartS 
         Caption         =   "启动"
         Height          =   495
         Left            =   -74760
         TabIndex        =   79
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Timer TimerDelayA 
         Interval        =   5000
         Left            =   -70440
         Top             =   2880
      End
      Begin VB.Timer TimerAvgA 
         Interval        =   50
         Left            =   -70080
         Top             =   2880
      End
      Begin VB.Timer TimerA 
         Interval        =   200
         Left            =   -74880
         Top             =   2880
      End
      Begin VB.Frame Frame5 
         Caption         =   "运动设定"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   49
         Top             =   1080
         Width           =   5055
         Begin VB.TextBox TxtStepYA 
            Height          =   300
            Left            =   3240
            TabIndex        =   56
            Text            =   "0"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtStepXA 
            Height          =   300
            Left            =   3240
            TabIndex        =   52
            Text            =   "0"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TxtendYA 
            Height          =   300
            Left            =   2040
            TabIndex        =   55
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TxtendXA 
            Height          =   300
            Left            =   2040
            TabIndex        =   51
            Text            =   "0"
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TxtStartYA 
            Height          =   300
            Left            =   840
            TabIndex        =   54
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox CmblUnitYMA 
            Height          =   300
            ItemData        =   "frmMain.frx":03C0
            Left            =   3960
            List            =   "frmMain.frx":03CD
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox CmblUnitXMA 
            Height          =   300
            ItemData        =   "frmMain.frx":03E0
            Left            =   3960
            List            =   "frmMain.frx":03ED
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox TxtStartXA 
            Height          =   300
            Left            =   840
            TabIndex        =   50
            Text            =   "0"
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "起点      终点     步长   单位"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            TabIndex        =   60
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Y方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   59
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "X方向"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "数据采集"
         Height          =   11535
         Left            =   -69480
         TabIndex        =   47
         Top             =   1080
         Width           =   8775
         Begin VB.CommandButton CmmdStopScanA 
            Caption         =   "中止扫描"
            Height          =   615
            Left            =   3000
            TabIndex        =   65
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdResetA 
            Caption         =   "清空数据"
            Height          =   615
            Left            =   4440
            TabIndex        =   64
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdSaveA 
            Caption         =   "保存数据"
            Height          =   615
            Left            =   5880
            TabIndex        =   63
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdConnectPNAMA 
            Caption         =   "连接网分"
            Height          =   615
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton CmmdScanA 
            Caption         =   "开始扫描"
            Height          =   615
            Left            =   1560
            TabIndex        =   61
            Top             =   360
            Width           =   1335
         End
         Begin RichTextLib.RichTextBox RichDataA 
            Height          =   10215
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   18018
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   1.00000e5
            TextRTF         =   $"frmMain.frx":0400
         End
      End
      Begin VB.Frame FrmScanM 
         Caption         =   "数据采集"
         Height          =   11535
         Left            =   -69480
         TabIndex        =   41
         Top             =   1080
         Width           =   8775
         Begin VB.CommandButton CmmdConnectPNAM 
            Caption         =   "连接网分"
            Height          =   615
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "保存数据"
            Height          =   615
            Left            =   5280
            TabIndex        =   45
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton CmmdDeletData 
            Caption         =   "清空数据"
            Height          =   615
            Left            =   3600
            TabIndex        =   44
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton CmmdGetData 
            Caption         =   "读取数据"
            Height          =   615
            Left            =   1920
            TabIndex        =   43
            Top             =   360
            Width           =   1455
         End
         Begin RichTextLib.RichTextBox RichDataM 
            Height          =   10215
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   18018
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   1.00000e5
            TextRTF         =   $"frmMain.frx":048F
         End
      End
      Begin VB.CommandButton CmmdStartM 
         Caption         =   "启动"
         Height          =   495
         Left            =   -74760
         TabIndex        =   40
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton CmmdFendM 
         Caption         =   "急停"
         Height          =   495
         Left            =   -73440
         TabIndex        =   39
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton CmmdResetZeroM 
         Caption         =   "全部置零"
         Height          =   495
         Left            =   -70800
         TabIndex        =   38
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton CmmdLendM 
         Caption         =   "缓停"
         Height          =   495
         Left            =   -72120
         TabIndex        =   37
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "运动设定"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   30
         Top             =   1080
         Width           =   5055
         Begin VB.TextBox TxtLenXM 
            Height          =   300
            Left            =   1560
            TabIndex        =   34
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox CmblUnitXM 
            Height          =   300
            ItemData        =   "frmMain.frx":051E
            Left            =   3360
            List            =   "frmMain.frx":052B
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox CmblUnitYM 
            Height          =   300
            ItemData        =   "frmMain.frx":053E
            Left            =   3360
            List            =   "frmMain.frx":054B
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox TxtLenYM 
            Height          =   300
            Left            =   1560
            TabIndex        =   31
            Text            =   "0"
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "X方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Y方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   35
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame FramePosiSet 
         Caption         =   "重设位置"
         Height          =   2115
         Left            =   240
         TabIndex        =   25
         Top             =   3720
         Width           =   5055
         Begin VB.TextBox TxtPosiSetZ 
            Height          =   300
            Left            =   1560
            TabIndex        =   199
            Text            =   "0"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton CmmdSetPosi 
            Caption         =   "设置为当前位置"
            Height          =   975
            Left            =   3240
            TabIndex        =   18
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TxtPosiSetY 
            Height          =   300
            Left            =   1560
            TabIndex        =   17
            Text            =   "0"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox TxtPosiSetX 
            Height          =   300
            Left            =   1560
            TabIndex        =   16
            Text            =   "0"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   200
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "Z方向位置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   198
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   29
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   28
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Y方向位置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   27
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X方向位置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   26
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame FrameSpeedSet 
         Caption         =   "速度设定"
         Height          =   2655
         Left            =   5520
         TabIndex        =   22
         Top             =   1080
         Width           =   5055
         Begin VB.ComboBox CmbSpeedUnitZ 
            Height          =   300
            ItemData        =   "frmMain.frx":055E
            Left            =   3600
            List            =   "frmMain.frx":056B
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox TxtSpeedNumZ 
            Height          =   300
            Left            =   1800
            TabIndex        =   208
            Text            =   "0"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton CmmdSpeedCancel 
            Caption         =   "取消"
            Height          =   495
            Left            =   1320
            TabIndex        =   13
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton CmmdSpeedSaveAs 
            Caption         =   "另存为"
            Height          =   495
            Left            =   3720
            TabIndex        =   15
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton CmmdSpeedSave 
            Caption         =   "保存为默认"
            Height          =   495
            Left            =   2520
            TabIndex        =   14
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton CmmdSpeedSure 
            Caption         =   "应用"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox TxtSpeedNumX 
            Height          =   300
            Left            =   1800
            TabIndex        =   8
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox CmbSpeedUnitX 
            Height          =   300
            ItemData        =   "frmMain.frx":0584
            Left            =   3600
            List            =   "frmMain.frx":0591
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox CmbSpeedUnitY 
            Height          =   300
            ItemData        =   "frmMain.frx":05AA
            Left            =   3600
            List            =   "frmMain.frx":05B7
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox TxtSpeedNumY 
            Height          =   300
            Left            =   1800
            TabIndex        =   10
            Text            =   "0"
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label LblSpeedSetZ 
            Alignment       =   2  'Center
            Caption         =   "Z方向速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   600
            TabIndex        =   207
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label LblSpeedSetX 
            Alignment       =   2  'Center
            Caption         =   "X方向速度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   600
            TabIndex        =   24
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label LblSpeedSetY 
            Alignment       =   2  'Center
            Caption         =   "Y方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   600
            TabIndex        =   23
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.CommandButton CmmdLend 
         Caption         =   "缓停"
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton CmmdResetZero 
         Caption         =   "全部置零"
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton CmmdFend 
         Caption         =   "急停"
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton CmmdStart 
         Caption         =   "启动"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Frame FrameSetPoint 
         Caption         =   "运动设定"
         Height          =   1935
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   5055
         Begin VB.ComboBox CmblUnitZ 
            Height          =   300
            ItemData        =   "frmMain.frx":05D0
            Left            =   3360
            List            =   "frmMain.frx":05DD
            Style           =   2  'Dropdown List
            TabIndex        =   197
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox TxtLenZ 
            Height          =   300
            Left            =   1560
            TabIndex        =   196
            Text            =   "0"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Timer TimerFend 
            Interval        =   200
            Left            =   0
            Top             =   1560
         End
         Begin VB.TextBox TxtLenY 
            Height          =   300
            Left            =   1560
            TabIndex        =   2
            Text            =   "0"
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox CmblUnitY 
            Height          =   300
            ItemData        =   "frmMain.frx":05F0
            Left            =   3360
            List            =   "frmMain.frx":05FD
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox CmblUnitX 
            Height          =   300
            ItemData        =   "frmMain.frx":0610
            Left            =   3360
            List            =   "frmMain.frx":061D
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxtLenX 
            Height          =   300
            Left            =   1560
            TabIndex        =   0
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label LblZRun 
            Alignment       =   2  'Center
            Caption         =   "Z方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   195
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label LblYRun 
            Alignment       =   2  'Center
            Caption         =   "Y方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   21
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label LblXRun 
            Alignment       =   2  'Center
            Caption         =   "X方向运动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Label Label19 
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   127
      Top             =   6240
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmblUnitXA_LostFocus()
    CmblUnitX.ListIndex = CmblUnitXA.ListIndex
    Call CmblUnitX_LostFocus
End Sub
Private Sub CmblUnitYA_LostFocus()
    CmblUnitY.ListIndex = CmblUnitYA.ListIndex
    Call CmblUnitY_LostFocus
End Sub

Private Sub CmblUnitXMA_LostFocus()
    CmblUnitX.ListIndex = CmblUnitXMA.ListIndex
    Call CmblUnitX_LostFocus
End Sub

Private Sub CmblUnitXSB_LostFocus()
    CmblUnitX.ListIndex = CmblUnitXSB.ListIndex
    Call CmblUnitX_LostFocus
End Sub
Private Sub CmblUnitYSB_LostFocus()
    CmblUnitY.ListIndex = CmblUnitYSB.ListIndex
    Call CmblUnitY_LostFocus
End Sub

Private Sub CmblUnitYMA_LostFocus()
    CmblUnitY.ListIndex = CmblUnitYMA.ListIndex
    Call CmblUnitY_LostFocus
End Sub
Private Sub CmblUnitXSA_LostFocus()
    CmblUnitX.ListIndex = CmblUnitXSA.ListIndex
    Call CmblUnitX_LostFocus
End Sub
Private Sub CmblUnitYSA_LostFocus()
    CmblUnitY.ListIndex = CmblUnitYSA.ListIndex
    Call CmblUnitY_LostFocus
End Sub

Private Sub CmblUnitXS_LostFocus()
    CmblUnitX.ListIndex = CmblUnitXS.ListIndex
    Call CmblUnitX_LostFocus
End Sub
Private Sub CmblUnitYS_LostFocus()
    CmblUnitY.ListIndex = CmblUnitYS.ListIndex
    Call CmblUnitY_LostFocus
End Sub
Private Sub CmblUnitX_LostFocus()
    CmblUnitXMA.ListIndex = CmblUnitX.ListIndex
    CmblUnitXM.ListIndex = CmblUnitX.ListIndex
    CmblUnitXSA.ListIndex = CmblUnitX.ListIndex
    CmblUnitXSB.ListIndex = CmblUnitX.ListIndex
    CmblUnitXS.ListIndex = CmblUnitX.ListIndex
    
End Sub
Private Sub CmblUnitY_LostFocus()
    CmblUnitYMA.ListIndex = CmblUnitY.ListIndex
    CmblUnitYM.ListIndex = CmblUnitY.ListIndex
    CmblUnitYSA.ListIndex = CmblUnitY.ListIndex
    CmblUnitYSB.ListIndex = CmblUnitY.ListIndex
    CmblUnitYS.ListIndex = CmblUnitY.ListIndex
End Sub

Private Sub CmblUnitZ_Change()
    CmblUnitZMA.ListIndex = CmblUnitZ.ListIndex
    CmblUnitZM.ListIndex = CmblUnitZ.ListIndex
    CmblUnitZSA.ListIndex = CmblUnitZ.ListIndex
    CmblUnitZSB.ListIndex = CmblUnitZ.ListIndex
    CmblUnitZS.ListIndex = CmblUnitZ.ListIndex
End Sub

Private Sub CmbSpeedUnitX_Click()
    If CmbSpeedUnitX.ListIndex = 0 Then
        TxtSpeedNumX.Text = mmConSpeedX
    ElseIf CmbSpeedUnitX.ListIndex = 1 Then
        TxtSpeedNumX.Text = (mmConSpeedX / 10)
    Else
        TxtSpeedNumX.Text = pConSpeedX
    End If
End Sub


Private Sub CmbSpeedUnitY_Click()
    If CmbSpeedUnitY.ListIndex = 0 Then
        TxtSpeedNumY.Text = mmConSpeedY
    ElseIf CmbSpeedUnitY.ListIndex = 1 Then
        TxtSpeedNumY.Text = (mmConSpeedY / 10)
    Else
        TxtSpeedNumY.Text = pConSpeedY
    End If
End Sub

Private Sub CmmdConnectPNAM_Click()
    On Error GoTo err3

    StartNA
    connectPNAMark = True
    PicErrPrinter(0).Print "连接网分成功！"
    

    If connectPNAMark = True Then
        startFreq = ch.StartFrequency
        stopFreq = ch.StopFrequency
        numOfPoints = ch.NumberOfPoints
        TxtStartFreq.Text = startFreq
        TxtStopFreq.Text = stopFreq
        TxtNumOfPoints.Text = numOfPoints
        TxtStartFreqA.Text = startFreq
        TxtStopFreqA.Text = stopFreq
        TxtNumOfPointsA.Text = numOfPoints
        

        RichDataS.Text = "x/mm" & Chr(9) & "y/mm" & Chr(9) & "Frequency/Hz" & Chr(9) & "linmag" & Chr(9) & "logmag" & Chr(9) & "phase" & Chr(9) & "real" & Chr(9) & "img"
        sweepType = ch.sweepType
        If sweepType = 0 Then
            If numOfPoints > 1 Then
                For tmpi = 0 To numOfPoints - 1
                    freqx = startFreq + ((stopFreq - startFreq) * tmpi / (numOfPoints - 1))
                    'RichDataS.Text = RichDataS.Text & Chr(9) & Int(freqx)
                    freqss(tmpi) = Int(freqx)
                Next tmpi
            End If
            If numOfPoints = 1 Then
                'RichDataS.Text = RichDataS.Text & Chr(9) & Int(startFreq)
                freqss(0) = Int(startFreq)
            End If
        End If
        If sweepType = 1 Then
            If numOfPoints > 1 Then
                For tmpi = 0 To numOfPoints - 1
                    freqx = Exp(Log(startFreq) + (Log(stopFreq) - Log(startFreq)) * tmpi / (numOfPoints - 1))
                    'RichDataS.Text = RichDataS.Text & Chr(9) & Int(freqx)
                    freqss(tmpi) = Int(freqx)
                Next tmpi
            End If
            If numOfPoints = 1 Then
                'RichDataS.Text = RichDataS.Text & Chr(9) & Int(startFreq)
                freqss(0) = Int(startFreq)
            End If
        End If
        RichDataSA.Text = RichDataS.Text
    Else
        TxtStartFreq.Text = "0"
        TxtStopFreq.Text = "0"
        TxtNumOfPoints.Text = "0"
        TxtStartFreqA.Text = "0"
        TxtStopFreqA.Text = "0"
        TxtNumOfPointsA.Text = "0"
    End If
    
    Exit Sub
err3:
    MsgBox "请先开启网分！"

End Sub

Private Sub CmmdConnectPNAMA_Click()
    Call CmmdConnectPNAM_Click
End Sub

Private Sub CmmdConnectPNAS_Click()
    Call CmmdConnectPNAM_Click
End Sub

Private Sub CmmdConnectPNASB_Click()
    Call CmmdConnectPNAM_Click
End Sub

Private Sub CmmdDeletData_Click()
    RichDataM.Text = "x/mm" & Chr(9) & "y/mm" & Chr(9) & "linmag" & Chr(9) & "logmag" & Chr(9) & "phase"
    RichDataA.Text = "x/mm" & Chr(9) & "y/mm" & Chr(9) & "linmag" & Chr(9) & "logmag" & Chr(9) & "phase"
End Sub

Private Sub CmmdDeletDataS_Click()
    Call CmmdSetSpecFresh_Click
End Sub

Private Sub CmmdFend_Click()
    sudden_stop 1
    sudden_stop 2
    TimerFend.Enabled = True
End Sub

Private Sub CmmdResetM_Click()

End Sub

Private Sub CmmdFendM_Click()
    Call CmmdFend_Click
End Sub

Private Sub CmmdGetData_Click()
    If connectPNAMark = False Then
        MsgBox "请先连接网分！"
        Exit Sub
    End If
    
    linMagMarker = meas.marker(1).Value(naMarkerFormat_LinMag)

    logMagMarker = meas.marker(1).Value(naMarkerFormat_LogMag)

    phaseMarker = meas.marker(1).Value(naMarkerFormat_Phase)
    
    outPosiX = P2mmX(pPosiX) + mmNPosiX
    outPosiY = P2mmY(pPosiY) + mmNPosiY
    IAvg = 0
    TimerDelayA.Enabled = True
    CmmdGetData.Enabled = False
End Sub

Private Sub CmmdGetDataSure_Click()
    Dim tmpDelay As Integer
    Dim tmpAvgNum As Integer
    Dim tmpAvgDelayTime As Integer
    
    PicErrPrinter(0).Cls
    If Not IsInt(TxtGetDataDelayTime.Text) Then
        PicErrPrinter(0).Print "网分稳定时间错误！必须为整数！"
        TxtGetDataDelayTime.SelStart = 0
        TxtGetDataDelayTime.SelLength = Len(TxtGetDataDelayTime.Text)
        TxtGetDataDelayTime.SetFocus
        Exit Sub
    End If
    
    If Not IsInt(TxtGetDataAvgNum.Text) Then
        PicErrPrinter(0).Print "求均值采样次数错误！必须为整数！"
        TxtGetDataAvgNum.SelStart = 0
        TxtGetDataAvgNum.SelLength = Len(TxtGetDataAvgNum.Text)
        TxtGetDataAvgNum.SetFocus
        Exit Sub
    End If
    
    If Not IsInt(TxtGetDataAvgDelayTime.Text) Then
        PicErrPrinter(0).Print "采样间隔时间错误！必须为整数！"
        TxtGetDataAvgDelayTime.SelStart = 0
        TxtGetDataAvgDelayTime.SelLength = Len(TxtGetDataAvgDelayTime.Text)
        TxtGetDataAvgDelayTime.SetFocus
        Exit Sub
    End If
    getDataDelayTime = Int(Val(TxtGetDataDelayTime.Text))
    getDataAvgNum = Int(Val(TxtGetDataAvgNum.Text))
    getDataAvgDelayTime = Int(Val(TxtGetDataAvgDelayTime.Text))
    
    TimerAvgA.Interval = getDataAvgDelayTime
    TimerDelayA.Interval = getDataDelayTime
    TimerAvgB.Interval = getDataAvgDelayTime
    TimerDelayB.Interval = getDataDelayTime
End Sub

Private Sub CmmdLend_Click()

End Sub

Private Sub CmmdResetA_Click()
    Call CmmdDeletData_Click
End Sub

Private Sub CmmdResetSA_Click()
    Call CmmdDeletDataS_Click
End Sub

Private Sub CmmdResetZero_Click()
    mmPosiX = 0
    pPosiX = 0
    mmPosiY = 0
    pPosiY = 0
    mmPosiZ = 0
    pPosiZ = 0
    mmNPosiX = 0
    pNPosiX = 0
    mmNPosiY = 0
    pNPosiY = 0
    reset_pos 1
    reset_pos 2
    TxtLenX.Text = "0"
    TxtLenY.Text = "0"
    TxtLenXM.Text = "0"
    TxtLenYM.Text = "0"
    TxtPosiSetX.Text = "0"
    TxtPosiSetY.Text = "0"
End Sub

Private Sub CmmdResetZeroM_Click()
    Call CmmdResetZero_Click
End Sub

Private Sub CmmdScanA_Click()

    
    If connectPNAMark = False Then
        MsgBox "请先连接网分！"
        Exit Sub
    End If
    markStopScanA = False
    PicErrPrinter(0).Cls
    endXA = All2p(CmblUnitXMA.ListIndex, 0, Val(TxtendXA.Text))
    startXA = All2p(CmblUnitXMA.ListIndex, 0, Val(TxtStartXA.Text))
    stepXA = All2p(CmblUnitXMA.ListIndex, 0, Val(TxtStepXA.Text))
    endYA = All2p(CmblUnitXMA.ListIndex, 1, Val(TxtendYA.Text))
    startYA = All2p(CmblUnitXMA.ListIndex, 1, Val(TxtStartYA.Text))
    stepYA = All2p(CmblUnitXMA.ListIndex, 1, Val(TxtStepYA.Text))
    If startXA = endXA Then
        If startYA = endYA Then
            stepsDA = 0
        Else
            stepsDA = (endYA - startYA) / stepYA
        End If
        
    Else
        stepsDA = (endXA - startXA) / stepXA
    End If
    stepsIA = Int(stepsDA)
    IA = 0
    addXA = (startXA + IA * stepXA) - pPosiX
    addYA = (startYA + IA * stepYA) - pPosiY
    TxtLenX.Text = addXA
    TxtLenY.Text = addYA
    CmblUnitX.ListIndex = 2
    CmblUnitY.ListIndex = 2
    Call CmmdStart_Click
    
    markGetData = False
    
    TimerA.Enabled = True
    CmmdScanA.Enabled = False
End Sub

Private Sub CmmdScanSA_Click()

    
    If connectPNAMark = False Then
        MsgBox "请先连接网分！"
        Exit Sub
    End If
    markStopScanSA = False
    PicErrPrinter(0).Cls
    endXSA = All2p(CmblUnitXMA.ListIndex, 0, Val(TxtendXSA.Text))
    startXSA = All2p(CmblUnitXMA.ListIndex, 0, Val(TxtStartXSA.Text))
    stepXSA = All2p(CmblUnitXMA.ListIndex, 0, Val(TxtStepXSA.Text))
    endYSA = All2p(CmblUnitXMA.ListIndex, 1, Val(TxtendYSA.Text))
    startYSA = All2p(CmblUnitXMA.ListIndex, 1, Val(TxtStartYSA.Text))
    stepYSA = All2p(CmblUnitXMA.ListIndex, 1, Val(TxtStepYSA.Text))
    If startXSA = endXSA Then
        If startYSA = endYSA Then
            stepsDSA = 0
        Else
            stepsDSA = (endYSA - startYSA) / stepYSA
        End If
        
    Else
        stepsDSA = (endXSA - startXSA) / stepXSA
    End If
    stepsISA = Int(stepsDSA)
    ISA = 0
    addXSA = (startXSA + ISA * stepXSA) - pPosiX
    addYSA = (startYSA + ISA * stepYSA) - pPosiY
    TxtLenXS.Text = addXSA
    TxtLenYS.Text = addYSA
    CmblUnitX.ListIndex = 2
    CmblUnitY.ListIndex = 2
    Open TxtFileLoad.Text For Append As #1
    Print #1, RichDataSA.Text
    Close #1
    Call TxtLenXS_LostFocus
    Call TxtLenYS_LostFocus
    Call CmmdStartS_Click
    
    markGetDataS = False
    
    TimerB.Enabled = True
    CmmdScanSA.Enabled = False
End Sub

Private Sub CmmdScanSB_Click()
    If connectPNAMark = False Then
        MsgBox "请先连接网分！"
        Exit Sub
    End If
    markStopScanSB = False        '中止扫描灰掉，不能按
    PicErrPrinter(0).Cls          '报错窗口清零
    IASB = Val(TxtIASB.Text) - 1
    IBSB = Val(TxtIBSB.Text)
    startXSB = Val(TxtStartXSB.Text) 'Val()传值函数
    stepAXSB = Val(TxtStepAXSB.Text)
    stepBXSB = Val(TxtStepBXSB.Text)
    startYSB = Val(TxtStartYSB.Text)
    stepAYSB = Val(TxtStepAYSB.Text)
    stepBYSB = Val(TxtStepBYSB.Text)
    TxtStartXSA.Text = startXSB     'txtStartXSA自动扫谱的起点，A为auto，自动
    TxtStartYSA.Text = startYSB
    TxtendXSA.Text = startXSB + IASB * stepAXSB
    TxtendYSA.Text = startXSB + IASB * stepAYSB
    TxtStepXSA.Text = stepAXSB
    TxtStepYSA.Text = stepAYSB
    IB = 0
    Call CmmdScanSA_Click
    
    TimerSB.Enabled = True
    CmmdScanSB.Enabled = False
    
    
'    If startAXSB = endAXSB Then
'        If startYSA = endYSA Then
'            stepsDSA = 0
'        Else
'            stepsDSA = (endYSA - startYSA) / stepYSA
'        End If
'
'    Else
'        stepsDSA = (endXSA - startXSA) / stepXSA
'    End If
    
'    stepsISA = Int(stepsDSA)
'    ISA = 0
'    addXSA = (startXSA + ISA * stepXSA) - pPosiX
'    addYSA = (startYSA + ISA * stepYSA) - pPosiY
'    TxtLenXS.Text = addXSA
'    TxtLenYS.Text = addYSA
'    CmblUnitX.ListIndex = 2
'    CmblUnitY.ListIndex = 2
'    Open TxtFileLoad.Text For Append As #1
'    Print #1, RichDataSA.Text
'    Close #1
''    Call TxtLenXS_LostFocus
 '   Call TxtLenYS_LostFocus
 '   Call CmmdStartS_Click
 '
 '   markGetDataS = False
 '
 '   TimerB.Enabled = True
 '   CmmdScanSA.Enabled = False
End Sub

Private Sub CmmdSetPosi_Click()
    PicErrPrinter(0).Cls
    If Not IsNumeric(TxtPosiSetX.Text) Then
        PicErrPrinter(0).Print "X轴错误！位置必须是数字！"
        TxtPosiSetX.SelStart = 0
        TxtPosiSetX.SelLength = Len(TxtPosiSetX.Text)
        TxtPosiSetX.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtSpeedNumY.Text) Then
        PicErrPrinter(0).Print "Y轴错误！速度必须是数字！"
        TxtPosiSetY.SelStart = 0
        TxtPosiSetY.SelLength = Len(TxtPosiSetY.Text)
        TxtPosiSetY.SetFocus
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If Not IsNumeric(TxtSpeedNumZ.Text) Then
        PicErrPrinter(0).Print "Z轴错误！速度必须是数字！"
        TxtPosiSetZ.SelStart = 0
        TxtPosiSetZ.SelLength = Len(TxtPosiSetZ.Text)
        TxtPosiSetZ.SetFocus
        Exit Sub
    End If
    mmNPosiX = Val(TxtPosiSetX.Text) - mmPosiX
    mmNPosiY = Val(TxtPosiSetY.Text) - mmPosiY
    mmNPosiZ = Val(TxtPosiSetZ.Text) - mmPosiZ
End Sub

Private Sub CmmdSlowM_Click()

End Sub

Private Sub CmmdSetSpecFresh_Click()
    If connectPNAMark = True Then
        startFreq = ch.StartFrequency
        stopFreq = ch.StopFrequency
        numOfPoints = ch.NumberOfPoints
        TxtStartFreq.Text = startFreq
        TxtStopFreq.Text = stopFreq
        TxtNumOfPoints.Text = numOfPoints
        TxtStartFreqA.Text = startFreq
        TxtStopFreqA.Text = stopFreq
        TxtNumOfPointsA.Text = numOfPoints
        
        RichDataS.Text = "x/mm" & Chr(9) & "y/mm" & Chr(9) & "Frequency/Hz" & Chr(9) & "linmag" & Chr(9) & "logmag" & Chr(9) & "phase" & Chr(9) & "real" & Chr(9) & "img"
        sweepType = ch.sweepType
        If sweepType = 0 Then
            If numOfPoints > 1 Then
                For tmpi = 0 To numOfPoints - 1
                    freqx = startFreq + ((stopFreq - startFreq) * tmpi / (numOfPoints - 1))
                    'RichDataS.Text = RichDataS.Text & Chr(9) & Int(freqx)
                    freqss(tmpi) = Int(freqx)
                Next tmpi
            End If
            If numOfPoints = 1 Then
                'RichDataS.Text = RichDataS.Text & Chr(9) & Int(startFreq)
                freqss(0) = Int(startFreq)
            End If
        End If
        If sweepType = 1 Then
            If numOfPoints > 1 Then
                For tmpi = 0 To numOfPoints - 1
                    freqx = Exp(Log(startFreq) + (Log(stopFreq) - Log(startFreq)) * tmpi / (numOfPoints - 1))
                    'RichDataS.Text = RichDataS.Text & Chr(9) & Int(freqx)
                    freqss(tmpi) = Int(freqx)
                Next tmpi
            End If
            If numOfPoints = 1 Then
                'RichDataS.Text = RichDataS.Text & Chr(9) & Int(startFreq)
                freqss(0) = Int(startFreq)
            End If
        End If
        RichDataSA.Text = RichDataS.Text
    Else
        TxtStartFreq.Text = "0"
        TxtStopFreq.Text = "0"
        TxtNumOfPoints.Text = "0"
        TxtStartFreqA.Text = "0"
        TxtStopFreqA.Text = "0"
        TxtNumOfPointsA.Text = "0"
    End If
End Sub

Private Sub CmmdSetSpecFreshA_Click()
    Call CmmdSetSpecFresh_Click
End Sub

Private Sub CmmdSetSpecSure_Click()
    'startFreq = Val(TxtStartFreq.Text)
    'stopFreq = Val(TxtStopFreq.Text)
    rateFreq = Val(TxtRateFreq.Text)
    'ch.StartFrequency = startFreq
    'ch.StopFrequency = stopFreq
    'ch.NumberOfPoints = numOfPoints
    Call CmmdSetSpecFresh_Click
End Sub

Private Sub CmmdSetSpecSureA_Click()
    'startFreq = Val(TxtStartFreqA.Text)
    'stopFreq = Val(TxtStopFreqA.Text)
    'numOfPoints = Val(TxtNumOfPointsA.Text)
    rateFreq = Val(TxtRateFreqA.Text)
    'ch.StartFrequency = startFreq
    'ch.StopFrequency = stopFreq
    'ch.NumberOfPoints = numOfPoints
    Call CmmdSetSpecFresh_Click
End Sub

Private Sub CmmdSpeedCancel_Click()
    TxtSpeedNumX.Text = mmConSpeedX
    TxtSpeedNumY.Text = mmConSpeedY
    TxtSpeedNumZ.Text = mmConSpeedZ
End Sub

Private Sub CmmdSpeedSave_Click()
    Call CmmdSpeedSure_Click
    If speedOK = 1 Then
            Open App.Path + "\speedsets\default.txt" For Output As #1
                Write #1, pMaxSpeedX, pMaxSpeedY, pMaxSpeedZ, pConSpeedX, pConSpeedY, pConSpeedZ
            Close #1
    End If
End Sub

Private Sub CmmdSpeedSaveAs_Click()

End Sub

Private Sub CmmdSpeedSure_Click()
    Dim tmmConSpeedX As Double
    Dim tmmConSpeedY As Double
    Dim tmmConSpeedZ As Double
    Dim tpConSpeedX As Double
    Dim tpConSpeedY As Double
    Dim tpConSpeedZ As Double
    speedOK = 0
    PicErrPrinter(0).Cls
    If Not IsNumeric(TxtSpeedNumX.Text) Then
        PicErrPrinter(0).Print "X轴错误！速度必须是数字！"
        TxtSpeedNumX.SelStart = 0
        TxtSpeedNumX.SelLength = Len(TxtSpeedNumX.Text)
        TxtSpeedNumX.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtSpeedNumY.Text) Then
        PicErrPrinter(0).Print "Y轴错误！速度必须是数字！"
        TxtSpeedNumY.SelStart = 0
        TxtSpeedNumY.SelLength = Len(TxtSpeedNumY.Text)
        TxtSpeedNumY.SetFocus
        Exit Sub
    End If
     If Not IsNumeric(TxtSpeedNumZ.Text) Then
        PicErrPrinter(0).Print "Z轴错误！速度必须是数字！"
        TxtSpeedNumZ.SelStart = 0
        TxtSpeedNumZ.SelLength = Len(TxtSpeedNumZ.Text)
        TxtSpeedNumZ.SetFocus
        Exit Sub
    End If
    If CmbSpeedUnitX.ListIndex = 0 Then
        tmmConSpeedX = Val(TxtSpeedNumX.Text)
        tpConSpeedX = Mm2pX(tmmConSpeedX)
    ElseIf CmbSpeedUnitX.ListIndex = 1 Then
        tmmConSpeedX = Val(TxtSpeedNumX.Text) * 10
        tpConSpeedX = Mm2pX(tmmConSpeedX)
    Else
        tpConSpeedX = Val(TxtSpeedNumX.Text)
        tmmConSpeedX = P2mmX(tpConSpeedX)
    End If
    If CmbSpeedUnitY.ListIndex = 0 Then
        tmmConSpeedY = Val(TxtSpeedNumY.Text)
        tpConSpeedY = Mm2pY(tmmConSpeedY)
    ElseIf CmbSpeedUnitY.ListIndex = 1 Then
        tmmConSpeedY = Val(TxtSpeedNumY.Text) * 10
        tpConSpeedY = Mm2pY(tmmConSpeedY)
    Else
        tpConSpeedY = Val(TxtSpeedNumY.Text)
        tmmConSpeedY = P2mmY(tpConSpeedY)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If CmbSpeedUnitZ.ListIndex = 0 Then
        tmmConSpeedZ = Val(TxtSpeedNumZ.Text)
        tpConSpeedZ = Mm2pZ(tmmConSpeedZ)
    ElseIf CmbSpeedUnitZ.ListIndex = 1 Then
        tmmConSpeedZ = Val(TxtSpeedNumZ.Text) * 10
        tpConSpeedZ = Mm2pZ(tmmConSpeedZ)
    Else
        tpConSpeedZ = Val(TxtSpeedNumZ.Text)
        tmmConSpeedZ = P2mmY(tpConSpeedZ)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If tpConSpeedX > 25600 Then
        PicErrPrinter(0).Print "X轴速度过大，建议低于5mm/s"
        TxtSpeedNumX.SelStart = 0
        TxtSpeedNumX.SelLength = Len(TxtSpeedNumX.Text)
        TxtSpeedNumX.SetFocus
        Exit Sub
    End If
    If tpConSpeedY > 25600 Then
        PicErrPrinter(0).Print "Y轴速度过大，建议低于5mm/s"
        TxtSpeedNumY.SelStart = 0
        TxtSpeedNumY.SelLength = Len(TxtSpeedNumY.Text)
        TxtSpeedNumY.SetFocus
        Exit Sub
    End If
     If tpConSpeedZ > 25600 Then
        PicErrPrinter(0).Print "Z轴速度过大，建议低于5mm/s"
        TxtSpeedNumZ.SelStart = 0
        TxtSpeedNumZ.SelLength = Len(TxtSpeedNumZ.Text)
        TxtSpeedNumZ.SetFocus
        Exit Sub
    End If
        speedOK = 1
        pConSpeedX = tpConSpeedX
        pConSpeedY = tpConSpeedY
        pConSpeedZ = tpConSpeedZ
        mmConSpeedX = tmmConSpeedX
        mmConSpeedY = tmmConSpeedY
        mmConSpeedZ = tmmConSpeedZ
        set_conspeed 1, pConSpeedX
        set_conspeed 2, pConSpeedY
        set_conspeed 3, pConSpeedZ
End Sub

Private Sub CmmdStart_Click()
    PicErrPrinter(0).Cls
    pAddX = 0     'pulse add， x方向
    pAddY = 0
    pAddZ = 0
    If CmblUnitX.ListIndex = 2 Then
        If Not IsInt(TxtLenX.Text) Then
            PicErrPrinter(0).Print "X轴错误！Pulse位移必须为整数！"
            TxtLenX.SelStart = 0
            TxtLenX.SelLength = Len(TxtLenX.Text)
            Call TxtLenX.SetFocus
            Exit Sub
        End If
    Else
        If Not IsNumeric(TxtLenX.Text) Then
            PicErrPrinter(0).Print "X轴错误！位移必须为数字！"
            TxtLenX.SelStart = 0
            TxtLenX.SelLength = Len(TxtLenX.Text)
            TxtLenX.SetFocus
            Exit Sub
        End If
    End If
    
    If CmblUnitY.ListIndex = 2 Then
        If Not IsInt(TxtLenY.Text) Then
            PicErrPrinter(0).Print "Y轴错误！Pulse位移必须为整数！"
            TxtLenY.SelStart = 0
            TxtLenY.SelLength = Len(TxtLenY.Text)
            TxtLenY.SetFocus
            Exit Sub
        End If
    Else
        If Not IsNumeric(TxtLenY.Text) Then
            PicErrPrinter(0).Print "Y轴错误！位移必须为数字！"
            TxtLenY.SelStart = 0
            TxtLenY.SelLength = Len(TxtLenY.Text)
            TxtLenY.SetFocus
            Exit Sub
        End If
    End If
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If CmblUnitZ.ListIndex = 2 Then
        If Not IsInt(TxtLenZ.Text) Then
            PicErrPrinter(0).Print "Z轴错误！Pulse位移必须为整数！"
            TxtLenZ.SelStart = 0
            TxtLenZ.SelLength = Len(TxtLenZ.Text)
            Call TxtLenZ.SetFocus
            Exit Sub
        End If
    Else
        If Not IsNumeric(TxtLenZ.Text) Then
            PicErrPrinter(0).Print "Z轴错误！位移必须为数字！"
            TxtLenZ.SelStart = 0
            TxtLenZ.SelLength = Len(TxtLenX.Text)
            TxtLenZ.SetFocus
            Exit Sub
        End If
    End If
    pAddX = All2p(CmblUnitX.ListIndex, 0, Val(TxtLenX.Text))  '所有的单位都必须转换成pulse
    MoveX 1       '判断是否应该加上螺纹误差
    CmmdStart.Enabled = False
    CmmdStartM.Enabled = False
    CmmdStartS.Enabled = False
    markBack = 1
    markMoveBack = 1
    pAddY = All2p(CmblUnitY.ListIndex, 1, Val(TxtLenY.Text))
    MoveY 1
    CmmdStart.Enabled = False
    CmmdStartM.Enabled = False
    CmmdStartS.Enabled = False
    markBack = 1
    markMoveBack = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    pAddZ = All2p(CmblUnitZ.ListIndex, 2, Val(TxtLenZ.Text))
    MoveZ 1
    CmmdStart.Enabled = False
    CmmdStartM.Enabled = False
    CmmdStartS.Enabled = False
    markBack = 1
    markMoveBack = 1
End Sub

Private Sub CmmdStopM_Click()

End Sub

Private Sub CmmdStartM_Click()
    Call CmmdStart_Click
End Sub



Private Sub CmmdStopScanA_Click()
    markStopScanA = True
End Sub



Private Sub CmmdConnectPNASA_Click()
    Call CmmdConnectPNAM_Click
End Sub

Private Sub CmmdStopScanSA_Click()
    markStopScanSA = True
End Sub

Private Sub CmmdStartS_Click()
    Call CmmdStart_Click
End Sub

Private Sub CmmdFendS_Click()
    Call CmmdFend_Click
End Sub

Private Sub CmmdStopScanSB_Click()
    markStopScanSB = True
    Call CmmdStopScanSA_Click
End Sub

Private Sub Combo2_LostFocus()
CmblUnitZMA.ListIndex = CmblUnitX.ListIndex
    CmblUnitZM.ListIndex = CmblUnitZ.ListIndex
    CmblUnitZSA.ListIndex = CmblUnitZ.ListIndex
    CmblUnitZSB.ListIndex = CmblUnitZ.ListIndex
    CmblUnitZS.ListIndex = CmblUnitZ.ListIndex
End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command11_Click()
    Call CmmdDeletDataS_Click
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()

End Sub

Private Sub Command2_Click()
    PicErrPrinter(0).Print sweepType
End Sub





Private Sub Command3_Click()
    Call CmmdConnectPNAM_Click
End Sub

Private Sub Command4_Click()
    Call CmmdResetZero_Click
End Sub


Private Sub CmmdGetDataS_Click()
    If connectPNAMark = False Then
        MsgBox "请先连接网分！"
        Exit Sub
    End If
   
    Open TxtFileLoad.Text For Append As #1
    
    tmpSpecReal = meas.GetData(naRawData, naDataFormat_Real)
    tmpSpecImg = meas.GetData(naRawData, naDataFormat_Imaginary)
    tmpSpecLin = meas.GetData(naRawData, naDataFormat_LinMag)
    tmpSpecLog = meas.GetData(naRawData, naDataFormat_LogMag)
    tmpSpecPhase = meas.GetData(naRawData, naDataFormat_Phase)
    
    tmpnumOfPoints = ch.NumberOfPoints
    
    For tmpi = 0 To tmpnumOfPoints - 1
        specReal(tmpi) = 0 'tmpSpecReal(i)
        specImg(tempi) = 0 'tmpSpecImg(i)
        specLog(tempi) = 0 'tmpSpecLog(i)
        specLin(tempi) = 0 'tmpSpecLin(i)
        specPhase(tempi) = 0 'tmpSpecPhase(i)
        
    Next tmpi
    'PicErrPrinter(0).Print specReal(0)
        
    outPosiXS = P2mmX(pPosiX) + mmNPosiX
    outPosiYS = P2mmY(pPosiY) + mmNPosiY
    IAvgS = 0
    TimerDelayB.Enabled = True
    CmmdGetDataS.Enabled = False
End Sub



Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()

End Sub

Private Sub fileinfos_DragDrop(Source As Control, X As Single, Y As Single, Z As Single)

End Sub

Private Sub Form_Load()
    CmblUnitX.ListIndex = 0
    CmblUnitY.ListIndex = 0
    CmblUnitXM.ListIndex = 0
    CmblUnitYM.ListIndex = 0
    CmblUnitXMA.ListIndex = 0
    CmblUnitYMA.ListIndex = 0
    CmblUnitXS.ListIndex = 0
    CmblUnitYS.ListIndex = 0
    CmblUnitXSA.ListIndex = 0
    CmblUnitYSA.ListIndex = 0
    CmblUnitXSB.ListIndex = 0
    CmblUnitYSB.ListIndex = 0
    CmbSpeedUnitX.ListIndex = 0
    CmbSpeedUnitY.ListIndex = 0
    RichDataM.Text = "x/mm" & Chr(9) & "y/mm" & Chr(9) & "linmag" & Chr(9) & "logmag" & Chr(9) & "phase"
    RichDataA.Text = "x/mm" & Chr(9) & "y/mm" & Chr(9) & "linmag" & Chr(9) & "logmag" & Chr(9) & "phase"
    RichDataS.Text = "x/mm" & Chr(9) & "y/mm"
    RichDataSA.Text = "x/mm" & Chr(9) & "y/mm"
    mmPosiX = 0
    pPosiX = 0
    mmPosiY = 0
    pPosiY = 0
    mmNPosiX = 0
    pNPosiX = 0
    mmNPosiY = 0
    pNPosiY = 0
    rateFreq = 1
    getDataDelayTime = 5000
    getDataAvgNum = 10
    getDataAvgDelayTime = 50
    reset_pos 1
    reset_pos 2
    TimerAvgA.Enabled = False
    TimerAvgA.Interval = getDataAvgDelayTime
    TimerDelayA.Enabled = False
    TimerDelayA.Interval = getDataDelayTime
    TimerAvgB.Enabled = False
    TimerAvgB.Interval = getDataAvgDelayTime
    TimerDelayB.Enabled = False
    TimerDelayB.Interval = getDataDelayTime
    TimerA.Enabled = False
    TimerB.Enabled = False
    TimerSB.Enabled = False
    TimerFend.Enabled = False
    CmmdLend.Enabled = False
    CmmdLendM.Enabled = False
    TxtSpeedNumX.Text = mmConSpeedX
    TxtSpeedNumY.Text = mmConSpeedY
    TxtGetDataDelayTime.Text = getDataDelayTime
    TxtGetDataAvgNum.Text = getDataAvgNum
    TxtGetDataAvgDelayTime.Text = getDataAvgDelayTime
    TxtFileLoad.Text = "D:\data.txt"
    
    If connectPNAMark = True Then
        startFreq = ch.StartFrequency
        stopFreq = ch.StopFrequency
        numOfPoints = ch.NumberOfPoints
        TxtStartFreq.Text = startFreq
        TxtStopFreq.Text = stopFreq
        TxtNumOfPoints.Text = numOfPoints
        TxtStartFreqA.Text = startFreq
        TxtStopFreqA.Text = stopFreq
        TxtNumOfPointsA.Text = numOfPoints
        
        RichDataS.Text = "x/mm" & Chr(9) & "y/mm" & Chr(9) & "Frequency/Hz" & Chr(9) & "linmag" & Chr(9) & "logmag" & Chr(9) & "phase" & Chr(9) & "real" & Chr(9) & "img"
        sweepType = ch.sweepType
        If sweepType = 0 Then
            If numOfPoints > 1 Then
                For tmpi = 0 To numOfPoints - 1
                    freqx = startFreq + ((stopFreq - startFreq) * tmpi / (numOfPoints - 1))
                    'RichDataS.Text = RichDataS.Text & Chr(9) & Int(freqx)
                    freqss(tmpi) = Int(freqx)
                Next tmpi
            End If
            If numOfPoints = 1 Then
                'RichDataS.Text = RichDataS.Text & Chr(9) & Int(startFreq)
                freqss(0) = Int(startFreq)
            End If
        End If
        If sweepType = 1 Then
            If numOfPoints > 1 Then
                For tmpi = 0 To numOfPoints - 1
                    freqx = Exp(Log(startFreq) + (Log(stopFreq) - Log(startFreq)) * tmpi / (numOfPoints - 1))
                    'RichDataS.Text = RichDataS.Text & Chr(9) & Int(freqx)
                    freqss(tmpi) = Int(freqx)
                Next tmpi
            End If
            If numOfPoints = 1 Then
                'RichDataS.Text = RichDataS.Text & Chr(9) & Int(startFreq)
                freqss(0) = Int(startFreq)
            End If
        End If
        RichDataSA.Text = RichDataS.Text
    Else
        TxtStartFreq.Text = "0"
        TxtStopFreq.Text = "0"
        TxtNumOfPoints.Text = "0"
        TxtStartFreqA.Text = "0"
        TxtStopFreqA.Text = "0"
        TxtNumOfPointsA.Text = "0"
    End If
    'TimerPoint.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("是否退出MPC08控制程序？", vbYesNo, "提示") = vbNo Then Cancel = True
End Sub









Private Sub test_Click()
    PicErrPrinter(0).Print ch.sweepType
End Sub


Private Sub Frame12_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame14_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame15_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame8_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame9_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub FramePosiSet_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub FrameSetPoint_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label30_Click()

End Sub

Private Sub Label36_Click()

End Sub

Private Sub Label37_Click()

End Sub

Private Sub LblPosiNumX_Click()

End Sub

Private Sub LblPosiNumY_Click()

End Sub

Private Sub LblSpeedNumY_Click()

End Sub

Private Sub LblXRun_Click()

End Sub

Private Sub LblYRun_Click()

End Sub

Private Sub PicErrPrinter_Click(Index As Integer)

End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub TimerA_Timer()
    If CmmdStart.Enabled = True Then
        'PicErrPrinter(2).Print IA & stepsI
        If markGetData = False Then
            Call CmmdGetData_Click
            markGetData = True
        End If
        If CmmdGetData.Enabled = False Then
            Exit Sub
        End If
        IA = IA + 1
        If IA > stepsIA Then
            markStopScanA = True
        End If
        If markStopScanA = True Then
            TimerA.Enabled = False
            CmmdScanA.Enabled = True
            Call CmblUnitXMA_LostFocus
            Call CmblUnitYMA_LostFocus
            Exit Sub
        End If
        addXA = (startXA + IA * stepXA) - pPosiX
        addYA = (startYA + IA * stepYA) - pPosiY
        TxtLenX.Text = addXA
        TxtLenY.Text = addYA
        CmblUnitX.ListIndex = 2
        CmblUnitY.ListIndex = 2
        Call CmmdStart_Click
        markGetData = False
    End If
    
End Sub

Private Sub TimerAvgA_Timer()
    IAvg = IAvg + 1
    If IAvg = getDataAvgNum Then
        TimerAvgA.Enabled = False
        CmmdGetData.Enabled = True
        RichDataM.Text = RichDataM.Text & vbLf & outPosiX & Chr(9) & outPosiY & Chr(9) & Round(linMagMarker / getDataAvgNum, 3) & Chr(9) & Round(logMagMarker / getDataAvgNum, 3) & Chr(9) & Round(phaseMarker / getDataAvgNum, 3)
        RichDataA.Text = RichDataM.Text
    End If
    
    linMagMarker = linMagMarker + meas.marker(1).Value(naMarkerFormat_LinMag)

    logMagMarker = logMagMarker + meas.marker(1).Value(naMarkerFormat_LogMag)

    phaseMarker = phaseMarker + meas.marker(1).Value(naMarkerFormat_Phase)

End Sub

Private Sub TimerAvgB_Timer()

    IAvgS = IAvgS + 1
    If IAvgS > getDataAvgNum Then
        TimerAvgB.Enabled = False
        'RichDataS.Text = RichDataS.Text & vbLf & outPosiXS & Chr(9) & outPosiYS
        
    For tmpi = 0 To tmpnumOfPoints - 1
        specReal(tmpi) = specReal(tmpi) / getDataAvgNum
        specImg(tmpi) = specImg(tmpi) / getDataAvgNum
        specLog(tmpi) = specLog(tmpi) / getDataAvgNum
        specLin(tmpi) = specLin(tmpi) / getDataAvgNum
        specPhase(tmpi) = specPhase(tmpi) / getDataAvgNum
        'RichDataS.SelStart = Len(RichDataS.Text)
        'RichDataS.SelText = vbLf & outPosiXS & Chr(9) & outPosiYS & Chr(9) & Round(freqss(tmpi), 6) & Chr(9) & Round(specLin(tmpi), 6) & Chr(9) & Round(specLog(tmpi), 6) & Chr(9) & Round(specPhase(tmpi), 6) & Chr(9) & Round(specReal(tmpi), 6) & Chr(9) & Round(specImg(tmpi), 6)
        'RichDataS.SelStart = Len(RichDataS.Text)
        Print #1, vbLf & outPosiXS & Chr(9) & outPosiYS & Chr(9) & Round(freqss(tmpi), 6) & Chr(9) & Round(specLin(tmpi), 6) & Chr(9) & Round(specLog(tmpi), 6) & Chr(9) & Round(specPhase(tmpi), 6) & Chr(9) & Round(specReal(tmpi), 6) & Chr(9) & Round(specImg(tmpi), 6)
        'RichDataSA.SelStart = Len(RichDataSA.Text)
        'RichDataSA.SelText = vbLf & outPosiXS & Chr(9) & outPosiYS & Chr(9) & Round(freqss(tmpi), 6) & Chr(9) & Round(specLin(tmpi), 6) & Chr(9) & Round(specLog(tmpi), 6) & Chr(9) & Round(specPhase(tmpi), 6) & Chr(9) & Round(specReal(tmpi), 6) & Chr(9) & Round(specImg(tmpi), 6)
        'RichDataSA.SelStart = Len(RichDataSA.Text)
    Next tmpi
        'RichDataSA.Text = RichDataS.Text
        Close #1
        CmmdGetDataS.Enabled = True
    End If
    
    tmpSpecReal = meas.GetData(naRawData, naDataFormat_Real)
    tmpSpecImg = meas.GetData(naRawData, naDataFormat_Imaginary)
    tmpSpecLin = meas.GetData(naRawData, naDataFormat_LinMag)
    tmpSpecLog = meas.GetData(naRawData, naDataFormat_LogMag)
    tmpSpecPhase = meas.GetData(naRawData, naDataFormat_Phase)
    tmpnumOfPoints = ch.NumberOfPoints
    
    For tmpi = 0 To tmpnumOfPoints - 1
        specReal(tmpi) = specReal(tmpi) + tmpSpecReal(tmpi)
        specImg(tmpi) = specImg(tmpi) + tmpSpecImg(tmpi)
        specLog(tmpi) = specLog(tmpi) + tmpSpecLog(tmpi)
        specLin(tmpi) = specLin(tmpi) + tmpSpecLin(tmpi)
        specPhase(tmpi) = specPhase(tmpi) + tmpSpecPhase(tmpi)
    Next tmpi
    
End Sub

Private Sub TimerB_Timer()
    If CmmdStart.Enabled = True Then
        If markGetDataS = False Then
            Call CmmdGetDataS_Click
            markGetDataS = True
        End If
        If CmmdGetDataS.Enabled = False Then
            Exit Sub
        End If
        ISA = ISA + 1
        If ISA > stepsISA Then
            markStopScanSA = True
        End If
        If markStopScanSA = True Then
            TimerB.Enabled = False
            CmmdScanSA.Enabled = True
            'Close #1
            Exit Sub
        End If
        addXSA = (startXSA + ISA * stepXSA) - pPosiX
        addYSA = (startYSA + ISA * stepYSA) - pPosiY
        TxtLenXS.Text = addXSA
        TxtLenYS.Text = addYSA
        CmblUnitX.ListIndex = 2
        CmblUnitY.ListIndex = 2
        Call TxtLenXS_LostFocus
        Call TxtLenYS_LostFocus
        Call CmmdStartS_Click
        markGetDataS = False
    End If
End Sub

Private Sub TimerDelayA_Timer()
    TimerAvgA.Enabled = True
    TimerDelayA.Enabled = False
End Sub

Private Sub TimerDelayB_Timer()
    TimerAvgB.Enabled = True
    TimerDelayB.Enabled = False
End Sub

Private Sub TimerFend_Timer()
    mmPosiX = P2mmX(pPosiX)          'P2mm（pulse转换成mm）
    mmPosiY = P2mmY(pPosiY)
    mmPosiZ = P2mmZ(pPosiZ)
    TimerFend.Enabled = False
End Sub

Private Sub TimerPoint_Timer()
    get_abs_pos 1, pPosiX
    get_abs_pos 2, pPosiY
    get_abs_pos 3, pPosiZ
    LblPosiNumX.Caption = P2mmX(pPosiX) + mmNPosiX
    LblPosiNumY.Caption = P2mmY(pPosiY) + mmNPosiY
    LblPosiNumZ.Caption = P2mmZ(pPosiZ) + mmNPosiZ
    If (check_done(1)) + (check_done(2)) + (check_done(3)) = 0 Then 'check_done步径内置函数，查看一个轴的运动是否结束
        If markMoveBack = 0 Then
            markBack = 0
        End If
        
        If markMoveBack <> 0 Then
            Call moveBack
        End If
        If markBack = 0 Then
            CmmdStart.Enabled = True
            CmmdStartM.Enabled = True
        End If
    End If
    LblSpeedNumX.Caption = P2mmX(get_rate(1))
    LblSpeedNumY.Caption = P2mmY(get_rate(2))
    LblSpeedNumZ.Caption = P2mmZ(get_rate(3))
End Sub


Private Sub TimerSB_Timer()
    If CmmdScanSA.Enabled = True Then
        
        If IB >= IBSB - 1 Then
            markStopScanSB = True
        End If
        If markStopScanSB = True Then
            TimerSB.Enabled = False
            CmmdScanSB.Enabled = True
            Exit Sub
            
        End If
        
        IB = IB + 1
        TxtStartXSA.Text = startXSB + stepBXSB * IB
        TxtStartYSA.Text = startYSB + stepBYSB * IB
        TxtendXSA.Text = startXSB + IASB * stepAXSB + stepBXSB * IB
        TxtendYSA.Text = startXSB + IASB * stepAYSB + stepBYSB * IB
        TxtStepXSA.Text = stepAXSB
        TxtStepYSA.Text = stepAYSB
        Call CmmdScanSA_Click
    
    End If
    
    
End Sub

Private Sub TxtFileLoad_Change()

End Sub

Private Sub TxtIASB_Change()

End Sub

Private Sub TxtIBSB_Change()

End Sub

Private Sub TxtLenX_GotFocus()
    TxtLenX.SelStart = 0
    TxtLenX.SelLength = Len(TxtLenX.Text)
End Sub

Private Sub TxtLenXM_LostFocus()
    TxtLenX.Text = TxtLenXM.Text
    TxtLenXS.Text = TxtLenXM.Text
End Sub
Private Sub TxtLenYM_LostFocus()
    TxtLenY.Text = TxtLenYM.Text
    TxtLenYS.Text = TxtLenYM.Text
End Sub
Private Sub TxtLenX_LostFocus()
    TxtLenXM.Text = TxtLenX.Text
    TxtLenXS.Text = TxtLenX.Text
End Sub
Private Sub TxtLenY_LostFocus()
    TxtLenYM.Text = TxtLenY.Text
    TxtLenYS.Text = TxtLenY.Text
End Sub
Private Sub TxtLenXS_LostFocus()
    TxtLenX.Text = TxtLenXS.Text
    TxtLenXM.Text = TxtLenXS.Text
End Sub
Private Sub TxtLenYS_LostFocus()
    TxtLenY.Text = TxtLenYS.Text
    TxtLenYM.Text = TxtLenYS.Text
End Sub

Private Sub TxtLeny_GotFocus()
    TxtLenY.SelStart = 0
    TxtLenY.SelLength = Len(TxtLenY.Text)
End Sub

Private Sub TxtLenZ_GotFocus()
    TxtLenZ.SelStart = 0
    TxtLenZ.SelLength = Len(TxtLenZ.Text)
End Sub

Private Sub TxtSpeedNumX_GotFocus()
    TxtSpeedNumX.SelStart = 0
    TxtSpeedNumX.SelLength = Len(TxtSpeedNumX.Text)
End Sub
Private Sub TxtSpeedNumY_GotFocus()
    TxtSpeedNumY.SelStart = 0
    TxtSpeedNumY.SelLength = Len(TxtSpeedNumY.Text)
End Sub
Private Sub TxtPosiSetX_GotFocus()
    TxtPosiSetX.SelStart = 0
    TxtPosiSetX.SelLength = Len(TxtPosiSetX.Text)
End Sub
Private Sub TxtPosiSetY_GotFocus()
    TxtPosiSetY.SelStart = 0
    TxtPosiSetY.SelLength = Len(TxtPosiSetY.Text)
End Sub

Private Sub TxtStartFreq_Change()

End Sub

Private Sub TxtStartFreqA_Change()

End Sub

Private Sub TxtStartXSB_Change()

End Sub

Private Sub Z方向速度_Click()

End Sub
