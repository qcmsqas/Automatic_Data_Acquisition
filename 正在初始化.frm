VERSION 5.00
Begin VB.Form frmIni 
   Caption         =   "Initializing"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ReTryini 
      Caption         =   "重试"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Cancelini 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Sureini 
      Caption         =   "确定"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public numTotalAxes As Long '总轴数
Public numCard As Long '总卡数
Dim markX As Long

Private Sub Label1_Click()

End Sub
Private Function IniBoard() As Long '初始化步进控制卡
    numTotalAxes = auto_set()
    If numTotalAxes < 0 Then
        IniBoard = -2
        Exit Function
    End If
    If numTotalAxes = 0 Then
        IniBoard = 0
        Exit Function
    End If
    numCard = init_board()
    If numCard < 0 Then
       IniBoard = -2
       Exit Function
    End If
    If numCard = 0 Then
        IniBoard = -1
        Exit Function
    End If
    IniBoard = 1
End Function

Private Sub Cancelini_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ReTryini.Enabled = False
    Sureini.Enabled = False
    Cls
    Print Tab(5)
    Print Tab(5)
    Print Tab(5); "正在初始化……"
    Print Tab(5); "正在开启步进控制卡"
    markX = IniBoard
    If markX = 0 Then
        Print Tab(5); "检测不到控制卡，请检查MPC08运动控制卡是否正确插入您的计算机。"
        ReTryini.Enabled = True
        Sureini.Enabled = True
        Exit Sub
    End If
    If markX < 0 Then
        Print Tab(5); "控制程序调用错误，请检查MPC08相关软件是否正确安装。"
        ReTryini.Enabled = True
        Sureini.Enabled = True
        Exit Sub
    End If
    Sureini.Enabled = True
    
End Sub

Private Sub Form_Load()
    ReTryini.Enabled = False
    Sureini.Enabled = False
    frmIni.AutoRedraw = True
End Sub

Private Sub Form_Unload(cancel As Integer)
    If markX <= 0 Then
        If MsgBox("是否退出MPC08控制程序？", vbYesNo, "提示") = vbNo Then cancel = True
    End If
End Sub

Private Sub ReTryini_Click()
    Call Form_Activate
End Sub

Private Sub Sureini_Click()
    If markX <= 0 Then
        Unload Me
    End If
    frmMain.Show
    Unload Me
End Sub
