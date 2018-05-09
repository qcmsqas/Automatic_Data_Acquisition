VERSION 5.00
Begin VB.Form frmIni 
   Caption         =   "Initializing"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   8640
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton connect 
      Caption         =   "连接网分"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Timer TimerIni 
      Interval        =   200
      Left            =   8160
      Top             =   2880
   End
   Begin VB.CommandButton ReTryini 
      Caption         =   "重试"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Cancelini 
      Caption         =   "取消"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Sureini 
      Caption         =   "确定"
      Height          =   375
      Left            =   4560
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
Dim numAxes() As Long '第一个卡的轴数

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
'    frmIni.Print Tab(5); numTotalAxes
    numCard = init_board()
    If numCard < 0 Then
       IniBoard = -2
'    frmIni.Print Tab(5); IniBoard
       Exit Function
    End If
    If numCard = 0 Then
        IniBoard = -1
'    frmIni.Print Tab(5); IniBoard
        Exit Function
    End If
'    frmIni.Print Tab(5); numCard
    IniBoard = 1
    ReDim numAxes(numCard)
    For i = 1 To numCard
        numAxes(i) = get_axe(i)
    Next i
End Function

Private Sub Cancelini_Click()
    markCancel = 1
    Unload Me
End Sub

Private Sub connect_Click()
    ' frmIni.Print Tab(5); "正在初始化网分"
    On Error GoTo err1

    StartNA
    connectPNAMark = True
    frmIni.Print Tab(5); "连接网分成功！"
    Exit Sub
err1:
    MsgBox "请先开启网分！"
End Sub

Private Sub Form_Activate()
    ReTryini.Enabled = False
    connectPNAMark = False
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
    Print Tab(5); "检测到卡数目"; numCard; "总轴数目"; numTotalAxes
    Print Tab(5); "正在读取默认校正文件"
    markZ = ReadCorrect
    If markZ = -1 Then
        Print Tab(5); "默认校正文件不存在，使用默认校正参数1mm = 3200pulse"
    ElseIf markZ = 0 Then
        Print Tab(5); "读取默认校正文件错误，使用默认校正参数1mm = 3200pulse"
        Call DefaultCorrect
    Else
        Print Tab(5); "读取默认校正文件成功，X轴 "; mmMaxX; "mm = "; pMaxX; "pulse，Y轴 "; mmMaxY; "mm = "; pMaxY; "pulse。"
    End If
    Print Tab(5); "正在读取默认速度设置文件"
    markA = ReadSpeedSet
    If markA = -1 Then
        Print Tab(5); "默认速度设置文件不存在，使用默认参数"
        pConSpeedX = 2000
        pConSpeedY = 2000
        pMaxSpeedX = 80000
        pMaxSpeedY = 80000
    ElseIf markA = 0 Then
        Print Tab(5); "读取默认速度设置文件错误，使用默认参数"
        pConSpeedX = 2000
        pConSpeedY = 2000
        pMaxSpeedX = 80000
        pMaxSpeedY = 80000
    Else
        Print Tab(5); "读取默认速度设置文件成功"
        If SetBoard = 0 Then
            Print Tab(5); "设置速度参数成功"
        Else
            Print Tab(5); "设置速度参数失败，使用默认参数"
            pConSpeedX = 2000
            pConSpeedY = 2000
            pMaxSpeedX = 80000
            pMaxSpeedY = 80000
        End If
    End If
    Print Tab(5); "正在消除螺距误差"
    con_pmove 1, -errA
    con_pmove 2, -errA
    markTimer = 0
    TimerIni.Enabled = True
    mmConSpeedX = pConSpeedX * mmMaxX / pMaxX
    mmMaxSpeedX = pMaxSpeedX * mmMaxX / pMaxX
    mmConSpeedY = pConSpeedY * mmMaxY / pMaxY
    mmMaxSpeedY = pMaxSpeedY * mmMaxY / pMaxY

End Sub

Private Sub Form_Load()
    ReTryini.Enabled = False
    Sureini.Enabled = False
    frmIni.AutoRedraw = True
    markCancel = 0
    markX = -1
    errA = 1600
    TimerIni.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If markCancel = 1 Then
        If MsgBox("是否退出MPC08控制程序？", vbYesNo, "提示") = vbNo Then Cancel = True
        markCancel = 0
        Exit Sub
    End If
    If markX <= 0 Then
        If MsgBox("是否退出MPC08控制程序？", vbYesNo, "提示") = vbNo Then Cancel = True
    End If
End Sub

Private Sub ReTryini_Click()
    Call Form_Activate
End Sub

Private Sub Sureini_Click()
    If markX <= 0 Then
        Unload Me
    End If
    
    If markX > 0 Then
        frmMain.Show
        Unload Me
    End If
End Sub

Private Sub TimerIni_Timer()
    If (check_done(1) + check_done(2)) = 0 Then
        If markTimer = 0 Then
            con_pmove 1, errA
            con_pmove 2, errA
            markTimer = 1
        Else
            Print Tab(5); "初始化完成，请点击 确定 继续"
            Sureini.Enabled = True
            TimerIni.Enabled = False
        End If
    End If
    
End Sub
