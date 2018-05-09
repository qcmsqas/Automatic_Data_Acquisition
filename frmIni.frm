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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton connect 
      Caption         =   "��������"
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
      Caption         =   "����"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Cancelini 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Sureini 
      Caption         =   "ȷ��"
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
Dim numAxes() As Long '��һ����������

Private Function IniBoard() As Long '��ʼ���������ƿ�
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
    ' frmIni.Print Tab(5); "���ڳ�ʼ������"
    On Error GoTo err1

    StartNA
    connectPNAMark = True
    frmIni.Print Tab(5); "�������ֳɹ���"
    Exit Sub
err1:
    MsgBox "���ȿ������֣�"
End Sub

Private Sub Form_Activate()
    ReTryini.Enabled = False
    connectPNAMark = False
    Sureini.Enabled = False
    Cls
    Print Tab(5)
    Print Tab(5)
    Print Tab(5); "���ڳ�ʼ������"
    Print Tab(5); "���ڿ����������ƿ�"
    markX = IniBoard
    If markX = 0 Then
        Print Tab(5); "��ⲻ�����ƿ�������MPC08�˶����ƿ��Ƿ���ȷ�������ļ������"
        ReTryini.Enabled = True
        Sureini.Enabled = True
        Exit Sub
    End If
    If markX < 0 Then
        Print Tab(5); "���Ƴ�����ô�������MPC08�������Ƿ���ȷ��װ��"
        ReTryini.Enabled = True
        Sureini.Enabled = True
        Exit Sub
    End If
    Print Tab(5); "��⵽����Ŀ"; numCard; "������Ŀ"; numTotalAxes
    Print Tab(5); "���ڶ�ȡĬ��У���ļ�"
    markZ = ReadCorrect
    If markZ = -1 Then
        Print Tab(5); "Ĭ��У���ļ������ڣ�ʹ��Ĭ��У������1mm = 3200pulse"
    ElseIf markZ = 0 Then
        Print Tab(5); "��ȡĬ��У���ļ�����ʹ��Ĭ��У������1mm = 3200pulse"
        Call DefaultCorrect
    Else
        Print Tab(5); "��ȡĬ��У���ļ��ɹ���X�� "; mmMaxX; "mm = "; pMaxX; "pulse��Y�� "; mmMaxY; "mm = "; pMaxY; "pulse��"
    End If
    Print Tab(5); "���ڶ�ȡĬ���ٶ������ļ�"
    markA = ReadSpeedSet
    If markA = -1 Then
        Print Tab(5); "Ĭ���ٶ������ļ������ڣ�ʹ��Ĭ�ϲ���"
        pConSpeedX = 2000
        pConSpeedY = 2000
        pMaxSpeedX = 80000
        pMaxSpeedY = 80000
    ElseIf markA = 0 Then
        Print Tab(5); "��ȡĬ���ٶ������ļ�����ʹ��Ĭ�ϲ���"
        pConSpeedX = 2000
        pConSpeedY = 2000
        pMaxSpeedX = 80000
        pMaxSpeedY = 80000
    Else
        Print Tab(5); "��ȡĬ���ٶ������ļ��ɹ�"
        If SetBoard = 0 Then
            Print Tab(5); "�����ٶȲ����ɹ�"
        Else
            Print Tab(5); "�����ٶȲ���ʧ�ܣ�ʹ��Ĭ�ϲ���"
            pConSpeedX = 2000
            pConSpeedY = 2000
            pMaxSpeedX = 80000
            pMaxSpeedY = 80000
        End If
    End If
    Print Tab(5); "���������ݾ����"
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
        If MsgBox("�Ƿ��˳�MPC08���Ƴ���", vbYesNo, "��ʾ") = vbNo Then Cancel = True
        markCancel = 0
        Exit Sub
    End If
    If markX <= 0 Then
        If MsgBox("�Ƿ��˳�MPC08���Ƴ���", vbYesNo, "��ʾ") = vbNo Then Cancel = True
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
            Print Tab(5); "��ʼ����ɣ����� ȷ�� ����"
            Sureini.Enabled = True
            TimerIni.Enabled = False
        End If
    End If
    
End Sub
