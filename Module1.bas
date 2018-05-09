Attribute VB_Name = "Module1"
Public numTotalAxes As Long '总轴数
Public numCard As Long '总卡数
Public markX As Long '控制卡读取标签
Public markY As Long '文件读取标签
Public markZ As Long '参数读取标签
Public mmMaxY As Double 'Y轴mm校正值
Public pMaxY As Long 'Y轴Pulse校正值
Public mmMaxX As Double 'X轴mm校正值
Public pMaxX As Long 'X轴Pulse校正值
Public markCancel As Long '取消键用的标签
Public mmMaxSpeedX As Double 'x方向最大速度毫米
Public pMaxSpeedX As Double 'x方向最大速度脉冲
Public mmMaxSpeedY As Double
Public pMaxSpeedY As Double
Public mmConSpeedX As Double
Public pConSpeedX As Double
Public mmConSpeedY As Double
Public pConSpeedY As Double
Public pAddX As Long
Public mmAddX As Double
Public pAddY As Long
Public mmAddY As Double
Public pAddZ As Double  'Z轴pulse的增加
Public mmAddZ As Double 'Z轴mm的增加
Public pPosiX As Long
Public mmPosiX As Double
Public pPosiY As Long
Public mmPosiY As Double
Public pNPosiX As Long
Public mmNPosiX As Double
Public pNPosiY As Long
Public mmNPosiY As Double
Public errA As Long
Public markBack As Long
Public markMoveBack As Long
Public speedOK As Long
Public markTimer As Long
Public Declare Sub Sleep Lib "kernel32 " (ByVal dwMilliseconds As Long)
'以下为网分控制所用参数
Public na As IApplication
Public meas As IMeasurement
Public ch As IChannel
Public scpi As ScpiStringParser
Public maxMkrValue As Single
Public pnaID As String
'以上为网分控制所用参数
'以下为手动扫marker所用参数
Public connectPNAMark As Boolean
Public linMagMarker As Double
Public logMagMarker As Double
Public phaseMarker As Double
Public outPosiX As Double
Public outPosiY As Double
Public IAvg As Long
'以上为手动扫Marker所用参数
'以下为自动扫Marker所用参数
Public stepsDA As Double
Public stepsIA As Integer
Public IA As Integer
Public addXA As Double
Public addYA As Double
Public startXA As Double
Public stepXA As Double
Public endXA As Double
Public endYA As Double
Public startYA As Double
Public stepYA As Double
Public markGetData As Boolean
Public markStopScanA As Boolean
'以上为自动扫Marker所用参数
'以下为网分采集数据所用参数
Public getDataDelayTime As Integer '等待网分稳定的时间
Public getDataAvgNum As Integer '取平均用到的测量次数
Public getDataAvgDelayTime As Integer '两次测量的时间间隔
'以上为网分采集数据所用参数
'以下为手动扫谱所用参数
Public rateFreq As Long
Public specLog(8192) As Double
Public specPhase(8192) As Double
Public specLin(8192) As Double
Public specReal(8192) As Double
Public specImg(8192) As Double
Public freqss(8192) As Double
Public outPosiXS As Double
Public outPosiYS As Double
Public IAvgS As Long
Public startFreq As Double
Public stopFreq As Double
Public numOfPoints As Long
Public tmpSpecReal As Variant
Public tmpSpecImg As Variant
Public tmpSpecPhase As Variant
Public tmpSpecLog As Variant
Public tmpSpecLin As Variant
Public tmpnumOfPoints As Long
Public tmpi As Long
Public sweepType As Integer
Public freqx
'以上为手动扫谱所用参数
'以下为自动扫谱所用参数
Public stepsDSA As Double
Public stepsISA As Integer
Public ISA As Integer
Public addXSA As Double
Public addYSA As Double
Public startXSA As Double
Public stepXSA As Double
Public endXSA As Double
Public endYSA As Double
Public startYSA As Double
Public stepYSA As Double
Public markGetDataS As Boolean
Public markStopScanSA As Boolean
'以上为自动扫谱所用参数
'以下为二维扫谱所用参数
Public stepsDSB As Double
Public stepsISB As Integer
Public ISB As Integer
Public addXSB As Double
Public addYSB As Double
Public startXSB As Double
Public stepAXSB As Double
Public stepBXSB As Double
Public IASB As Integer
Public IBSB As Integer
Public IB As Integer
Public startYSB As Double
Public stepAYSB As Double
Public stepBYSB As Double
Public markGetDataSB As Boolean
Public markStopScanSB As Boolean
'以上为二维扫谱所用参数



Public Function ReadCorrect() As Long
    If Len(dir(App.Path + "\correctfiles\default.txt")) <= 0 Then
        Call DefaultCorrect
        ReadCorrect = -1
        Exit Function
    End If
    Open App.Path + "\correctfiles\default.txt" For Input As #1
    If Not EOF(1) Then
        Input #1, pMaxX
    Else
        Call DefaultCorrect
        ReadCorrect = 0
        Close #1
        Exit Function
    End If
    If Not EOF(1) Then
        Input #1, mmMaxX
    Else
        Call DefaultCorrect
        ReadCorrect = 0
        Close #1
        Exit Function
    End If
    If Not EOF(1) Then
        Input #1, pMaxY
    Else
        Call DefaultCorrect
        ReadCorrect = 0
        Close #1
        Exit Function
    End If
    If Not EOF(1) Then
        Input #1, mmMaxY
    Else
        Call DefaultCorrect
        ReadCorrect = 0
        Close #1
        Exit Function
    End If
    Close #1
    ReadCorrect = 1
End Function
Public Sub DefaultCorrect()
        mmMaxY = 100
        pMaxY = 320000
        mmMaxX = 100
        pMaxX = 320000
End Sub
Public Function ReadSpeedSet() As Long
    Dim pmx As Double
    Dim pmy As Double
    Dim pcx As Double
    Dim pcy As Double
    If Len(dir(App.Path + "\speedsets\default.txt")) <= 0 Then
        ReadSpeedSet = -1
        Exit Function
    End If
    Open App.Path + "\speedsets\default.txt" For Input As #1
    If Not EOF(1) Then
        Input #1, pmx
    Else
        Call DefaultCorrect
        ReadSpeedSet = 0
        Close #1
        Exit Function
    End If
    If Not EOF(1) Then
        Input #1, pmy
    Else
        Call DefaultCorrect
        ReadSpeedSet = 0
        Close #1
        Exit Function
    End If
    If Not EOF(1) Then
        Input #1, pcx
    Else
        Call DefaultCorrect
        ReadSpeedSet = 0
        Close #1
        Exit Function
    End If
    If Not EOF(1) Then
        Input #1, pcy
    Else
        Call DefaultCorrect
        ReadSpeedSet = 0
        Close #1
        Exit Function
    End If
    Close #1
    pMaxSpeedX = pmx
    pMaxSpeedY = pmy
    pConSpeedX = pcx
    pConSpeedY = pcy
    set_maxspeed 1, pMaxSpeedX
    set_maxspeed 2, pMaxSpeedY
    set_conspeed 1, pConSpeedX
    set_conspeed 2, pConSpeedY
    ReadSpeedSet = 1
End Function
Public Function SetBoard() As Long
    If set_maxspeed(1, pMaxSpeedX) = -1 Then
        SetBoard = -1
        Exit Function
    End If
    If set_conspeed(1, pConSpeedX) = -1 Then
        SetBoard = -1
        Exit Function
    End If
    If set_maxspeed(2, pMaxSpeedY) = -1 Then
        SetBoard = -1
        Exit Function
    End If
    If set_conspeed(2, pConSpeedY) = -1 Then
        SetBoard = -1
        Exit Function
    End If
    SetBoard = 0
End Function
Public Function P2mmX(ByVal X As Double) As Double
    P2mmX = X * mmMaxX / pMaxX
End Function
Public Function Mm2pX(ByVal X As Double) As Double
    Mm2pX = X * pMaxX / mmMaxX
End Function
Public Function P2mmY(ByVal X As Double) As Double
    P2mmY = X * mmMaxY / pMaxY
End Function
Public Function Mm2pY(ByVal X As Double) As Double
    Mm2pY = X * pMaxY / mmMaxY
End Function
Public Function P2mmZ(ByVal X As Double) As Double
    P2mmZ = X * mmMaxZ / pMaxZ
End Function
Public Function Mm2pZ(ByVal X As Double) As Double
    Mm2pZ = X * pMaxZ / mmMaxZ
End Function

Public Function IsInt(ByVal X As String) As Boolean
    Dim Y As Double
    If IsNumeric(X) = False Then
        IsInt = False
        Exit Function
    End If
    Y = Val(X)
    If Y - Int(Y) <> 0 Then
        IsInt = False
        Exit Function
    End If
    IsInt = True
    
End Function

Public Sub MoveX(ByVal pormm As Long)
    Dim markmx As Long
    If pormm = 2 Then                   '这句话什么鬼，有等于二的可能吗/
        mmPosiX = mmPosiX + mmAddX
        pAddX = Int(Mm2pX(mmPosiX)) - pPosiX
    End If
    If pAddX >= 0 Then
        markmx = con_pmove(1, pAddX)    '步径内置函数，一个轴以常速做点位运动
    Else
        markmx = con_pmove(1, pAddX - errA)
    End If
End Sub
Public Sub MoveY(ByVal pormm As Long)
    Dim markmy As Long
    If pormm = 2 Then
        mmPosiY = mmPosiY + mmAddY
        pAddY = Int(Mm2pY(mmPosiY)) - pPosiY
    End If
    If pAddY >= 0 Then
        markmy = con_pmove(2, pAddY)
    Else
        markmy = con_pmove(2, pAddY - errA)
    End If
End Sub
Public Sub MoveZ(ByVal pormm As Long)
   Dim markmz As Long
    If pormm = 2 Then
        mmPosiZ = mmPosiZ + mmAddZ
        pAddZ = Int(Mm2pZ(mmPosiZ)) - pPosiZ
    End If
    If pAddZ >= 0 Then
        markmz = con_pmove(3, pAddY)
    Else
        markmz = con_pmove(3, pAddY - errA)
    End If
End Sub
Public Sub moveBack()
    If pAddX < 0 Then
        con_pmove 1, errA
        pAddX = 0
    End If
    If pAddY < 0 Then
        con_pmove 2, errA
        pAddY = 0
    End If
    markMoveBack = 0
End Sub

Public Sub StartNA() '网分初始化
 ' Connects to the PNA application, presets, and defines some parameters
 
 Set na = CreateObject("AgilentPNA835X.Application", "192.168.0.2")
 ' Set na = CreateObject("AgilentPNA835X.Application", "xxxxx") ' Use this method to select a specific destination
 ' The above line needs the destination's "Computer Name" or IP Address used in place of "xxxxx"
 ' na.Preset
 Set scpi = na.ScpiStringParser
 Set ch = na.ActiveChannel
 Set meas = na.ActiveMeasurement
End Sub
Public Sub ScanMoveA()

End Sub
Public Sub ScanReakA()
    
End Sub
Public Function All2p(ByVal Index As Integer, ByVal xy As Integer, ByVal a As Double) As Long
    If xy = 0 Then
    If Index = 2 Then
            All2p = Int(a)
        ElseIf Index = 1 Then
            All2p = Int(Mm2pX(10 * a))
        Else
            All2p = Int(Mm2pX(a))
        End If
    End If
    If xy = 1 Then
        If Index = 2 Then
            All2p = Int(a)
        ElseIf Index = 1 Then
            All2p = Int(Mm2pY(10 * a))
        Else
            All2p = Int(Mm2pY(a))
        End If
    End If
    '''''''''''''''''''''''''
        If xy = 2 Then
        If Index = 2 Then
            All2p = Int(a)
        ElseIf Index = 1 Then
            All2p = Int(Mm2pZ(10 * a))
        Else
            All2p = Int(Mm2pZ(a))
        End If
    End If
End Function
