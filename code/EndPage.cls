VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Words() As String
    Dim ETime As Long
    Dim mode As Integer
    
    Private Type PosD
        X As Single
        y As Single
        time As Long
    End Type
    Private Type Writer
        p() As PosD
        StartTime As Long
    End Type
    Dim pMark(4000) As Boolean
    Dim w As Writer, WTime As Long
    Dim HW As ScrollArea
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    If mode = 1 Then
        Page.Clear argb(255, 0, 0, 0)
        StartScrollArea Page, HW
            For Win = 1 To UBound(w.p)
                If GetTickCount - WTime >= w.p(Win).time - w.StartTime And pMark(Win) = False Then
                    Page.PaintLine w.p(Win - 1).X / 12 / 1440 * (Screen.Width / Screen.TwipsPerPixelX), w.p(Win - 1).y / 12 / 900 * (Screen.Height / Screen.TwipsPerPixelY), w.p(Win).X / 12 / 1440 * (Screen.Width / Screen.TwipsPerPixelX), w.p(Win).y / 12 / 900 * (Screen.Height / Screen.TwipsPerPixelY), argb(255, 255, 255, 255), 10
                    pMark(Win) = True
                End If
            Next
        EndScrollArea Page, 0, 0, 0, 0, GW, GH, 1
        Exit Sub
    End If
    
    Dim pro As Single, DY As Single, y As Single, pro2 As Single
    pro = (GetTickCount - ETime) / 60000
    If pro > 1 Then pro = 1
    DY = pro * (UBound(Words) * 50 + GH / 2)
    
    pro2 = (GetTickCount - 60000 - ETime) / 3000
    If pro2 < 0 Then pro2 = 0
    If pro2 > 1 Then pro2 = 1
    pro2 = Cubic(pro2, 0, 1, 1, 1)
    
    If GetTickCount - 60000 - ETime >= 3000 Then
        mode = 1: WTime = GetTickCount + 2000: ECore.NewTransform transFadeIn, 1000
    End If
    
    Dim h As Long
    h = Page.Res.ImgSize("back.png", imgGetHeight)
    
    Page.DrawImage "back.png", 0, 0, 0, pro * (h - GH), GW, GH, 1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100 + 150 * pro2, 0, 0, 0)
    
    For I = 1 To UBound(Words)
        y = I * 50 + GH - DY
        If y > 0 And y < GH Then
            pro = 1 - Cubic(Abs(y - GH / 2) / (GH / 2), 0, 1, 1, 1) ^ 2
            If pro > 1 Then pro = 1
            If I <> UBound(Words) And pro2 > 0 Then pro = pro * (1 - pro2)
            Page.Writes Words(I), GW / 2, y, 24 + IIf(I = UBound(Words), 24 * pro2, 0), IIf(Words(I) = "DETERMINATION", argb(255 * pro, 255, 0, 0), argb(255 * pro, 255, 255, 255)), align:=StringAlignmentCenter, style:=FontStyleBold
        End If
    Next
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    WTime = GetTickCount
    BGM.Dispose
    Set BGM = New GMusic
    BGM.Create App.path & "\music\ending.mp3"
    BGM.rate = 1
    BGM.Volume = 0.5
    BGM.Play
    ETime = GetTickCount
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.newImage App.path & "\assets\end\back.png", GW, GW / 2000 * 5685
    ReDim Words(0)
    Dim strs As String
    Open App.path & "\note\ending.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, strs
        If InStr(strs, "//") = 0 Then ReDim Preserve Words(UBound(Words) + 1): Words(UBound(Words)) = strs
    Loop
    Close #1
    Open App.path & "\note\handwrite.hw" For Binary As #1
    Get #1, , w
    Close #1
    For I = 1 To UBound(w.p)
        w.p(I).time = w.StartTime + (w.p(I).time - w.StartTime) * 0.3
    Next
    'w.StartTime = w.p(1).Time
    HW = CreateScrollArea(GW, GH)
    '创建页面
    ECore.Add Page, "EndPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
