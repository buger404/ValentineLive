VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StartupPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Step As Single, LTime As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear 0 '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim pro As Single, tpro As Single
    
    pro = (GetTickCount - LTime) / IIf(Step = 11, 1000, 200)
    tpro = 1
    If Step = 11 Then tpro = 1 - pro
    If pro > 1 Then
        If Step = 11 Then ECore.NewTransform transFadeIn, 1000, "GamePage"
        If Step = 10 Then SE.Play "battle.wav"
        If Step <= 6 Then SE.Play "switch.wav"
        If Step < 11 Then LTime = GetTickCount: pro = 1: Step = Step + 0.5
    End If
    
    If Step > 1 Then Page.DrawImage "h2.png", GW / 2 - 100 * tpro, GH / 2 - 50 * tpro, alpha:=IIf(Step < 2, IIf(GetTickCount Mod 200 < 100, 1, 0), 1), pos:=posOnCenter
    If Step > 2 Then Page.DrawImage "h3.png", GW / 2 + 100 * tpro, GH / 2 - 50 * tpro, alpha:=IIf(Step < 3, IIf(GetTickCount Mod 200 < 100, 1, 0), 1), pos:=posOnCenter
    If Step > 3 Then Page.DrawImage "h4.png", GW / 2 - 100 * tpro, GH / 2, alpha:=IIf(Step < 4, IIf(GetTickCount Mod 200 < 100, 1, 0), 1), pos:=posOnCenter
    If Step > 4 Then Page.DrawImage "h5.png", GW / 2 + 100 * tpro, GH / 2, alpha:=IIf(Step < 5, IIf(GetTickCount Mod 200 < 100, 1, 0), 1), pos:=posOnCenter
    If Step > 5 Then Page.DrawImage "h6.png", GW / 2 - 100 * tpro, GH / 2 + 50 * tpro, alpha:=IIf(Step < 6, IIf(GetTickCount Mod 200 < 100, 1, 0), 1), pos:=posOnCenter
    If Step > 6 Then Page.DrawImage "h7.png", GW / 2 + 100 * tpro, GH / 2 + 50 * tpro, alpha:=IIf(Step < 7, IIf(GetTickCount Mod 200 < 100, 1, 0), 1), pos:=posOnCenter
    If Step > 0 Then Page.DrawImage "h1.png", GW / 2, GH / 2 - 100 * tpro, alpha:=IIf(Step < 1, IIf(GetTickCount Mod 200 < 100, 1, 0), 1), pos:=posOnCenter

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
    LTime = GetTickCount
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
    Page.Res.NewImages App.path & "\assets", 0.35
    '创建页面
    ECore.Add Page, "StartupPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
